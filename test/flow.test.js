import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DataSource } from '../server/source.js';

const quiet = { log() {}, warn() {} };
const fila = (codigo, nombre, dias, pend, deudaIni, deudaRest) => ({ '@odata.etag': '', ItemInternalId: 'x', Column1: '1', 'Código': codigo, Nombre: nombre, 'Días adeudados iniciales': String(dias), 'Deuda inicial (RD$)': String(deudaIni), 'Días pendientes': String(pend), 'Deuda restante (RD$)': String(deudaRest), Estado: deudaRest > 0 ? 'Pendiente' : 'Saldado' });

function fakeFlow() {
  const f = { online: true, rows: [fila('DR0001', 'ANA', -10, 7.5, 10000, 7500), fila('DR0002', 'LUIS', 0, 0, 0, 0)], calls: 0 };
  f.fetch = async (url, init) => {
    f.calls++;
    if (!f.online) throw new Error('fetch failed');
    assert.equal(init.method, 'POST');
    return { ok: true, status: 200, text: async () => JSON.stringify(f.rows) };
  };
  return f;
}

test('la app llama al flujo, guarda copia, detecta cambios y aguanta sin conexion', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dv-'));
  const f = fakeFlow();
  const ds = new DataSource({ dataDir: dir, flowUrl: 'https://flujo.example/invoke', fetchImpl: f.fetch, minCheckMs: 0, log: quiet });
  await ds.init();
  assert.equal(ds.state.online, true);
  assert.equal(ds.state.via, 'power automate (flujo)');
  assert.equal(f.calls, 1);
  const primeraFecha = ds.state.fileModified;
  assert.ok(fs.existsSync(ds.rowsPath));

  // sin cambios en la tabla: se llama al flujo pero no cambia la fecha
  await ds.lookup('DR0001');
  assert.equal(f.calls, 2);
  assert.equal(ds.state.fileModified, primeraFecha);

  // RRHH aplica un pago
  f.rows[0] = fila('DR0001', 'ANA', -10, 5, 10000, 5000);
  const e = await ds.lookup('1');
  assert.equal(e.deudaPagada, 5000);
  assert.notEqual(ds.state.fileModified, primeraFecha);

  // se cae la conexion: responde con lo ultimo y marca offline
  f.online = false;
  assert.equal((await ds.lookup('DR0002')).nombre, 'LUIS');
  assert.equal(ds.state.online, false);
  assert.match(ds.state.lastError, /fetch failed/);

  // vuelve
  f.online = true;
  await ds.lookup('DR0002');
  assert.equal(ds.state.online, true);
  ds.close();

  // reinicio sin conexion: arranca con la copia del flujo
  const f2 = fakeFlow(); f2.online = false;
  const ds2 = new DataSource({ dataDir: dir, flowUrl: 'https://flujo.example/invoke', fetchImpl: f2.fetch, minCheckMs: 0, log: quiet });
  await ds2.init();
  assert.equal(ds2.state.via, 'power automate (flujo) (copia local)');
  assert.equal((await ds2.lookup('DR0001')).deudaPagada, 5000);
  ds2.close();
});

test('si el flujo devuelve algo que no son filas de Control, se conservan los datos anteriores', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dv-'));
  const f = fakeFlow();
  const ds = new DataSource({ dataDir: dir, flowUrl: 'https://flujo.example/invoke', fetchImpl: f.fetch, minCheckMs: 0, log: quiet });
  await ds.init();
  f.rows = [{ Nombre: 'sin codigo' }];
  await ds.lookup('DR0001');
  assert.equal(ds.state.online, false);
  assert.match(ds.state.lastError, /Falta la columna/);
  assert.equal((await ds.lookup('DR0001')).nombre, 'ANA');
  ds.close();
});
