import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { DataSource } from '../server/source.js';
import { fakeWorkbook } from './helpers.js';

const quiet = { log() {}, warn() {} };

/** Graph falso: se puede poner offline y cambiar la fecha de modificacion del archivo. */
function fakeGraph(buffer) {
  const g = {
    online: true, modified: '2026-11-01T10:00:00Z', buffer, downloads: 0, metaCalls: 0,
    async locateItem() { if (!g.online) throw new Error('ENOTFOUND graph.microsoft.com'); return { driveId: 'd', itemId: 'i', via: 'ruta' }; },
    async api(p) {
      if (!g.online) throw new Error('ENOTFOUND graph.microsoft.com');
      g.metaCalls++;
      return { json: async () => ({ id: 'i', name: 'Control_Deuda_Vacaciones.xlsb', lastModifiedDateTime: g.modified }) };
    },
    async download() { g.downloads++; return { buffer: g.buffer, name: 'Control_Deuda_Vacaciones.xlsb', lastModified: g.modified, via: 'ruta' }; },
  };
  return g;
}

test('descarga en vivo, guarda copia y usa la copia cuando no hay internet', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dv-'));
  const g = fakeGraph(fakeWorkbook());
  const ds = new DataSource({ dataDir: dir, graph: g, minCheckMs: 0, log: quiet });
  await ds.init();
  assert.equal(ds.state.online, true);
  assert.equal(g.downloads, 1);
  assert.ok(fs.existsSync(ds.cachePath), 'guarda copia local');

  // misma fecha -> no vuelve a descargar, solo consulta metadatos
  const e1 = await ds.lookup('DR0001');
  assert.equal(e1.nombre, 'ANA PRUEBA');
  assert.equal(g.downloads, 1);

  // RRHH modifica el Excel -> se descarga de nuevo
  g.modified = '2026-11-02T09:00:00Z';
  await ds.lookup('DR0001');
  assert.equal(g.downloads, 2);
  assert.equal(ds.state.fileModified, '2026-11-02T09:00:00Z');

  // se cae internet -> sigue respondiendo con lo que hay y marca offline
  g.online = false;
  const e2 = await ds.lookup('DR0002');
  assert.equal(e2.nombre, 'LUIS SALDADO');
  assert.equal(ds.state.online, false);
  assert.match(ds.state.lastError, /ENOTFOUND/);

  // vuelve internet -> la siguiente consulta vuelve a leer en vivo
  g.online = true;
  await ds.lookup('DR0002');
  assert.equal(ds.state.online, true);
  ds.close();

  // reinicio sin internet: arranca con la copia guardada
  const g2 = fakeGraph(fakeWorkbook()); g2.online = false;
  const ds2 = new DataSource({ dataDir: dir, graph: g2, minCheckMs: 0, log: quiet });
  await ds2.init();
  assert.equal(ds2.state.via, 'copia local');
  assert.equal(ds2.state.fileModified, '2026-11-02T09:00:00Z');
  assert.equal((await ds2.lookup('DR0001')).codigo, 'DR0001');
  assert.equal(await ds2.lookup('DR9999'), null);
  ds2.close();
});

test('sin SharePoint ni archivo falla con un mensaje claro; con subida manual funciona', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dv-'));
  const ds = new DataSource({ dataDir: dir, log: quiet });
  await assert.rejects(ds.init(), /No hay datos/);
  const info = ds.uploadBuffer(fakeWorkbook(), 'manual.xlsx');
  assert.equal(info.empleados, 4);
  assert.equal(ds.state.via, 'subido');
  assert.equal((await ds.lookup('333')), null);
});
