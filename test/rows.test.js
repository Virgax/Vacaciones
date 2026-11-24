import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseRows } from '../server/excel.js';
import { DataSource } from '../server/source.js';

// Asi devuelve "Obtener filas de una tabla": claves = cabeceras, valores como texto, fechas como numero de serie
const controlRows = [
  { '@odata.etag': '', ItemInternalId: '1', 'Código': 'DR0001', 'Nombre': 'ANA PRUEBA', 'Tipo de personal': 'Directo', 'Departamento': 'Calidad', 'Días adeudados iniciales': '-10', 'Salario diario (RD$)': '1000', 'Deuda inicial (RD$)': '10000', 'Horas extras aplicadas (RD$)': '2500', 'Días compensados': '2.5', 'Días pendientes': '7.5', 'Deuda restante (RD$)': '7500', 'Estado': 'Pendiente' },
  { 'Código': 'DR0002', 'Nombre': 'LUIS SALDADO', 'Días adeudados iniciales': '0', 'Salario diario (RD$)': '2000', 'Deuda inicial (RD$)': '', 'Días pendientes': '', 'Deuda restante (RD$)': '', 'Estado': '' },
  { 'Código': '', 'Nombre': '', 'Días adeudados iniciales': '-10' }, // fila de totales
];
const movRows = [
  { 'Fecha': '46326', 'Código': 'DR0001', 'Nombre': 'ANA PRUEBA', 'Horas extras trabajadas': '10', 'Tipo de hora': '100%', 'Monto aplicado a la deuda (RD$) – calculado': '2500', 'Observación': '' },
  { 'Fecha': '2026-11-02T00:00:00.000Z', 'Código': 'DR0001', 'Horas extras trabajadas': '2', 'Tipo de hora': '100%', 'Monto aplicado a la deuda (RD$) – calculado': '500' },
];

test('parseRows entiende la salida de Obtener filas de una tabla', () => {
  const { empleados, movimientos, info } = parseRows({ control: { value: controlRows }, movimientos: movRows });
  assert.equal(info.empleados, 2);
  const a = empleados.get('DR0001');
  assert.equal(a.diasPagados, 2.5); assert.equal(a.deudaPagada, 2500); assert.equal(a.porcentajePagado, 25);
  const l = empleados.get('DR0002');
  assert.equal(l.estado, 'Saldado'); assert.equal(l.deudaInicial, 0);
  assert.deepEqual(movimientos.get('DR0001').map((m) => m.fecha), ['2026-10-31', '2026-11-02']);
  assert.throws(() => parseRows({ control: [{ Nombre: 'x' }] }), /Falta la columna/);
  assert.throws(() => parseRows({}), /No hay filas/);
  // solo la tabla Control, pegando la salida de la accion tal cual
  const solo = parseRows({ value: controlRows });
  assert.equal(solo.info.empleados, 2); assert.equal(solo.info.movimientos, 0);
  assert.equal(parseRows(controlRows).info.empleados, 2);
});

test('las filas se guardan en el volumen y se recuperan al reiniciar', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dv-'));
  const ds = new DataSource({ dataDir: dir, log: { log() {}, warn() {} } });
  ds.uploadRows({ control: controlRows, movimientos: movRows }, { modified: '2026-11-24T10:00:00Z' });
  assert.equal(ds.state.via, 'power automate (filas)');
  assert.equal(ds.state.fileModified, '2026-11-24T10:00:00.000Z');
  const ds2 = new DataSource({ dataDir: dir, log: { log() {}, warn() {} } });
  await ds2.init();
  assert.equal(ds2.state.via, 'power automate (filas) (copia local)');
  assert.equal((await ds2.lookup('1')).nombre, 'ANA PRUEBA');
});
