import { test } from 'node:test';
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import { parseWorkbook, normalizeCode } from '../server/excel.js';
import { fakeWorkbook } from './helpers.js';

test('normalizeCode acepta las variantes del lector', () => {
  assert.equal(normalizeCode('DR0333'), 'DR0333');
  assert.equal(normalizeCode('dr333'), 'DR0333');
  assert.equal(normalizeCode(' 333 '), 'DR0333');
  assert.equal(normalizeCode(333), 'DR0333');
  assert.equal(normalizeCode('DR-0333\r'), 'DR0333');
  assert.equal(normalizeCode('DR12345'), 'DR12345');
  assert.equal(normalizeCode(''), '');
  assert.doesNotMatch(normalizeCode('Código'), /^DR\d{4}$/);
});

test('parsea Control y Movimientos y calcula lo pagado', () => {
  const { empleados, movimientos, info } = parseWorkbook(fakeWorkbook());
  assert.equal(info.empleados, 4);
  assert.equal(info.conDeuda, 3);
  assert.equal(empleados.get('DR0004').nombre, 'DUPLICADO CON DEUDA');
  const a = empleados.get('DR0001');
  assert.equal(a.nombre, 'ANA PRUEBA');
  assert.equal(a.diasAdeudados, 10);
  assert.equal(a.deudaInicial, 10000);
  assert.equal(a.diasPendientes, 7.5);
  assert.equal(a.deudaRestante, 7500);
  assert.equal(a.diasPagados, 2.5);
  assert.equal(a.deudaPagada, 2500);
  assert.equal(a.porcentajePagado, 25);
  assert.equal(a.estado, 'Pendiente');
  assert.equal(movimientos.get('DR0001').length, 1);
  assert.equal(movimientos.get('DR0001')[0].fecha, '2026-11-01');
  assert.equal(movimientos.get('DR0001')[0].monto, 2500);

  const l = empleados.get('DR0002');
  assert.equal(l.deudaPagada, 0); assert.equal(l.porcentajePagado, 100); assert.equal(l.estado, 'Saldado');

  // celdas de formula vacias -> se recalculan como en el libro
  const s = empleados.get('DR0003');
  assert.equal(s.deudaInicial, 2000);     // 4 dias * 500
  assert.equal(s.diasPendientes, 3);      // 4 - 1 compensado
  assert.equal(s.deudaRestante, 1500);    // 2000 - 500 aplicados
  assert.equal(s.deudaPagada, 500);
  assert.equal(s.estado, 'Pendiente');
  assert.equal(movimientos.get('DR0003')[0].observacion, 'nota');
});

test('rechaza un libro sin la hoja de control', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['a', 'b'], [1, 2]]), 'Hoja1');
  assert.throws(() => parseWorkbook(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })), /Control/);
});
