import XLSX from 'xlsx';

export function fakeWorkbook() {
  const control = [
    [null, 'CONTROL DE DÍAS DE VACACIONES A COMPENSAR EN DINERO'],
    [],
    [null, 'Código', 'Nombre', 'Tipo de personal', 'Departamento', 'Días adeudados iniciales', 'Cantidad de horas normales', 'Salario mensual (RD$)', 'Salario diario (RD$)', 'Salario por hora  (RD$)', 'Deuda inicial (RD$)', 'Horas extras aplicadas (RD$)', 'Días compensados', 'Días pendientes', 'Deuda restante (RD$)', 'Estado'],
    [1, 'DR0001', 'ANA PRUEBA', 'Directo', 'Calidad', -10, 88, 30000, 1000, 125, 10000, 2500, 2.5, 7.5, 7500, 'Pendiente'],
    [2, 'DR0002', 'LUIS SALDADO', 'Indirecto', 'Software', 0, 0, 50000, 2000, 250, 0, 0, 0, 0, 0, 'Saldado'],
    [3, 'DR0003', 'SIN CALCULO', 'Directo', null, -4, 35.2, 20000, 500, 62.5, null, 500, 1, null, null, null],
    [4, 'DR0004', 'DUPLICADO SALDADO', 'Directo', null, 0, 0, 20000, 500, 62.5, 0, 0, 0, 0, 0, 'Saldado'],
    [5, 'DR0004', 'DUPLICADO CON DEUDA', 'Directo', null, -2, 17.6, 20000, 500, 62.5, 1000, 0, 0, 2, 1000, 'Pendiente'],
    [null, null, null, null, null, -14, 123.2, null, null, null, 10000, 3000, 3.5, 7.5, 7500, null], // totales
  ];
  const mov = [
    ['REGISTRO DE HORAS EXTRAS'],
    ['instrucciones'],
    [],
    ['Fecha', 'Código', 'Nombre', 'Horas extras trabajadas', 'Tipo de hora', 'Monto aplicado a la deuda (RD$) – calculado', 'Observación'],
    [new Date(Date.UTC(2026, 10, 1)), 'DR0001', 'ANA PRUEBA', 10, '100%', 2500, null],
    [new Date(Date.UTC(2026, 9, 31)), 'DR0003', 'SIN CALCULO', 4, '100%', 500, 'nota'],
    [null, null, '', null, null, '', null],
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(control, { cellDates: true }), 'Control');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mov, { cellDates: true }), 'Movimientos');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

