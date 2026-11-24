// Lectura del libro "Control_Deuda_Vacaciones" (.xlsb / .xlsx) y normalizacion de datos.
import XLSX from 'xlsx';

/** Quita acentos, pasa a minusculas y colapsa espacios: para comparar cabeceras. */
export function norm(s) {
  return String(s ?? '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

/**
 * Normaliza lo que llega del lector de codigos: "dr0333", " DR 333", "333", "0333" -> "DR0333".
 * Devuelve '' si no parece un codigo.
 */
export function normalizeCode(input) {
  let s = String(input ?? '').toUpperCase().replace(/[\s\-_.]/g, '');
  // algunos lectores anteponen/anaden caracteres de control
  s = s.replace(/[^A-Z0-9]/g, '');
  const m = s.match(/^(?:DR)?(\d{1,6})$/);
  if (!m) return s.length ? s : '';
  return 'DR' + m[1].padStart(4, '0');
}

function num(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const n = Number(String(v).replace(/[^\d.,-]/g, '').replace(/\.(?=.*\.)/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function findHeaderRow(rows, required) {
  for (let i = 0; i < Math.min(rows.length, 30); i++) {
    const cells = rows[i].map(norm);
    if (required.every((req) => cells.some((c) => c.includes(req)))) return i;
  }
  return -1;
}

function columnMap(headerRow, spec) {
  const cells = headerRow.map(norm);
  const map = {};
  for (const [key, needles] of Object.entries(spec)) {
    const idx = cells.findIndex((c) => needles.some((n) => c.includes(n)));
    if (idx >= 0) map[key] = idx;
  }
  return map;
}

function pickSheet(wb, preferred, required) {
  const byName = wb.SheetNames.find((n) => norm(n) === norm(preferred));
  const candidates = byName ? [byName, ...wb.SheetNames.filter((n) => n !== byName)] : wb.SheetNames;
  for (const name of candidates) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true });
    const h = findHeaderRow(rows, required);
    if (h >= 0) return { name, rows, headerIdx: h };
  }
  return null;
}

const CONTROL_COLS = {
  codigo: ['codigo'],
  nombre: ['nombre'],
  tipo: ['tipo de personal'],
  departamento: ['departamento'],
  diasAdeudados: ['dias adeudados'],
  salarioDiario: ['salario diario'],
  deudaInicial: ['deuda inicial'],
  horasExtras: ['horas extras aplicadas'],
  diasCompensados: ['dias compensados'],
  diasPendientes: ['dias pendientes'],
  deudaRestante: ['deuda restante'],
  estado: ['estado'],
};

const MOV_COLS = {
  fecha: ['fecha'],
  codigo: ['codigo'],
  horas: ['horas extras trabajadas', 'horas'],
  tipoHora: ['tipo de hora'],
  monto: ['monto aplicado'],
  observacion: ['observacion'],
};

function toISODate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.toISOString().slice(0, 10);
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}

const round2 = (n) => (n == null ? null : Math.round(n * 100) / 100);
const round3 = (n) => (n == null ? null : Math.round(n * 1000) / 1000);

/**
 * Parsea el libro y devuelve { empleados: Map<codigo, registro>, movimientos: Map<codigo, []>, info }.
 * Los valores calculados por Excel (deuda inicial, dias pendientes, deuda restante) se usan tal cual;
 * si una celda viene vacia se recalculan con las mismas formulas del libro.
 */
export function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const control = pickSheet(wb, 'Control', ['codigo', 'dias adeudados', 'deuda inicial']);
  if (!control) throw new Error('No se encontro la hoja "Control" (cabeceras Codigo / Dias adeudados / Deuda inicial)');
  const cols = columnMap(control.rows[control.headerIdx], CONTROL_COLS);
  for (const k of ['codigo', 'nombre', 'diasAdeudados']) {
    if (cols[k] == null) throw new Error(`Falta la columna "${k}" en la hoja ${control.name}`);
  }

  const empleados = new Map();
  for (const row of control.rows.slice(control.headerIdx + 1)) {
    const code = normalizeCode(row[cols.codigo]);
    if (!/^DR\d{4,}$/.test(code)) continue; // filas de totales / vacias
    const get = (k) => (cols[k] == null ? null : row[cols[k]]);

    const diasAdeudados = Math.abs(num(get('diasAdeudados')) ?? 0);
    const salarioDiario = num(get('salarioDiario'));
    let deudaInicial = num(get('deudaInicial'));
    if (deudaInicial == null) deudaInicial = salarioDiario != null ? diasAdeudados * salarioDiario : 0;
    const horasExtras = num(get('horasExtras')) ?? 0;
    const diasCompensados = num(get('diasCompensados')) ?? 0;
    let diasPendientes = num(get('diasPendientes'));
    if (diasPendientes == null) diasPendientes = Math.max(diasAdeudados - diasCompensados, 0);
    let deudaRestante = num(get('deudaRestante'));
    if (deudaRestante == null) deudaRestante = Math.max(deudaInicial - horasExtras, 0);

    const diasPagados = Math.max(diasAdeudados - diasPendientes, 0);
    const deudaPagada = Math.max(deudaInicial - deudaRestante, 0);
    let estado = String(get('estado') ?? '').trim();
    if (!estado) estado = deudaRestante <= 0 ? 'Saldado' : 'Pendiente';

    // codigo repetido en el libro: se queda la fila que aun tiene deuda (o la primera)
    const prev = empleados.get(code);
    if (prev && !(prev.deudaRestante <= 0 && deudaRestante > 0)) continue;
    empleados.set(code, {
      codigo: code,
      nombre: String(get('nombre') ?? '').replace(/\s+/g, ' ').trim(),
      tipo: get('tipo') ? String(get('tipo')).trim() : null,
      departamento: get('departamento') ? String(get('departamento')).trim() : null,
      diasAdeudados: round3(diasAdeudados),
      deudaInicial: round2(deudaInicial),
      diasPendientes: round3(diasPendientes),
      deudaRestante: round2(deudaRestante),
      diasPagados: round3(diasPagados),
      deudaPagada: round2(deudaPagada),
      porcentajePagado: deudaInicial > 0 ? Math.min(100, Math.round((deudaPagada / deudaInicial) * 1000) / 10) : 100,
      estado,
    });
  }

  const movimientos = new Map();
  const mov = pickSheet(wb, 'Movimientos', ['fecha', 'codigo', 'monto aplicado']);
  if (mov) {
    const mc = columnMap(mov.rows[mov.headerIdx], MOV_COLS);
    for (const row of mov.rows.slice(mov.headerIdx + 1)) {
      const code = normalizeCode(row[mc.codigo]);
      if (!/^DR\d{4,}$/.test(code)) continue;
      const monto = num(mc.monto == null ? null : row[mc.monto]);
      const horas = num(mc.horas == null ? null : row[mc.horas]);
      if (monto == null && horas == null) continue;
      if (!movimientos.has(code)) movimientos.set(code, []);
      movimientos.get(code).push({
        fecha: toISODate(mc.fecha == null ? null : row[mc.fecha]),
        horas: round2(horas),
        tipoHora: mc.tipoHora == null || row[mc.tipoHora] == null ? null : String(row[mc.tipoHora]).trim(),
        monto: round2(monto),
        observacion: mc.observacion == null || row[mc.observacion] == null ? null : String(row[mc.observacion]).trim(),
      });
    }
    for (const list of movimientos.values()) list.sort((a, b) => String(a.fecha ?? '').localeCompare(String(b.fecha ?? '')));
  }

  const info = {
    hojaControl: control.name,
    hojaMovimientos: mov?.name ?? null,
    empleados: empleados.size,
    conDeuda: [...empleados.values()].filter((e) => e.deudaRestante > 0).length,
    movimientos: [...movimientos.values()].reduce((a, l) => a + l.length, 0),
  };
  return { empleados, movimientos, info };
}
