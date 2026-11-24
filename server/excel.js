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
  if (/^\d{4,6}(\.\d+)?$/.test(s)) return toISODate(Number(s));
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}

const round2 = (n) => (n == null ? null : Math.round(n * 100) / 100);
const round3 = (n) => (n == null ? null : Math.round(n * 1000) / 1000);

/** Convierte una fila (array) + cabecera en un objeto { "texto de cabecera": valor }. */
function rowToObject(header, row) {
  const o = {};
  header.forEach((h, i) => { if (h != null && String(h).trim()) o[String(h)] = row[i] ?? null; });
  return o;
}

/** Busca en un objeto la clave cuya cabecera normalizada contiene alguna de las agujas. */
function pick(obj, needles) {
  for (const k of Object.keys(obj)) {
    const nk = norm(k);
    if (needles.some((n) => nk.includes(n))) return obj[k];
  }
  return null;
}

/**
 * Construye los datos a partir de filas como objetos (clave = texto de la cabecera), que es lo que devuelve
 * tanto la lectura del libro como la accion "Obtener filas de una tabla" de Power Automate.
 * Los valores calculados por Excel (deuda inicial, dias pendientes, deuda restante) se usan tal cual;
 * si vienen vacios se recalculan con las mismas formulas del libro.
 */
export function buildData({ control = [], movimientos: movRows = [] } = {}) {
  if (!Array.isArray(control) || !control.length) throw new Error('No hay filas de la tabla Control');
  const sample = control.find((r) => r && typeof r === 'object') || {};
  for (const [k, needles] of [['Codigo', CONTROL_COLS.codigo], ['Nombre', CONTROL_COLS.nombre], ['Dias adeudados', CONTROL_COLS.diasAdeudados]]) {
    if (!Object.keys(sample).some((h) => needles.some((n) => norm(h).includes(n)))) throw new Error(`Falta la columna "${k}" en las filas de Control`);
  }

  const empleados = new Map();
  for (const row of control) {
    if (!row || typeof row !== 'object') continue;
    const get = (k) => pick(row, CONTROL_COLS[k]);
    const code = normalizeCode(get('codigo'));
    if (!/^DR\d{4,}$/.test(code)) continue; // filas de totales / vacias

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

    // codigo repetido: se queda la fila que aun tiene deuda (o la primera)
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
  for (const row of Array.isArray(movRows) ? movRows : []) {
    if (!row || typeof row !== 'object') continue;
    const code = normalizeCode(pick(row, MOV_COLS.codigo));
    if (!/^DR\d{4,}$/.test(code)) continue;
    const monto = num(pick(row, MOV_COLS.monto));
    const horas = num(pick(row, MOV_COLS.horas));
    if (monto == null && horas == null) continue;
    if (!movimientos.has(code)) movimientos.set(code, []);
    const tipoHora = pick(row, MOV_COLS.tipoHora);
    const obs = pick(row, MOV_COLS.observacion);
    movimientos.get(code).push({
      fecha: toISODate(pick(row, MOV_COLS.fecha)),
      horas: round2(horas),
      tipoHora: tipoHora == null || tipoHora === '' ? null : String(tipoHora).trim(),
      monto: round2(monto),
      observacion: obs == null || obs === '' ? null : String(obs).trim(),
    });
  }
  for (const list of movimientos.values()) list.sort((a, b) => String(a.fecha ?? '').localeCompare(String(b.fecha ?? '')));

  const info = {
    empleados: empleados.size,
    conDeuda: [...empleados.values()].filter((e) => e.deudaRestante > 0).length,
    movimientos: [...movimientos.values()].reduce((a, l) => a + l.length, 0),
  };
  return { empleados, movimientos, info };
}

/** Lee el libro (.xlsb/.xlsx) y devuelve { empleados, movimientos, info }. */
export function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const control = pickSheet(wb, 'Control', ['codigo', 'dias adeudados', 'deuda inicial']);
  if (!control) throw new Error('No se encontro la hoja "Control" (cabeceras Codigo / Dias adeudados / Deuda inicial)');
  const header = control.rows[control.headerIdx];
  const controlRows = control.rows.slice(control.headerIdx + 1).map((r) => rowToObject(header, r));
  let movRows = [];
  const mov = pickSheet(wb, 'Movimientos', ['fecha', 'codigo', 'monto aplicado']);
  if (mov) movRows = mov.rows.slice(mov.headerIdx + 1).map((r) => rowToObject(mov.rows[mov.headerIdx], r));
  const out = buildData({ control: controlRows, movimientos: movRows });
  out.info = { hojaControl: control.name, hojaMovimientos: mov?.name ?? null, ...out.info };
  return out;
}

/**
 * Acepta el JSON que manda Power Automate con "Obtener filas de una tabla":
 *   { "control": [ {fila}, ... ], "movimientos": [ {fila}, ... ] }
 * Cada lista puede venir tambien como la salida completa de la accion ({ "value": [...] }).
 */
export function parseRows(json) {
  const unwrap = (x) => {
    if (Array.isArray(x)) return x;
    if (x && typeof x === 'object' && Array.isArray(x.value)) return x.value;
    if (typeof x === 'string') { try { return unwrap(JSON.parse(x)); } catch { return []; } }
    return [];
  };
  const body = typeof json === 'string' ? JSON.parse(json) : (json || {});
  // tambien vale pegar directamente la salida de la accion ({ "value": [...] }) o el array de filas
  const control = unwrap(body.control ?? body.Control ?? body.empleados ?? body);
  const movimientos = unwrap(body.movimientos ?? body.Movimientos ?? body.pagos);
  const clean = (rows) => rows.map((r) => {
    const o = {};
    for (const [k, v] of Object.entries(r || {})) if (!k.startsWith('@odata') && k !== 'ItemInternalId') o[k] = v;
    return o;
  });
  return buildData({ control: clean(control), movimientos: clean(movimientos) });
}
