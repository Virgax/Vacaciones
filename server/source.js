// Fuente de datos con "failsafe":
//  - En cada consulta se comprueba en vivo (Graph) si el Excel de SharePoint cambio; si cambio se descarga.
//  - Cada descarga se guarda en DATA_DIR como copia local.
//  - Si SharePoint/Internet falla, se responde con la ultima copia (y se avisa de la fecha de esa copia).
import fs from 'node:fs';
import path from 'node:path';
import { parseWorkbook, parseRows, normalizeCode } from './excel.js';

const CACHE_NAME = 'Control_Deuda_Vacaciones.cache.xlsb';
const META_NAME = 'Control_Deuda_Vacaciones.cache.json';
const ROWS_NAME = 'Control_Deuda_Vacaciones.rows.json';   // filas recibidas de Power Automate
const ROWS_META = 'Control_Deuda_Vacaciones.rows.meta.json';

export class DataSource {
  /**
   * @param {object} o
   * @param {string} o.dataDir           carpeta para la copia local
   * @param {string} [o.localFile]       archivo fijo a usar si no hay SharePoint (DATA_FILE)
   * @param {import('./graph.js').GraphClient|null} [o.graph]
   * @param {number} [o.minCheckMs]      no volver a preguntar a SharePoint antes de este tiempo (por defecto 15 s)
   * @param {number} [o.refreshMs]       refresco periodico en segundo plano (0 = desactivado)
   */
  constructor({ dataDir, localFile = '', graph = null, minCheckMs = 15_000, refreshMs = 0, log = console }) {
    this.dataDir = dataDir;
    this.localFile = localFile;
    this.graph = graph;
    this.minCheckMs = minCheckMs;
    this.refreshMs = refreshMs;
    this.log = log;
    this.data = null;           // { empleados, movimientos, info }
    this.state = {
      source: graph ? 'sharepoint' : 'local',   // de donde viene la configuracion
      online: null,                              // ultimo intento a SharePoint: true/false
      loadedAt: null,                            // cuando se parseo lo que esta en memoria
      fileModified: null,                        // lastModified del Excel en SharePoint (o mtime del archivo)
      fileName: null,
      via: null,                                 // ruta / enlace / busqueda / copia local / archivo local / subido
      lastCheck: null,                           // ultima comprobacion en vivo
      lastError: null,
      empleados: 0,
    };
    this._checking = null;
    fs.mkdirSync(dataDir, { recursive: true });
  }

  get cachePath() { return path.join(this.dataDir, CACHE_NAME); }
  get metaPath() { return path.join(this.dataDir, META_NAME); }
  get rowsPath() { return path.join(this.dataDir, ROWS_NAME); }
  get rowsMetaPath() { return path.join(this.dataDir, ROWS_META); }

  // ---------- arranque ----------
  async init() {
    // 1) copia local primero, para poder responder aunque SharePoint tarde o falle
    this._loadLocalCopy();
    // 2) intento en vivo
    if (this.graph) await this.checkLive(true).catch(() => {});
    if (this.refreshMs > 0 && this.graph) {
      this._timer = setInterval(() => this.checkLive(true).catch(() => {}), this.refreshMs);
      this._timer.unref?.();
    }
    if (!this.data) throw new Error(this.state.lastError || 'No hay datos: configura SharePoint (GRAPH_*/SP_FILE_URL), DATA_FILE o sube el Excel desde /admin.html');
    return this.state;
  }

  _readMeta(file) {
    try { return fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : {}; } catch { return {}; }
  }

  _loadLocalCopy() {
    const candidates = [];
    if (this.localFile) candidates.push({ file: this.localFile, via: 'archivo local', kind: 'xlsx', savedAt: Infinity });
    const metaX = this._readMeta(this.metaPath);
    const metaR = this._readMeta(this.rowsMetaPath);
    candidates.push({ file: this.cachePath, via: 'copia local', kind: 'xlsx', meta: metaX, savedAt: Date.parse(metaX.savedAt || 0) || 0 });
    candidates.push({ file: this.rowsPath, via: 'copia local', kind: 'rows', meta: metaR, savedAt: Date.parse(metaR.savedAt || 0) || 0 });
    candidates.sort((a, b) => b.savedAt - a.savedAt); // la copia mas reciente primero
    for (const c of candidates) {
      try {
        if (!fs.existsSync(c.file)) continue;
        const buf = fs.readFileSync(c.file);
        const parsed = c.kind === 'rows' ? parseRows(buf.toString('utf8')) : parseWorkbook(buf);
        const meta = c.meta || {};
        this._apply(parsed, {
          via: meta.via && c.via === 'copia local' ? `${meta.via} (copia local)` : c.via,
          fileName: meta.fileName || path.basename(c.file),
          fileModified: meta.fileModified || fs.statSync(c.file).mtime.toISOString(),
        });
        this.log.log(`[datos] ${c.via}: ${c.file} (${parsed.info.empleados} empleados, Excel de ${this.state.fileModified})`);
        return true;
      } catch (e) {
        this.state.lastError = `${c.via}: ${e.message}`;
        this.log.warn(`[datos] no se pudo usar ${c.file}: ${e.message}`);
      }
    }
    return false;
  }

  _apply(parsed, { via, fileName, fileModified }) {
    this.data = parsed;
    this.state.loadedAt = new Date().toISOString();
    this.state.via = via;
    this.state.fileName = fileName;
    this.state.fileModified = fileModified || null;
    this.state.empleados = parsed.info.empleados;
    this.state.lastError = null;
  }

  _saveCopy(buffer, meta, { file = this.cachePath, metaFile = this.metaPath } = {}) {
    try {
      const tmp = file + '.tmp';
      fs.writeFileSync(tmp, buffer);
      fs.renameSync(tmp, file);
      fs.writeFileSync(metaFile, JSON.stringify({ ...meta, savedAt: new Date().toISOString() }, null, 2));
    } catch (e) {
      this.log.warn(`[datos] no se pudo guardar la copia local: ${e.message}`);
    }
  }

  // ---------- en vivo ----------
  /**
   * Comprueba SharePoint. Si el Excel cambio (lastModified distinto) lo descarga y lo parsea.
   * Varias consultas simultaneas comparten la misma comprobacion; y no se pregunta mas de una vez por minCheckMs.
   * Nunca lanza: deja el resultado en state.online / state.lastError.
   */
  async checkLive(force = false) {
    if (!this.graph) return false;
    if (this._checking) return this._checking;
    const last = this.state.lastCheck ? Date.parse(this.state.lastCheck) : 0;
    if (!force && Date.now() - last < this.minCheckMs) return false;
    this._checking = this._doCheck().finally(() => { this._checking = null; });
    return this._checking;
  }

  async _doCheck() {
    this.state.lastCheck = new Date().toISOString();
    try {
      const item = await this.graph.locateItem();
      // metadatos (barato) -> solo descargamos si cambio
      const meta = await (await this.graph.api(`/drives/${item.driveId}/items/${item.itemId}?$select=id,name,lastModifiedDateTime,size`)).json();
      const changed = !this.data || meta.lastModifiedDateTime !== this.state.fileModified || !String(this.state.via).startsWith('sharepoint');
      if (changed) {
        const dl = await this.graph.download();
        const parsed = parseWorkbook(dl.buffer);
        this._apply(parsed, { via: `sharepoint (${dl.via})`, fileName: dl.name, fileModified: dl.lastModified });
        this._saveCopy(dl.buffer, { fileName: dl.name, fileModified: dl.lastModified, webUrl: dl.webUrl });
        this.log.log(`[datos] SharePoint: descargado ${dl.name} (modificado ${dl.lastModified}, ${parsed.info.empleados} empleados)`);
      }
      this.state.online = true;
      this.state.lastError = null;
      return changed;
    } catch (e) {
      this.state.online = false;
      this.state.lastError = e.message;
      this.log.warn(`[datos] SharePoint no disponible: ${e.message}${this.data ? ' -> usando ' + this.state.via : ''}`);
      return false;
    }
  }

  // ---------- subida manual (admin) ----------
  uploadBuffer(buffer, fileName = 'subido.xlsb', { modified = null, via = 'subido' } = {}) {
    const parsed = parseWorkbook(buffer); // valida antes de guardar
    const d = modified ? new Date(modified) : null;
    const fileModified = d && !Number.isNaN(d.getTime()) ? d.toISOString() : new Date().toISOString();
    this._apply(parsed, { via, fileName, fileModified });
    this._saveCopy(buffer, { fileName, fileModified, via });
    return parsed.info;
  }

  /** Filas en JSON (Power Automate "Obtener filas de una tabla"). */
  uploadRows(json, { modified = null, via = 'power automate (filas)' } = {}) {
    const parsed = parseRows(json); // valida antes de guardar
    const d = modified ? new Date(modified) : null;
    const fileModified = d && !Number.isNaN(d.getTime()) ? d.toISOString() : new Date().toISOString();
    const fileName = 'filas.json';
    this._apply(parsed, { via, fileName, fileModified });
    this._saveCopy(Buffer.from(typeof json === 'string' ? json : JSON.stringify(json)), { fileName, fileModified, via }, { file: this.rowsPath, metaFile: this.rowsMetaPath });
    return parsed.info;
  }

  // ---------- consultas ----------
  /** Busca un empleado; antes intenta refrescar en vivo. */
  async lookup(code) {
    await this.checkLive(false);
    if (!this.data) return null;
    const emp = this.data.empleados.get(normalizeCode(code));
    if (!emp) return null;
    return { ...emp, movimientos: this.data.movimientos.get(code) || [] };
  }

  status() {
    return { ...this.state, info: this.data?.info ?? null, ahora: new Date().toISOString() };
  }

  close() { if (this._timer) clearInterval(this._timer); }
}
