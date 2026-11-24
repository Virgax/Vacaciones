import express from 'express';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DataSource } from './source.js';
import { GraphClient, graphConfigFromEnv } from './graph.js';
import { normalizeCode } from './excel.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const ADMIN_KEY = process.env.ADMIN_KEY || '';
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, '..', 'data');

const graphCfg = graphConfigFromEnv();
const source = new DataSource({
  dataDir: DATA_DIR,
  localFile: process.env.DATA_FILE || '',
  graph: graphCfg ? new GraphClient(graphCfg) : null,
  minCheckMs: Number(process.env.MIN_CHECK_SECONDS || 15) * 1000,
  refreshMs: Number(process.env.REFRESH_MINUTES || 0) * 60_000,
});

if (!ADMIN_KEY) console.warn('[aviso] ADMIN_KEY vacia: las acciones de administracion quedan abiertas');
if (!graphCfg) console.warn('[aviso] SharePoint no configurado (GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET / SP_FILE_URL): se usara DATA_FILE o el Excel subido');

const app = express();
app.disable('x-powered-by');
app.set('etag', false);
app.use('/api', (req, res, next) => { res.set('Cache-Control', 'no-store'); next(); });

// limite sencillo por IP para que nadie "barra" codigos
const hits = new Map();
function rateLimit(max, windowMs) {
  return (req, res, next) => {
    const ip = req.ip || req.socket.remoteAddress || '?';
    const now = Date.now();
    const h = hits.get(ip) || { n: 0, t: now };
    if (now - h.t > windowMs) { h.n = 0; h.t = now; }
    h.n++; hits.set(ip, h);
    if (hits.size > 5000) hits.clear();
    if (h.n > max) return res.status(429).json({ error: 'Demasiadas consultas, espera un momento' });
    next();
  };
}

function requireAdmin(req, res, next) {
  if (!ADMIN_KEY) return next();
  const k = req.get('x-admin-key') || (req.get('authorization') || '').replace(/^bearer\s+/i, '') || req.query.key;
  if (k === ADMIN_KEY) return next();
  res.status(401).json({ error: 'clave de administracion incorrecta' });
}

app.get('/health', (req, res) => res.json({ ok: !!source.data, ...source.status() }));

app.get('/api/status', (req, res) => {
  const s = source.status();
  // no exponer detalles internos del error a la pantalla publica
  res.json({ ...s, lastError: s.lastError ? s.lastError.slice(0, 200) : null });
});

app.get('/api/empleado/:codigo', rateLimit(40, 60_000), async (req, res) => {
  const code = normalizeCode(req.params.codigo);
  if (!/^DR\d{4,}$/.test(code)) return res.status(400).json({ error: 'Codigo no valido', codigo: code });
  const emp = await source.lookup(code);
  const s = source.status();
  const datos = { via: s.via, fileModified: s.fileModified, online: s.online, loadedAt: s.loadedAt };
  if (!emp) return res.status(404).json({ error: 'Codigo no encontrado en el control de deuda', codigo: code, datos });
  res.json({ ...emp, datos });
});

// ---------- administracion ----------
app.post('/api/admin/refresh', requireAdmin, async (req, res) => {
  const changed = await source.checkLive(true);
  res.json({ ok: true, changed, ...source.status() });
});
app.post('/api/admin/upload', requireAdmin, express.raw({ type: () => true, limit: '25mb' }), (req, res) => {
  if (!Buffer.isBuffer(req.body) || req.body.length < 100) return res.status(400).json({ error: 'Envia el archivo .xlsb/.xlsx en el cuerpo de la peticion' });
  try {
    const info = source.uploadBuffer(req.body, req.get('x-file-name') || 'subido.xlsb');
    res.json({ ok: true, info, ...source.status() });
  } catch (e) {
    res.status(400).json({ error: `No se pudo leer el Excel: ${e.message}` });
  }
});
app.get('/api/admin/status', requireAdmin, (req, res) => res.json(source.status()));

app.use(express.static(path.join(__dirname, '..', 'public'), { etag: true, maxAge: '1h', index: 'index.html' }));
app.use((req, res) => res.status(404).json({ error: 'not found' }));

source.init()
  .then((s) => console.log(`[datos] listo: ${s.empleados} empleados via ${s.via}; Excel modificado ${s.fileModified}`))
  .catch((e) => console.error(`[datos] ${e.message}`))
  .finally(() => {
    app.listen(PORT, () => console.log(`Consulta de deuda de vacaciones escuchando en http://localhost:${PORT}`));
  });
