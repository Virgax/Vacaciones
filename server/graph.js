// Descarga del Excel desde SharePoint mediante Microsoft Graph (credenciales de aplicacion).
// Permisos de aplicacion necesarios en Entra ID: Sites.Read.All (o Sites.Selected concedido al sitio).

export function graphConfigFromEnv(env = process.env) {
  const { GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, SP_FILE_URL, SP_FILE_PATH } = env;
  if (!GRAPH_TENANT_ID || !GRAPH_CLIENT_ID || !GRAPH_CLIENT_SECRET || !SP_FILE_URL) return null;
  return { tenantId: GRAPH_TENANT_ID, clientId: GRAPH_CLIENT_ID, clientSecret: GRAPH_CLIENT_SECRET, fileUrl: SP_FILE_URL, filePath: SP_FILE_PATH || '' };
}

/** De la URL de SharePoint saca host, ruta del sitio y nombre del archivo. */
export function parseSharePointUrl(fileUrl) {
  const u = new URL(fileUrl);
  const m = u.pathname.match(/^(\/(?:sites|teams)\/[^/]+)/i);
  const sitePath = m ? m[1] : '';
  let fileName = u.searchParams.get('file') || '';
  if (!fileName) {
    const last = decodeURIComponent(u.pathname.split('/').pop() || '');
    if (/\.xls[xbm]$/i.test(last)) fileName = last;
  }
  return { host: u.host, sitePath, fileName };
}

export class GraphClient {
  constructor(cfg, { fetchImpl = fetch } = {}) {
    this.cfg = cfg;
    this.fetch = fetchImpl;
    this.token = null;
    this.tokenExp = 0;
    this.itemCache = null; // { driveId, itemId }
  }

  async getToken() {
    if (this.token && Date.now() < this.tokenExp - 60_000) return this.token;
    const body = new URLSearchParams({
      client_id: this.cfg.clientId,
      client_secret: this.cfg.clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    });
    const r = await this.fetch(`https://login.microsoftonline.com/${encodeURIComponent(this.cfg.tenantId)}/oauth2/v2.0/token`, {
      method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body,
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.access_token) throw new Error(`Graph token: ${r.status} ${j.error_description || j.error || ''}`.trim());
    this.token = j.access_token;
    this.tokenExp = Date.now() + (Number(j.expires_in) || 3600) * 1000;
    return this.token;
  }

  async api(path, init = {}) {
    const token = await this.getToken();
    const url = path.startsWith('http') ? path : `https://graph.microsoft.com/v1.0${path}`;
    const r = await this.fetch(url, { ...init, headers: { authorization: `Bearer ${token}`, ...(init.headers || {}) } });
    if (!r.ok) {
      const txt = await r.text().catch(() => '');
      const err = new Error(`Graph ${r.status} ${path}: ${txt.slice(0, 300)}`);
      err.status = r.status;
      throw err;
    }
    return r;
  }

  /** Localiza el driveItem del Excel: por ruta, por enlace compartido o buscando por nombre en el sitio. */
  async locateItem() {
    if (this.itemCache) return this.itemCache;
    const { host, sitePath, fileName } = parseSharePointUrl(this.cfg.fileUrl);
    const errors = [];

    if (this.cfg.filePath && sitePath) {
      try {
        const p = this.cfg.filePath.split('/').map(encodeURIComponent).join('/');
        const r = await this.api(`/sites/${host}:${sitePath}:/drive/root:/${p}`);
        const it = await r.json();
        return (this.itemCache = { driveId: it.parentReference.driveId, itemId: it.id, name: it.name, lastModified: it.lastModifiedDateTime, webUrl: it.webUrl, via: 'ruta' });
      } catch (e) { errors.push(`ruta: ${e.message}`); }
    }

    try {
      const enc = 'u!' + Buffer.from(this.cfg.fileUrl).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
      const r = await this.api(`/shares/${enc}/driveItem`);
      const it = await r.json();
      return (this.itemCache = { driveId: it.parentReference.driveId, itemId: it.id, name: it.name, lastModified: it.lastModifiedDateTime, webUrl: it.webUrl, via: 'enlace' });
    } catch (e) { errors.push(`enlace: ${e.message}`); }

    if (sitePath && fileName) {
      try {
        const site = await (await this.api(`/sites/${host}:${sitePath}`)).json();
        const drives = (await (await this.api(`/sites/${site.id}/drives?$select=id,name`)).json()).value || [];
        for (const d of drives) {
          const q = encodeURIComponent(fileName.replace(/'/g, "''"));
          const res = (await (await this.api(`/drives/${d.id}/root/search(q='${q}')?$select=id,name,lastModifiedDateTime,webUrl,parentReference`)).json()).value || [];
          const it = res.find((x) => x.name?.toLowerCase() === fileName.toLowerCase()) || res.find((x) => /\.xls[xbm]$/i.test(x.name || ''));
          if (it) return (this.itemCache = { driveId: d.id, itemId: it.id, name: it.name, lastModified: it.lastModifiedDateTime, webUrl: it.webUrl, via: `busqueda (${d.name})` });
        }
        errors.push(`busqueda: "${fileName}" no aparece en ${drives.length} biblioteca(s) de ${sitePath}`);
      } catch (e) { errors.push(`busqueda: ${e.message}`); }
    }
    throw new Error('No se pudo localizar el Excel en SharePoint. ' + errors.join(' | '));
  }

  /** Descarga el archivo. Devuelve { buffer, name, lastModified, webUrl, via }. */
  async download() {
    const item = await this.locateItem();
    try {
      const meta = await (await this.api(`/drives/${item.driveId}/items/${item.itemId}?$select=id,name,lastModifiedDateTime,webUrl`)).json();
      item.lastModified = meta.lastModifiedDateTime; item.name = meta.name; item.webUrl = meta.webUrl;
    } catch (e) {
      if (e.status === 404) this.itemCache = null; // se movio/renombro: la proxima vez se vuelve a buscar
      throw e;
    }
    const r = await this.api(`/drives/${item.driveId}/items/${item.itemId}/content`);
    const buffer = Buffer.from(await r.arrayBuffer());
    return { buffer, name: item.name, lastModified: item.lastModified, webUrl: item.webUrl, via: item.via };
  }
}
