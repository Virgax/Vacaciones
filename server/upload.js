// Decodifica el cuerpo de una subida de Excel. Power Automate, segun como se configure la accion HTTP,
// manda el archivo en crudo, envuelto en JSON {"$content-type":..., "$content":"<base64>"} o como base64 a secas.
const isWorkbook = (b) => b.length > 4 && (b[0] === 0x50 && b[1] === 0x4b /* PK (xlsx/xlsb) */ || b[0] === 0xd0 && b[1] === 0xcf /* OLE (xls) */);

export function decodeUploadBody(body) {
  let buffer = Buffer.isBuffer(body) ? body : Buffer.from(String(body ?? ''), 'utf8');
  if (isWorkbook(buffer)) return buffer;
  const text = buffer.toString('latin1').trim();
  if (text.startsWith('{')) {
    try {
      const j = JSON.parse(buffer.toString('utf8'));
      const b64 = j.$content ?? j.content ?? j.body?.$content ?? j.body?.content;
      if (typeof b64 === 'string') {
        const out = Buffer.from(b64, 'base64');
        if (isWorkbook(out)) return out;
      }
    } catch { /* no era JSON valido */ }
  }
  if (/^[A-Za-z0-9+/=\s"]+$/.test(text.slice(0, 4096))) {
    const out = Buffer.from(text.replace(/["\s]+/g, ''), 'base64');
    if (isWorkbook(out)) return out;
  }
  return buffer; // se devuelve tal cual: parseWorkbook dira que no es un Excel
}
