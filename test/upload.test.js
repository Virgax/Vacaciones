import { test } from 'node:test';
import assert from 'node:assert/strict';
import { decodeUploadBody } from '../server/upload.js';
import { fakeWorkbook } from './helpers.js';

test('acepta el archivo en crudo, envuelto en JSON de Power Automate o en base64', () => {
  const wb = fakeWorkbook();
  assert.ok(decodeUploadBody(wb).equals(wb), 'crudo');
  const pa = Buffer.from(JSON.stringify({ '$content-type': 'application/octet-stream', $content: wb.toString('base64') }));
  assert.ok(decodeUploadBody(pa).equals(wb), 'json power automate');
  assert.ok(decodeUploadBody(Buffer.from(wb.toString('base64'))).equals(wb), 'base64');
  assert.ok(decodeUploadBody(Buffer.from('"' + wb.toString('base64') + '"')).equals(wb), 'base64 entre comillas');
  const basura = Buffer.from('hola mundo');
  assert.ok(decodeUploadBody(basura).equals(basura), 'lo que no es Excel se deja igual');
});
