/**
 * Sample page -> Word export pipeline.
 * Confirms one title, an embedded image, a real table, and cell borders.
 *
 * Usage: pnpm exec tsx scripts/verify-docx-export.ts
 */
import { crc32, deflateSync } from 'node:zlib';
import JSZip from 'jszip';
import {
  convertHtmlExportZipToDocx,
  DOCX_CONTENT_WIDTH_TWIPS,
} from '../apps/server/src/integrations/export/pandoc-docx';

function pngChunk(type: string, data: Buffer): Buffer {
  const body = Buffer.concat([Buffer.from(type), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, checksum]);
}

function widePng(): Buffer {
  const width = 1200;
  const height = 400;
  const row = Buffer.alloc(1 + width * 3, 0);
  for (let x = 0; x < width; x++) {
    row[1 + x * 3] = 0x33;
    row[2 + x * 3] = 0x66;
    row[3 + x * 3] = 0x99;
  }
  const raw = Buffer.concat(Array.from({ length: height }, () => Buffer.from(row)));
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

const attachmentId = '11111111-1111-4111-8111-111111111111';
const html = `<!DOCTYPE html>
<html>
  <head><title>Пример страницы</title></head>
  <body>
    <h1>Пример страницы</h1>
    <p></p>
    <table>
      <colgroup><col /><col /></colgroup>
      <tbody>
        <tr><th><p>Name</p></th><th><p>Поступления до / после, млн ₽</p></th></tr>
        <tr><td><p>Ada</p></td><td><p>Engineer</p></td></tr>
      </tbody>
    </table>
    <p><img src="files/${attachmentId}/photo.png" alt="diagram" /></p>
  </body>
</html>`;

async function main(): Promise<void> {
  const zip = new JSZip();
  zip.file('Пример страницы.html', html);
  zip.file(`/files/${attachmentId}/photo.png`, widePng());
  const docx = await convertHtmlExportZipToDocx(
    await zip.generateAsync({ type: 'nodebuffer' }),
  );
  const packed = await JSZip.loadAsync(docx);
  const media = Object.keys(packed.files).filter((name) =>
    name.startsWith('word/media/'),
  );
  if (media.length === 0) {
    throw new Error('docx is missing word/media/* image parts');
  }
  const documentXml = await packed.file('word/document.xml')?.async('string');
  if (!documentXml?.includes('w:tbl')) {
    throw new Error('docx document.xml is missing a w:tbl table');
  }
  if (!documentXml.includes('<w:tblBorders>') || !documentXml.includes('w:insideH')) {
    throw new Error('docx table is missing cell borders');
  }
  const titles = [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
    .map((match) =>
      [...match[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
        .map((text) => text[1])
        .join('')
        .replace(/\s+/g, ' ')
        .trim(),
    )
    .filter((text) => text === 'Пример страницы');
  if (titles.length !== 1) {
    throw new Error(`expected one title paragraph, found ${titles.length}`);
  }
  const columns = [...documentXml.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"/g)].map(
    (match) => Number(match[1]),
  );
  const width = columns.reduce((sum, column) => sum + column, 0);
  if (width !== DOCX_CONTENT_WIDTH_TWIPS) {
    throw new Error(`table width ${width} is not the A4 text width`);
  }

  console.log(`embedded images: ${media.join(', ')}`);
  console.log('title: one paragraph');
  console.log('word table: w:tbl with borders');
  console.log('docx export verification passed');
}

main().catch((error: unknown) => {
  console.error(error);
  process.exit(1);
});
