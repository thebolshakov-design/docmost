import { execFileSync } from 'child_process';
import { readdirSync } from 'fs';
import { tmpdir } from 'os';
import { crc32, deflateSync } from 'zlib';
import * as JSZip from 'jszip';
import {
  convertHtmlExportZipToDocx,
  DOCX_CONTENT_WIDTH_TWIPS,
  DOCX_PAGE_HEIGHT_TWIPS,
  DOCX_PAGE_WIDTH_TWIPS,
  docxContentDisposition,
  docxDownloadFileName,
  PandocNotFoundError,
  prepareHtmlForPandoc,
  toLocalResourceSrc,
} from './pandoc-docx';

function pngChunk(type: string, data: Buffer): Buffer {
  const body = Buffer.concat([Buffer.from(type), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, checksum]);
}

/** 1200×400 PNG, wider than an A4 text column at 96 dpi. */
function widePng(): Buffer {
  const width = 1200;
  const height = 400;
  const row = Buffer.alloc(1 + width * 3);
  row[0] = 0;
  for (let x = 0; x < width; x++) {
    row[1 + x * 3] = 0x33;
    row[2 + x * 3] = 0x66;
    row[3 + x * 3] = 0x99;
  }
  const raw = Buffer.concat(Array.from({ length: height }, () => row));
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  return Buffer.concat([
    signature,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}
const ATTACHMENT_ID = '11111111-1111-4111-8111-111111111111';

function pandocAvailable(): boolean {
  try {
    execFileSync(process.env.PANDOC_PATH || 'pandoc', ['--version'], {
      stdio: 'ignore',
    });
    return true;
  } catch {
    return false;
  }
}

function tempDirs(): string[] {
  return readdirSync(tmpdir()).filter((name) => name.startsWith('docmost-docx-'));
}

async function sampleExportZip(): Promise<Buffer> {
  const html = `<!DOCTYPE html>
<html>
  <head><title>Пример страницы</title></head>
  <body>
    <h1>Пример страницы</h1>
    <p></p>
    <p><br></p>
    <h2>Section</h2>
    <ul><li><p>First item</p></li></ul>
    <ol><li><p>Step one</p></li></ol>
    <pre><code class="language-js">const answer = 42;</code></pre>
    <table>
      <colgroup><col style="min-width: 25px;" /><col style="min-width: 25px;" /></colgroup>
      <tbody>
        <tr><th><p>Name</p></th><th><p>Поступления до / после, млн ₽</p></th></tr>
        <tr><td><p>Ada</p></td><td><p>Engineer</p></td></tr>
      </tbody>
    </table>
    <p><img src="files/${ATTACHMENT_ID}/photo.png?t=123" alt="diagram" /></p>
    <p><img src="https://example.com/remote.png" alt="remote diagram" /></p>
  </body>
</html>`;

  const zip = new JSZip();
  zip.file('Пример страницы.html', html);
  zip.file(`/files/${ATTACHMENT_ID}/photo.png`, widePng());
  zip.file('docmost-metadata.json', '{}');
  return zip.generateAsync({ type: 'nodebuffer' });
}

function paragraphTexts(documentXml: string): string[] {
  return [...documentXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((match) =>
    [...match[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
      .map((text) => text[1])
      .join(''),
  );
}

describe('docx download names', () => {
  it('keeps Cyrillic titles and percent-encodes the header', () => {
    const fileName = docxDownloadFileName('Отчёт за квартал');
    expect(fileName).toBe('Отчёт за квартал.docx');

    const header = docxContentDisposition(fileName);
    expect([...header].every((char) => char.charCodeAt(0) <= 127)).toBe(true);
    const encoded = header.split('filename=')[1].replace(/"/g, '');
    expect(decodeURIComponent(encoded)).toBe('Отчёт за квартал.docx');
  });

  it('falls back when the title is empty or unsafe', () => {
    expect(docxDownloadFileName(null)).toBe('untitled.docx');
    expect(docxDownloadFileName('   ')).toBe('untitled.docx');
    expect(docxDownloadFileName('..')).toBe('untitled.docx');
  });
});

describe('prepareHtmlForPandoc', () => {
  it('rewrites attachment urls onto the exported files/ tree', () => {
    expect(toLocalResourceSrc(`/api/files/${ATTACHMENT_ID}/photo.png?t=9`)).toBe(
      `files/${ATTACHMENT_ID}/photo.png`,
    );
    expect(toLocalResourceSrc(`/files/${ATTACHMENT_ID}/my%20photo.png`)).toBe(
      `files/${ATTACHMENT_ID}/my photo.png`,
    );
    expect(toLocalResourceSrc('https://example.com/a.png')).toBeNull();
    expect(toLocalResourceSrc('//cdn.example.com/a.png')).toBeNull();
    expect(toLocalResourceSrc('files/../../etc/passwd')).toBeNull();
    expect(toLocalResourceSrc('data:image/png;base64,aaaa')).toBe(
      'data:image/png;base64,aaaa',
    );
  });

  it('keeps one title, promotes header rows, and drops empty paragraphs', () => {
    const html = prepareHtmlForPandoc(
      '<html><head><title>Пример страницы</title></head><body>' +
        '<h1>Пример страницы</h1><p></p><p><br></p>' +
        '<table><colgroup><col /></colgroup><tbody>' +
        '<tr><th>Name</th><th>Role</th></tr>' +
        '<tr><td>Ada</td><td>Engineer</td></tr>' +
        '</tbody></table></body></html>',
    );

    expect(html).not.toContain('<title');
    expect(html).toContain('<h1>Пример страницы</h1>');
    expect(html).toContain('<thead>');
    expect(html).not.toContain('<colgroup');
    expect(html).not.toContain('<p></p>');
    expect(html).not.toContain('<br');
  });

  it('turns remote images into links and strips cache-busting query strings', () => {
    const html = prepareHtmlForPandoc(
      `<p><img src="files/${ATTACHMENT_ID}/photo.png?t=123" alt="diagram" /></p>` +
        '<p><img src="https://example.com/remote.png" alt="remote diagram" /></p>',
    );

    expect(html).toContain(`src="files/${ATTACHMENT_ID}/photo.png"`);
    expect(html).not.toContain('photo.png?t=');
    expect(html).not.toContain('src="https://example.com/remote.png"');
    expect(html).toContain('href="https://example.com/remote.png"');
    expect(html).toContain('remote diagram');
  });
});

describe('convertHtmlExportZipToDocx', () => {
  const hasPandoc = pandocAvailable();

  it('removes the temp directory when pandoc cannot be started', async () => {
    const previous = process.env.PANDOC_PATH;
    process.env.PANDOC_PATH = '/nonexistent/pandoc-binary';
    const before = new Set(tempDirs());
    try {
      await expect(convertHtmlExportZipToDocx(await sampleExportZip())).rejects.toBeInstanceOf(
        PandocNotFoundError,
      );
    } finally {
      if (previous === undefined) {
        delete process.env.PANDOC_PATH;
      } else {
        process.env.PANDOC_PATH = previous;
      }
    }
    expect(tempDirs().filter((name) => !before.has(name))).toEqual([]);
  });

  (hasPandoc ? it : it.skip)(
    'embeds the image and keeps the table',
    async () => {
      const before = new Set(tempDirs());
      const docx = await convertHtmlExportZipToDocx(await sampleExportZip());
      expect(tempDirs().filter((name) => !before.has(name))).toEqual([]);

      const zip = await JSZip.loadAsync(docx);
      const media = Object.keys(zip.files).filter((name) =>
        name.startsWith('word/media/'),
      );
      expect(media).toHaveLength(1);

      const documentXml = await zip.file('word/document.xml')?.async('string');
      const stylesXml = await zip.file('word/styles.xml')?.async('string');
      expect(documentXml).toEqual(expect.any(String));
      expect(stylesXml).toContain('Calibri');
      expect(documentXml).toContain('w:tbl');
      expect(documentXml).toContain('<w:tblBorders>');
      expect(documentXml).toContain('w:insideH');
      expect(documentXml).toContain('w:insideV');
      expect(documentXml).toContain('w:fill="D6DCE4"');
      expect(documentXml).toContain('<w:tblHeader');
      expect(documentXml).toContain(`w:w="${DOCX_PAGE_WIDTH_TWIPS}"`);
      expect(documentXml).toContain(`w:h="${DOCX_PAGE_HEIGHT_TWIPS}"`);
      expect(documentXml).toContain('Ada');
      expect(documentXml).toContain('First item');
      expect(documentXml).toContain('42');
      expect(documentXml).not.toContain('example.com/remote.png');

      const titles = paragraphTexts(documentXml!).filter(
        (text) => text.replace(/\s+/g, ' ').trim() === 'Пример страницы',
      );
      expect(titles).toHaveLength(1);
      expect(documentXml).not.toContain('w:val="Title"');

      const columns = [
        ...documentXml!.matchAll(/<w:gridCol\b[^>]*w:w="(\d+)"/g),
      ].map((match) => Number(match[1]));
      expect(columns.reduce((sum, width) => sum + width, 0)).toBe(
        DOCX_CONTENT_WIDTH_TWIPS,
      );
      expect(Math.max(...columns)).toBeGreaterThan(Math.min(...columns));

      const image = documentXml!.match(/<wp:extent\b[^>]*cx="(\d+)"/);
      expect(image).not.toBeNull();
      const widthInches = Number(image![1]) / 914400;
      expect(widthInches).toBeGreaterThan(6);
      expect(widthInches).toBeLessThan(6.5);

      const emptyParagraphs = [
        ...documentXml!.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g),
      ].filter((match) => {
        const text = [...match[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
          .map((part) => part[1])
          .join('')
          .trim();
        return !text && !match[0].includes('<w:drawing');
      });
      expect(emptyParagraphs).toHaveLength(0);
    },
    20_000,
  );

  it('rejects an export archive that has no html page', async () => {
    const zip = new JSZip();
    zip.file('notes.txt', 'no page');
    const before = new Set(tempDirs());
    await expect(
      convertHtmlExportZipToDocx(await zip.generateAsync({ type: 'nodebuffer' })),
    ).rejects.toThrow(/did not contain a page/);
    expect(tempDirs().filter((name) => !before.has(name))).toEqual([]);
  });
});
