import { execFileSync } from 'child_process';
import { readdirSync } from 'fs';
import { tmpdir } from 'os';
import * as JSZip from 'jszip';
import {
  convertHtmlExportZipToDocx,
  docxContentDisposition,
  docxDownloadFileName,
  PandocNotFoundError,
  prepareHtmlForPandoc,
  toLocalResourceSrc,
} from './pandoc-docx';

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);
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
    <h2>Section</h2>
    <ul><li><p>First item</p></li></ul>
    <ol><li><p>Step one</p></li></ol>
    <pre><code class="language-js">const answer = 42;</code></pre>
    <table>
      <colgroup><col /><col /></colgroup>
      <tbody>
        <tr><th><p>Name</p></th><th><p>Role</p></th></tr>
        <tr><td><p>Ada</p></td><td><p>Engineer</p></td></tr>
      </tbody>
    </table>
    <p><img src="files/${ATTACHMENT_ID}/photo.png?t=123" alt="diagram" /></p>
    <p><img src="https://example.com/remote.png" alt="remote diagram" /></p>
  </body>
</html>`;

  const zip = new JSZip();
  zip.file('Пример страницы.html', html);
  zip.file(`/files/${ATTACHMENT_ID}/photo.png`, PNG);
  zip.file('docmost-metadata.json', '{}');
  return zip.generateAsync({ type: 'nodebuffer' });
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
      expect(documentXml).toEqual(expect.any(String));
      expect(documentXml).toContain('w:tbl');
      expect(documentXml).toContain('Ada');
      expect(documentXml).toContain('First item');
      expect(documentXml).toContain('42');
      expect(documentXml).toContain('Пример');
      expect(documentXml).not.toContain('example.com/remote.png');
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
