/**
 * Sample HTML page -> pandoc -> .docx.
 * Confirms embedded images (word/media/*) and a real Word table (w:tbl).
 *
 * Usage: node scripts/verify-docx-export.mjs
 */
import { execFileSync } from 'node:child_process';
import {
  mkdtempSync,
  mkdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

const dir = mkdtempSync(path.join(tmpdir(), 'verify-docx-'));

try {
  const imageDir = path.join(
    dir,
    'files',
    '11111111-1111-4111-8111-111111111111',
  );
  mkdirSync(imageDir, { recursive: true });
  writeFileSync(path.join(imageDir, 'photo.png'), PNG);

  const htmlPath = path.join(dir, 'page.html');
  writeFileSync(
    htmlPath,
    `<!DOCTYPE html>
<html>
  <head><title>Пример страницы</title></head>
  <body>
    <h1>Пример страницы</h1>
    <h2>Section</h2>
    <ul><li><p>First item</p></li></ul>
    <pre><code>const answer = 42;</code></pre>
    <table>
      <tbody>
        <tr><th><p>Name</p></th><th><p>Role</p></th></tr>
        <tr><td><p>Ada</p></td><td><p>Engineer</p></td></tr>
      </tbody>
    </table>
    <p><img src="files/11111111-1111-4111-8111-111111111111/photo.png" alt="diagram" /></p>
  </body>
</html>`,
  );

  const docxPath = path.join(dir, 'page.docx');
  execFileSync(
    process.env.PANDOC_PATH || 'pandoc',
    [
      htmlPath,
      '--from=html',
      '--to=docx',
      '--output',
      docxPath,
      '--resource-path',
      dir,
    ],
    { cwd: dir, stdio: 'inherit' },
  );

  const listing = execFileSync('unzip', ['-Z1', docxPath], { encoding: 'utf8' });
  const media = listing
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.startsWith('word/media/'));
  if (media.length === 0) {
    throw new Error('docx is missing word/media/* image parts');
  }

  const documentXml = execFileSync(
    'unzip',
    ['-p', docxPath, 'word/document.xml'],
    { encoding: 'utf8' },
  );
  if (!documentXml.includes('w:tbl')) {
    throw new Error('docx document.xml is missing a w:tbl table');
  }

  console.log(`embedded images: ${media.join(', ')}`);
  console.log('word table: w:tbl present');
  console.log('docx export verification passed');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
