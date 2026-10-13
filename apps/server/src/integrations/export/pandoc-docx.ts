import { execFile } from 'child_process';
import { mkdir, mkdtemp, readdir, readFile, rm, writeFile } from 'fs/promises';
import { tmpdir } from 'os';
import * as path from 'path';
import { promisify } from 'util';
import { load } from 'cheerio';
import * as JSZip from 'jszip';
import { sanitizeFileName } from '../../common/helpers/utils';

const execFileAsync = promisify(execFile);

const PANDOC_TIMEOUT_MS = 120_000;
const WORK_DIR_PREFIX = 'docmost-docx-';

export class PandocNotFoundError extends Error {
  constructor() {
    super('pandoc is not installed');
    this.name = 'PandocNotFoundError';
  }
}

export class DocxConversionError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'DocxConversionError';
  }
}

/**
 * Download name for a page export. Unicode titles, including Cyrillic, are
 * kept; the caller percent-encodes the name in Content-Disposition.
 */
export function docxDownloadFileName(title: string | null | undefined): string {
  const safe = sanitizeFileName((title ?? '').trim() || 'untitled', {
    preserveSpaces: true,
  });
  return `${safe || 'untitled'}.docx`;
}

/**
 * Matches the other export endpoints: an ASCII `filename` parameter that the
 * client `decodeURIComponent`s, so non-ASCII titles survive the header.
 */
export function docxContentDisposition(fileName: string): string {
  return `attachment; filename="${encodeURIComponent(fileName)}"`;
}

/**
 * Turn an HTML export zip (page HTML plus `files/{id}/{name}` attachments)
 * into a .docx buffer. Temporary files are removed before this returns.
 */
export async function convertHtmlExportZipToDocx(
  zipBuffer: Buffer,
): Promise<Buffer> {
  const workDir = await mkdtempSafe();
  try {
    await extractZipSafely(zipBuffer, workDir);
    const htmlName = await findTopLevelHtml(workDir);
    const htmlPath = path.join(workDir, htmlName);
    const html = await readFile(htmlPath, 'utf8');
    await writeFile(htmlPath, prepareHtmlForPandoc(html), 'utf8');

    const docxPath = path.join(workDir, 'page.docx');
    await runPandoc(htmlPath, docxPath, workDir);

    const docx = await readFile(docxPath);
    if (docx.length < 4 || docx[0] !== 0x50 || docx[1] !== 0x4b) {
      throw new DocxConversionError('pandoc did not produce a docx file');
    }
    return docx;
  } finally {
    await rm(workDir, { recursive: true, force: true });
  }
}

export function prepareHtmlForPandoc(html: string): string {
  const $ = load(html);
  $('img').each((_, element) => {
    const img = $(element);
    const src = img.attr('src');
    if (!src) {
      return;
    }

    const localSrc = toLocalResourceSrc(src);
    if (localSrc == null) {
      const alt = img.attr('alt')?.trim() || src;
      const link = $('<a></a>').attr('href', src).text(alt);
      img.replaceWith(link);
      return;
    }

    if (localSrc !== src) {
      img.attr('src', localSrc);
    }
  });
  return $.html();
}

/**
 * Local attachment paths pandoc can embed. Remote URLs return null so the
 * caller can leave them as links instead of letting pandoc fetch them.
 * `data:` images are left untouched; pandoc embeds those itself.
 */
export function toLocalResourceSrc(src: string): string | null {
  const trimmed = src.trim();
  if (!trimmed || /[\0\\]/.test(trimmed)) {
    return null;
  }
  if (/^data:/i.test(trimmed)) {
    return trimmed;
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed) || trimmed.startsWith('//')) {
    return null;
  }

  const withoutSuffix = trimmed.split('#')[0].split('?')[0];
  let decoded = withoutSuffix;
  try {
    decoded = decodeURIComponent(withoutSuffix);
  } catch {
    decoded = withoutSuffix;
  }

  let relative = decoded;
  if (relative.startsWith('/api/files/')) {
    relative = `files/${relative.slice('/api/files/'.length)}`;
  } else if (relative.startsWith('/files/')) {
    relative = `files/${relative.slice('/files/'.length)}`;
  } else if (relative.startsWith('api/files/')) {
    relative = `files/${relative.slice('api/files/'.length)}`;
  } else if (relative.startsWith('/')) {
    return null;
  }

  if (relative.split('/').includes('..')) {
    return null;
  }

  return relative;
}

function mkdtempSafe(): Promise<string> {
  return mkdtemp(path.join(tmpdir(), WORK_DIR_PREFIX));
}

async function extractZipSafely(
  zipBuffer: Buffer,
  root: string,
): Promise<void> {
  const zip = await JSZip.loadAsync(zipBuffer);
  for (const entry of Object.values(zip.files)) {
    const relative = normalizeZipEntry(entry.name);
    if (!relative) {
      continue;
    }
    const target = resolveInside(root, relative);
    if (entry.dir || relative.endsWith('/')) {
      await mkdir(target, { recursive: true });
      continue;
    }
    await mkdir(path.dirname(target), { recursive: true });
    const data = await entry.async('nodebuffer');
    await writeFile(target, data);
  }
}

function normalizeZipEntry(name: string): string | null {
  const normalized = name.replace(/\\/g, '/').replace(/^\/+/, '');
  if (
    !normalized ||
    normalized.includes('\0') ||
    normalized.split('/').includes('..')
  ) {
    return null;
  }
  return normalized;
}

function resolveInside(root: string, relative: string): string {
  const target = path.resolve(root, relative);
  const rootWithSep = root.endsWith(path.sep) ? root : root + path.sep;
  if (target !== root && !target.startsWith(rootWithSep)) {
    throw new DocxConversionError('HTML export archive contains an invalid path');
  }
  return target;
}

async function findTopLevelHtml(dir: string): Promise<string> {
  const entries = await readdir(dir, { withFileTypes: true });
  const html = entries.find(
    (entry) => entry.isFile() && entry.name.toLowerCase().endsWith('.html'),
  );
  if (!html) {
    throw new DocxConversionError('HTML export did not contain a page');
  }
  return html.name;
}

async function runPandoc(
  htmlPath: string,
  docxPath: string,
  workDir: string,
): Promise<void> {
  const binary = process.env.PANDOC_PATH || 'pandoc';
  try {
    await execFileAsync(
      binary,
      [
        htmlPath,
        '--from=html',
        '--to=docx',
        '--output',
        docxPath,
        '--resource-path',
        workDir,
      ],
      {
        cwd: workDir,
        timeout: PANDOC_TIMEOUT_MS,
        maxBuffer: 8 * 1024 * 1024,
      },
    );
  } catch (error) {
    const err = error as NodeJS.ErrnoException & { stderr?: string | Buffer };
    if (err.code === 'ENOENT') {
      throw new PandocNotFoundError();
    }
    const stderr = Buffer.isBuffer(err.stderr)
      ? err.stderr.toString('utf8')
      : (err.stderr ?? '');
    const detail = stderr.trim() || err.message || 'pandoc failed';
    throw new DocxConversionError(detail.slice(0, 2000));
  }
}

export async function readStreamToBuffer(
  stream: NodeJS.ReadableStream,
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    stream.on('data', (chunk: Buffer | string) => {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
    });
    stream.on('end', () => resolve(Buffer.concat(chunks)));
    stream.on('error', reject);
  });
}
