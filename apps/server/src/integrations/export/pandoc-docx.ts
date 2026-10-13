import { execFile } from 'child_process';
import { existsSync } from 'fs';
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

/** A4 with 2.54 cm margins. Matches reference.docx. Widths are twips. */
export const DOCX_PAGE_WIDTH_TWIPS = 11906;
export const DOCX_PAGE_HEIGHT_TWIPS = 16838;
export const DOCX_MARGIN_TWIPS = 1440;
export const DOCX_CONTENT_WIDTH_TWIPS =
  DOCX_PAGE_WIDTH_TWIPS - DOCX_MARGIN_TWIPS * 2;
const HEADER_FILL = 'D6DCE4';
const TABLE_BORDER =
  '<w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="808080"/><w:left w:val="single" w:sz="8" w:space="0" w:color="808080"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="808080"/><w:right w:val="single" w:sz="8" w:space="0" w:color="808080"/><w:insideH w:val="single" w:sz="8" w:space="0" w:color="808080"/><w:insideV w:val="single" w:sz="8" w:space="0" w:color="808080"/></w:tblBorders>';

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

    const docx = await polishDocx(await readFile(docxPath));
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

  splitImageGalleries($);
  removeEmptyParagraphs($);
  normalizeTables($);
  dropDuplicateTitle($);
  return $.html();
}

function splitImageGalleries($: ReturnType<typeof load>): void {
  $('p').each((_, element) => {
    const paragraph = $(element);
    const images = paragraph.children('img');
    if (images.length < 2) {
      return;
    }
    const onlyImages = paragraph.contents().toArray().every((node) => {
      if (node.type === 'text') {
        return !node.data?.trim();
      }
      return node.type !== 'tag' || node.name === 'img' || node.name === 'br';
    });
    if (!onlyImages) {
      return;
    }
    const replacement = images
      .toArray()
      .map((img) => `<p>${$.html(img)}</p>`)
      .join('');
    paragraph.replaceWith(replacement);
  });
}

function removeEmptyParagraphs($: ReturnType<typeof load>): void {
  $('p').each((_, element) => {
    const paragraph = $(element);
    if (paragraph.find('img, table, pre, ul, ol, svg').length > 0) {
      return;
    }
    const text = paragraph
      .text()
      .replace(/\u00a0/g, '')
      .trim();
    if (!text) {
      paragraph.remove();
    }
  });
}

function normalizeTables($: ReturnType<typeof load>): void {
  $('table').each((_, table) => {
    const $table = $(table);
    $table.find('colgroup').remove();
    if ($table.find('thead').length > 0) {
      return;
    }
    const firstRow = $table.find('tr').first();
    const cells = firstRow.children().toArray();
    const header =
      cells.length > 0 &&
      cells.every((cell) => cell.type === 'tag' && cell.name === 'th');
    if (!header) {
      return;
    }
    const thead = $('<thead></thead>');
    thead.append(firstRow);
    $table.prepend(thead);
  });
}

/**
 * The HTML export puts the page title in `<title>` and again as an H1.
 * Pandoc prints both. Keep the heading.
 */
function dropDuplicateTitle($: ReturnType<typeof load>): void {
  const normalize = (value: string) => value.replace(/\s+/g, ' ').trim();
  const title = normalize($('title').first().text());
  const heading = normalize($('h1').first().text());
  if (title && heading && title === heading) {
    $('title').remove();
  }
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

export function docxReferencePath(): string {
  return path.join(__dirname, 'reference.docx');
}

export function docxLuaFilterPath(): string {
  return path.join(__dirname, 'docx-export.lua');
}

async function runPandoc(
  htmlPath: string,
  docxPath: string,
  workDir: string,
): Promise<void> {
  const binary = process.env.PANDOC_PATH || 'pandoc';
  const referenceDoc = docxReferencePath();
  const luaFilter = docxLuaFilterPath();
  for (const asset of [referenceDoc, luaFilter]) {
    if (!existsSync(asset)) {
      throw new DocxConversionError(
        `Word export asset is missing: ${path.basename(asset)}`,
      );
    }
  }
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
        '--reference-doc',
        referenceDoc,
        '--lua-filter',
        luaFilter,
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

/**
 * Stamp visible borders, a shaded bold header, and page-width columns onto
 * every table. Pandoc's table style does not put those on the cells themselves.
 */
export async function polishDocx(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx);
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) {
    throw new DocxConversionError('pandoc did not produce a docx file');
  }
  const xml = await documentFile.async('string');
  zip.file('word/document.xml', polishDocumentXml(xml));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export function polishDocumentXml(xml: string): string {
  let cursor = 0;
  let result = '';
  while (cursor < xml.length) {
    const start = findTableOpen(xml, cursor);
    if (start < 0) {
      result += xml.slice(cursor);
      break;
    }
    result += xml.slice(cursor, start);
    const end = findTableClose(xml, start);
    result += polishTable(xml.slice(start, end));
    cursor = end;
  }
  return result;
}

function findTableOpen(xml: string, from: number): number {
  let index = from;
  while (index < xml.length) {
    const at = xml.indexOf('<w:tbl', index);
    if (at < 0) {
      return -1;
    }
    const next = xml[at + '<w:tbl'.length];
    if (next === '>' || next === ' ' || next === '\n' || next === '\r' || next === '\t') {
      return at;
    }
    index = at + '<w:tbl'.length;
  }
  return -1;
}

function findTableClose(xml: string, start: number): number {
  const open = '<w:tbl';
  const close = '</w:tbl>';
  let depth = 0;
  let index = start;
  while (index < xml.length) {
    const nextOpen = findTableOpen(xml, index === start ? start : index);
    const nextClose = xml.indexOf(close, index);
    if (nextClose < 0) {
      return xml.length;
    }
    if (nextOpen >= 0 && nextOpen < nextClose && nextOpen !== start) {
      depth += 1;
      index = nextOpen + open.length;
      continue;
    }
    if (nextOpen === start) {
      depth += 1;
      index = start + open.length;
      continue;
    }
    depth -= 1;
    index = nextClose + close.length;
    if (depth === 0) {
      return index;
    }
  }
  return xml.length;
}

function polishTable(table: string): string {
  const properties = table.match(/<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>/);
  if (!properties) {
    return scaleTableGrid(emphasizeHeaderRows(table));
  }
  let block = properties[0];
  if (!block.includes('<w:tblBorders')) {
    block = block.replace('</w:tblPr>', `${TABLE_BORDER}</w:tblPr>`);
  }
  const width = `<w:tblW w:w="${DOCX_CONTENT_WIDTH_TWIPS}" w:type="dxa"/>`;
  if (/<w:tblW\b[^>]*\/>/.test(block)) {
    block = block.replace(/<w:tblW\b[^>]*\/>/, width);
  } else if (/<w:tblW\b[\s\S]*?<\/w:tblW>/.test(block)) {
    block = block.replace(/<w:tblW\b[\s\S]*?<\/w:tblW>/, width);
  } else {
    block = block.replace(/(<w:tblPr\b[^>]*>)/, `$1${width}`);
  }
  if (!block.includes('<w:tblLayout')) {
    block = block.replace('</w:tblPr>', '<w:tblLayout w:type="fixed"/></w:tblPr>');
  }
  if (!block.includes('<w:tblCellMar')) {
    block = block.replace(
      '</w:tblPr>',
      '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>',
    );
  }
  return scaleTableGrid(emphasizeHeaderRows(table.replace(properties[0], block)));
}

function scaleTableGrid(table: string): string {
  const columns = [...table.matchAll(/<w:gridCol\b([^>]*)\/>/g)];
  if (columns.length === 0) {
    return table;
  }
  const widths = columns.map((match) => {
    const found = match[1].match(/w:w="(\d+)"/);
    return found ? Number(found[1]) : 1;
  });
  const sum = widths.reduce((total, width) => total + width, 0) || 1;
  let used = 0;
  const scaled = widths.map((width, index) => {
    if (index === widths.length - 1) {
      return Math.max(1, DOCX_CONTENT_WIDTH_TWIPS - used);
    }
    const next = Math.max(1, Math.round((width / sum) * DOCX_CONTENT_WIDTH_TWIPS));
    used += next;
    return next;
  });
  let index = 0;
  return table.replace(/<w:gridCol\b([^>]*)\/>/g, (_full, attrs: string) => {
    const width = scaled[index++];
    if (/w:w="\d+"/.test(attrs)) {
      return `<w:gridCol${attrs.replace(/w:w="\d+"/, `w:w="${width}"`)}/>`;
    }
    return `<w:gridCol${attrs} w:w="${width}"/>`;
  });
}

function emphasizeHeaderRows(table: string): string {
  return table.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
    if (!row.includes('<w:tblHeader')) {
      return row;
    }
    return row.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, emphasizeHeaderCell);
  });
}

function emphasizeHeaderCell(cell: string): string {
  const shading = `<w:shd w:val="clear" w:color="auto" w:fill="${HEADER_FILL}"/>`;
  let next = cell;
  if (!next.includes('<w:shd')) {
    if (/<w:tcPr\s*\/>/.test(next)) {
      next = next.replace(/<w:tcPr\s*\/>/, `<w:tcPr>${shading}</w:tcPr>`);
    } else if (/<w:tcPr\b[^>]*>/.test(next)) {
      next = next.replace(/<w:tcPr\b[^>]*>/, (open) => `${open}${shading}`);
    } else {
      next = next.replace(
        /^(<w:tc\b[^>]*>)/,
        `$1<w:tcPr>${shading}</w:tcPr>`,
      );
    }
  }
  next = next.replace(
    /<w:rPr(\s[^>]*)?>([\s\S]*?)<\/w:rPr>/g,
    (full, attrs: string | undefined, inner: string) => {
      if (/<w:b[\s/>]/.test(inner)) {
        return full;
      }
      return `<w:rPr${attrs ?? ''}><w:b/><w:bCs/>${inner}</w:rPr>`;
    },
  );
  next = next.replace(
    /<w:r(\s[^>]*)?>\s*(?!<w:rPr\b)/g,
    '<w:r$1><w:rPr><w:b/><w:bCs/></w:rPr>',
  );
  return next;
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
