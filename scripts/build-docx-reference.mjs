/**
 * Build the pandoc reference document used by Word export.
 * Requires pandoc. Writes apps/server/src/integrations/export/reference.docx.
 *
 * Usage: node scripts/build-docx-reference.mjs
 */
import { execFileSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const outputPath = resolve(
  root,
  'apps/server/src/integrations/export/reference.docx',
);

const heading1 = `<w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="BodyText" />
    <w:uiPriority w:val="9" />
    <w:qFormat />
    <w:pPr>
      <w:keepNext />
      <w:keepLines />
      <w:spacing w:before="120" w:after="120" />
      <w:outlineLvl w:val="0" />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri" />
      <w:b />
      <w:bCs />
      <w:color w:val="1A1A1A" />
      <w:sz w:val="32" />
      <w:szCs w:val="32" />
    </w:rPr>
  </w:style>`;

const heading2 = `<w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="BodyText" />
    <w:uiPriority w:val="9" />
    <w:unhideWhenUsed />
    <w:qFormat />
    <w:pPr>
      <w:keepNext />
      <w:keepLines />
      <w:spacing w:before="200" w:after="80" />
      <w:outlineLvl w:val="1" />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri" />
      <w:b />
      <w:bCs />
      <w:color w:val="1A1A1A" />
      <w:sz w:val="26" />
      <w:szCs w:val="26" />
    </w:rPr>
  </w:style>`;

const title = `<w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title" />
    <w:basedOn w:val="Normal" />
    <w:next w:val="BodyText" />
    <w:qFormat />
    <w:pPr>
      <w:keepNext />
      <w:keepLines />
      <w:spacing w:before="0" w:after="160" />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri" />
      <w:b />
      <w:bCs />
      <w:color w:val="1A1A1A" />
      <w:sz w:val="32" />
      <w:szCs w:val="32" />
    </w:rPr>
  </w:style>`;

const bodyText = `<w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text" />
    <w:basedOn w:val="Normal" />
    <w:link w:val="BodyTextChar" />
    <w:pPr>
      <w:spacing w:before="0" w:after="120" />
    </w:pPr>
    <w:qFormat />
  </w:style>`;

const table = `<w:style w:type="table" w:default="1" w:styleId="Table">
    <w:name w:val="Table" />
    <w:basedOn w:val="TableNormal" />
    <w:qFormat />
    <w:tblPr>
      <w:tblW w:w="5000" w:type="pct" />
      <w:jc w:val="left" />
      <w:tblInd w:w="0" w:type="dxa" />
      <w:tblBorders>
        <w:top w:val="single" w:sz="8" w:space="0" w:color="808080" />
        <w:left w:val="single" w:sz="8" w:space="0" w:color="808080" />
        <w:bottom w:val="single" w:sz="8" w:space="0" w:color="808080" />
        <w:right w:val="single" w:sz="8" w:space="0" w:color="808080" />
        <w:insideH w:val="single" w:sz="8" w:space="0" w:color="808080" />
        <w:insideV w:val="single" w:sz="8" w:space="0" w:color="808080" />
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="60" w:type="dxa" />
        <w:left w:w="100" w:type="dxa" />
        <w:bottom w:w="60" w:type="dxa" />
        <w:right w:w="100" w:type="dxa" />
      </w:tblCellMar>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
      <w:rPr>
        <w:b />
        <w:bCs />
      </w:rPr>
      <w:tcPr>
        <w:shd w:val="clear" w:color="auto" w:fill="D6DCE4" />
        <w:vAlign w:val="center" />
      </w:tcPr>
    </w:tblStylePr>
  </w:style>`;

function replaceStyle(xml, styleId, block) {
  const token = `w:styleId="${styleId}"`;
  const idAt = xml.indexOf(token);
  if (idAt < 0) {
    throw new Error(`reference style not found: ${styleId}`);
  }
  const start = xml.lastIndexOf('<w:style', idAt);
  const end = xml.indexOf('</w:style>', idAt);
  if (start < 0 || end < 0) {
    throw new Error(`reference style bounds not found: ${styleId}`);
  }
  return xml.slice(0, start) + block + xml.slice(end + '</w:style>'.length);
}

const base = execFileSync('pandoc', ['--print-default-data-file', 'reference.docx']);
const zip = await JSZip.loadAsync(base);
let styles = await zip.file('word/styles.xml').async('string');
let documentXml = await zip.file('word/document.xml').async('string');

styles = replaceStyle(styles, 'Heading1', heading1);
styles = replaceStyle(styles, 'Heading2', heading2);
styles = replaceStyle(styles, 'Title', title);
styles = replaceStyle(styles, 'BodyText', bodyText);
styles = replaceStyle(styles, 'Table', table);
styles = styles.replace(
  /<w:rFonts\b[^>]*Theme="[^"]*"[^/]*\/>/g,
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri" />',
);
styles = styles.replace(
  /<w:color w:val="(?:4F81BD|345A8A)"[^/]*\/>/g,
  '<w:color w:val="1A1A1A" />',
);
styles = styles.replace(
  /<w:rFonts w:asciiTheme="minorHAnsi"[^/]*\/>/,
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri" />',
);
styles = styles.replace(
  /<w:sz w:val="24" \/>\s*<w:szCs w:val="24" \/>/,
  '<w:sz w:val="22" />\n        <w:szCs w:val="22" />',
);
styles = styles.replace(
  '<w:spacing w:after="200" />',
  '<w:spacing w:after="160" w:line="276" w:lineRule="auto" />',
);

if (!documentXml.includes('<w:sectPr />')) {
  throw new Error('reference document has no empty sectPr to replace');
}
documentXml = documentXml.replace(
  '<w:sectPr />',
  `<w:sectPr>
      <w:pgSz w:w="11906" w:h="16838" w:orient="portrait" />
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0" />
    </w:sectPr>`,
);

if (styles.includes('4F81BD') || styles.includes('asciiTheme')) {
  throw new Error('reference styles still use the default blue theme fonts');
}
if (!styles.includes('Calibri') || !styles.includes('w:insideH') || !styles.includes('D6DCE4')) {
  throw new Error('reference styles are missing Calibri, table borders, or header shading');
}

zip.file('word/styles.xml', styles);
zip.file('word/document.xml', documentXml);
const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
writeFileSync(outputPath, out);
console.log(`wrote ${outputPath} (${out.length} bytes)`);
