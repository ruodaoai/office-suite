#!/usr/bin/env node
/**
 * Ruodao office-suite 插件命令入口（零第三方运行依赖）
 *
 * vendor/ 内嵌纯 JS 库（随插件分发，不走 npm install）：
 *   - jszip.min.js  3.10.1（OOXML 读写）
 *   - xlsx.full.min.js 0.18.5（SheetJS：xlsx/xls/csv 读写）
 *
 * 命令（stdin 传 JSON，stdout 输出单行 JSON：state: succeeded | failed）：
 *   doctor        环境自检
 *   docx2text     docx → 结构化文本（段落/标题/表格）
 *   docx2rich     docx → 富文本模型（标题/列表/表格/图片/行内格式）
 *   docx-save-rich 富文本模型 → docx（可基于原文件回存，保留页面设置与样式）
 *   docx-edit     docx 文本替换后另存（保 OOXML 结构）
 *   docx-build    纯文本块 → docx
 *   xlsx2json     xlsx/xls/csv → 每个 sheet 的 JSON 行
 *   xlsx2csv      xlsx → csv
 *   xlsx-edit     修改单元格后另存
 *   csv2xlsx      csv → xlsx
 *   pptx2json     pptx → 幻灯片文本大纲（形状文本/表格/备注）
 *   pptx-build    大纲 JSON → 生成 pptx（支持主题、版式、16:9）
 *   pdf-note      PDF 元信息探测
 */

import { createRequire } from "node:module";
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chatEndpoint, runAgent } from "../lib/agent-core.mjs";
import { compact, contextText, preparePayload, systemPrompt, tools as agentTools } from "../lib/agent-tools.mjs";

const require = createRequire(import.meta.url);
const PLUGIN_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const VERSION = "1.3.0";
const JSZip = require(path.join(PLUGIN_DIR, "vendor", "jszip.min.js"));
const XLSX = require(path.join(PLUGIN_DIR, "vendor", "xlsx.full.min.js"));

const emit = (payload) => process.stdout.write(JSON.stringify(payload) + "\n");
const fail = (error, extra = {}) => { emit({ state: "failed", error, ...extra }); process.exit(1); };
const env = (k, d = "") => (process.env[k] ?? d).trim();

function readStdin() {
  return new Promise((resolve) => {
    let data = "";
    if (process.stdin.isTTY) return resolve({});
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (c) => (data += c));
    process.stdin.on("end", () => {
      try { resolve(data.trim() ? JSON.parse(data) : {}); }
      catch { resolve({ prompt: data.trim() }); }
    });
    process.stdin.on("error", () => resolve({}));
  });
}

async function ensureReadable(inputPath) {
  const target = path.resolve(String(inputPath ?? ""));
  const stat = await fsp.stat(target).catch(() => null);
  if (!stat || !stat.isFile()) throw new Error(`文件不存在：${target}`);
  if (stat.size > 200 * 1024 * 1024) throw new Error(`文件过大（${Math.round(stat.size / 1048576)}MB），上限 200MB`);
  return target;
}

/* ---------------- docx ---------------- */

// 段落文本：拼接 <w:t>，<w:tab>/<w:br> 转空格/换行
function paragraphText(paragraphXml) {
  let text = "";
  const re = /<w:(t|tab|br)\b[^>]*(?:\/>|>([\s\S]*?)<\/w:\1>)/g;
  let match;
  while ((match = re.exec(paragraphXml))) {
    if (match[1] === "tab") text += "\t";
    else if (match[1] === "br") text += "\n";
    else text += (match[2] ?? "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  }
  return text;
}

function headingLevel(paragraphXml) {
  const m = /<w:pStyle w:val="(?:Heading|heading)(\d)"/.exec(paragraphXml);
  return m ? Number(m[1]) : 0;
}

async function docx2text(input) {
  const file = await ensureReadable(input.path);
  const buf = await fsp.readFile(file);
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("不是有效的 docx（缺少 word/document.xml）");
  const blocks = [];
  const paragraphRe = /<w:p\b[^>]*(?:\/>|[\s\S]*?<\/w:p>)/g;
  let match;
  while ((match = paragraphRe.exec(docXml))) {
    const xml = match[0];
    const text = paragraphText(xml);
    const level = headingLevel(xml);
    const inTable = /<w:tbl>[\s\S]*<\/w:tbl>/.test(xml) || docXml.lastIndexOf("<w:tbl>", match.index) > docXml.lastIndexOf("</w:tbl>", match.index);
    if (text.trim()) blocks.push({ type: level ? "heading" : inTable ? "table-cell" : "paragraph", level, text });
  }
  const tables = (docXml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) ?? []).length;
  const media = Object.keys(zip.files).filter((name) => name.startsWith("word/media/")).length;
  return { file, blocks, stats: { blocks: blocks.length, tables, media, chars: blocks.reduce((sum, b) => sum + b.text.length, 0) } };
}

async function docxEdit(input) {
  const file = await ensureReadable(input.path);
  const replacements = Array.isArray(input.replacements) ? input.replacements : [];
  if (!replacements.length) throw new Error("replacements 不能为空：[{ from, to }]");
  const buf = await fsp.readFile(file);
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("不是有效的 docx");
  // Word 可能把一个词拆进多个 run；先按 <w:t> 内容做整体替换，未命中的再尝试去掉跨 run 拆分
  const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let next = docXml;
  let applied = 0;
  for (const { from, to } of replacements) {
    if (!from) continue;
    const re = new RegExp(escapeRe(from), "g");
    if (re.test(next)) { next = next.replace(re, () => String(to ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")); applied += 1; }
  }
  if (!applied) throw new Error("未找到任何待替换文本（Word 可能把文本拆分在多个 run 中，可尝试更短的片段）");
  zip.file("word/document.xml", next);
  const out = path.resolve(String(input.output ?? path.join(path.dirname(file), `${path.basename(file, path.extname(file))}-edited.docx`)));
  await fsp.mkdir(path.dirname(out), { recursive: true });
  await fsp.writeFile(out, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return { file, output: out, applied };
}

const DOCX_ML = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function docxParagraphXml(block) {
  const type = String(block?.type ?? "paragraph");
  const text = String(block?.text ?? "");
  if (!text.trim()) return "<w:p/>";
  const style = type === "heading" ? ` <w:pPr><w:pStyle w:val="Heading${Math.min(Math.max(Number(block?.level ?? 1), 1), 3)}"/></w:pPr>` : "";
  return `<w:p>${style}<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${DOCX_ML(text)}</w:t></w:r></w:p>`;
}

async function docxBuild(input) {
  const blocks = Array.isArray(input.blocks) ? input.blocks : [];
  if (!blocks.length) throw new Error("blocks 不能为空：[{ type: 'heading'|'paragraph', level?, text }]");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${DOCX_ML(input.title ?? "Ruodao 文档")}</dc:title><dc:creator>Ruodao office-suite</dc:creator></cp:coreProperties>`);
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Ruodao office-suite</Application></Properties>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  const headingStyle = (level, size) => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="${level - 1}"/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="${size}"/></w:rPr></w:style>`;
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:style>${headingStyle(1, 36)}${headingStyle(2, 30)}${headingStyle(3, 26)}</w:styles>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${blocks.slice(0, 5000).map(docxParagraphXml).join("")}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  const output = path.resolve(String(input.output ?? path.join(process.cwd(), `ruodao-${Date.now()}.docx`)));
  await fsp.mkdir(path.dirname(output), { recursive: true });
  await fsp.writeFile(output, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return { output, blocks: blocks.length };
}

/* ---------------- docx 富文本（保留格式的导入/回存） ---------------- */

const DOCX_XML_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

const IMAGE_MIME = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml",
};
const MAX_EMBED_IMAGES = 24;
const MAX_EMBED_IMAGE_BYTES = 12 * 1024 * 1024;

function decodeXmlText(value) {
  return String(value ?? "")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}

function escapeXmlText(value) {
  return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function xmlAttrs(fragment) {
  const attrs = {};
  for (const match of String(fragment ?? "").matchAll(/([A-Za-z_][\w:.-]*)="([^"]*)"/g)) attrs[match[1]] = decodeXmlText(match[2]);
  return attrs;
}

function onOffTag(rPr, tag) {
  const match = new RegExp(`<w:${tag}\\b([^>]*)>`).exec(rPr);
  if (!match) return false;
  const value = /w:val="([^"]*)"/.exec(match[1])?.[1];
  return value === undefined || !["0", "false", "off", "none"].includes(String(value).toLowerCase());
}

function parseParagraphRuns(rPr) {
  const props = {};
  if (onOffTag(rPr, "b")) props.b = true;
  if (onOffTag(rPr, "i")) props.i = true;
  if (onOffTag(rPr, "strike") || onOffTag(rPr, "dstrike")) props.s = true;
  const underline = /<w:u\b([^>]*)>/.exec(rPr);
  if (underline) {
    const value = /w:val="([^"]*)"/.exec(underline[1])?.[1] ?? "single";
    if (!["none", "0", "false"].includes(String(value).toLowerCase())) props.u = true;
  }
  const color = /<w:color\b[^>]*w:val="([0-9a-fA-F]{6})"/.exec(rPr)?.[1];
  if (color && color.toLowerCase() !== "auto") props.color = `#${color.toLowerCase()}`;
  const size = Number(/<w:sz\b[^>]*w:val="(\d+)"/.exec(rPr)?.[1] ?? 0);
  if (size > 0) props.size = Math.round((size / 2) * (4 / 3) * 10) / 10;
  const fonts = /<w:rFonts\b([^>]*)>/.exec(rPr)?.[1] ?? "";
  const fontAttrs = xmlAttrs(fonts);
  const font = fontAttrs["w:eastAsia"] || fontAttrs["w:ascii"] || fontAttrs["w:hAnsi"];
  if (font && !/default/i.test(font)) props.font = font;
  const highlight = /<w:highlight\b[^>]*w:val="([^"]+)"/.exec(rPr)?.[1];
  const highlightHex = { yellow: "#fef08a", green: "#bbf7d0", cyan: "#a5f3fc", magenta: "#fbcfe8", blue: "#bfdbfe", red: "#fecaca", darkBlue: "#1e3a8a", darkCyan: "#155e75", darkGreen: "#166534", darkMagenta: "#86198f", darkRed: "#991b1b", darkYellow: "#854d0e", gray: "#e5e7eb", lightGray: "#f3f4f6", black: "#000000" }[highlight];
  if (highlightHex) props.highlight = highlightHex;
  const shd = /<w:shd\b[^>]*w:fill="([0-9a-fA-F]{6})"/.exec(rPr)?.[1];
  if (shd && !props.highlight) props.highlight = `#${shd.toLowerCase()}`;
  return props;
}

function parseStylesHeadings(stylesXml) {
  const map = new Map();
  if (!stylesXml) return map;
  for (const match of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const attrs = xmlAttrs(match[1]);
    const id = attrs["w:styleId"];
    if (!id) continue;
    const name = xmlAttrs(/<w:name\b([^>]*)\/?>/.exec(match[2])?.[1] ?? "")["w:val"] ?? "";
    const heading = /heading\s*([1-9])/i.exec(name)?.[1] ?? /^Heading\s*([1-9])$/i.exec(id)?.[1];
    if (heading) map.set(id, Math.min(Number(heading), 3));
    else if (/^(?:Title|标题)$/i.test(name)) map.set(id, 1);
  }
  return map;
}

function parseNumberingXml(xml) {
  const numToAbstract = new Map();
  const formats = new Map();
  if (!xml) return { numToAbstract, formats };
  for (const match of xml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
    const abstractId = /<w:abstractNumId w:val="(\d+)"/.exec(match[2])?.[1];
    if (abstractId) numToAbstract.set(match[1], abstractId);
  }
  for (const match of xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
    const levels = [];
    for (const level of match[2].matchAll(/<w:lvl\b[^>]*w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g)) {
      levels[Number(level[1])] = /<w:numFmt\b[^>]*w:val="([^"]+)"/.exec(level[2])?.[1] ?? "bullet";
    }
    formats.set(match[1], levels);
  }
  return { numToAbstract, formats };
}

function docxRelationships(xml) {
  const map = new Map();
  if (!xml) return map;
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attrs = xmlAttrs(match[1]);
    if (attrs.Id) map.set(attrs.Id, { target: attrs.Target ?? "", type: attrs.Type ?? "" });
  }
  return map;
}

async function parseDocxRun(runXml, ctx) {
  const rPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(runXml)?.[0] ?? "";
  const props = parseParagraphRuns(rPr);
  const images = [];
  for (const blip of runXml.matchAll(/<a:blip\b[^>]*r:embed="([^"]+)"/g)) {
    const rel = ctx.rels.get(blip[1]);
    if (!rel?.target) continue;
    const target = path.posix.normalize(`word/${rel.target}`).replace(/^word\/\.\.\//, "");
    const entry = ctx.zip.file(target);
    if (!entry) continue;
    if (ctx.imageCount >= MAX_EMBED_IMAGES || ctx.imageBytes >= MAX_EMBED_IMAGE_BYTES) { ctx.skippedImages += 1; continue; }
    const ext = (target.split(".").pop() ?? "").toLowerCase();
    const mime = IMAGE_MIME[ext];
    if (!mime) { ctx.skippedImages += 1; continue; }
    const buffer = await entry.async("nodebuffer");
    if (ctx.imageBytes + buffer.length > MAX_EMBED_IMAGE_BYTES) { ctx.skippedImages += 1; continue; }
    ctx.imageBytes += buffer.length;
    ctx.imageCount += 1;
    const extent = /<wp:extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(runXml);
    images.push({ dataUrl: `data:${mime};base64,${buffer.toString("base64")}`, width: extent ? Math.round(Number(extent[1]) / 9525) : 0, height: extent ? Math.round(Number(extent[2]) / 9525) : 0 });
  }
  const textParts = [];
  for (const match of runXml.matchAll(/<w:(t|tab|br|cr|noBreakHyphen)\b[^>]*(?:\/>|>([\s\S]*?)<\/w:\1>)/g)) {
    if (match[1] === "t") textParts.push(decodeXmlText(match[2] ?? ""));
    else if (match[1] === "tab") textParts.push("\t");
    else if (match[1] === "noBreakHyphen") textParts.push("-");
    else textParts.push("\n");
  }
  const runs = [];
  const text = textParts.join("");
  if (text) runs.push({ text, ...props });
  for (const image of images) runs.push({ text: "", ...props, img: image });
  return runs;
}

async function parseDocxParagraph(pXml, ctx) {
  const pPr = /<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/.exec(pXml)?.[0] ?? "";
  const styleId = /<w:pStyle\b[^>]*w:val="([^"]+)"/.exec(pPr)?.[1] ?? "";
  let heading = ctx.styleHeadings.get(styleId) ?? 0;
  if (!heading) {
    const outline = /<w:outlineLvl\b[^>]*w:val="(\d+)"/.exec(pPr)?.[1];
    if (outline !== undefined) heading = Math.min(Number(outline) + 1, 3);
  }
  const jc = /<w:jc\b[^>]*w:val="([^"]+)"/.exec(pPr)?.[1];
  const align = jc === "center" ? "center" : jc === "right" ? "right" : jc === "left" ? "left" : jc === "both" || jc === "justify" ? "justify" : undefined;
  let list;
  const numPr = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr)?.[1];
  if (numPr) {
    const level = Math.min(Number(/<w:ilvl\b[^>]*w:val="(\d+)"/.exec(numPr)?.[1] ?? 0), 4);
    const numId = /<w:numId\b[^>]*w:val="(\d+)"/.exec(numPr)?.[1] ?? "";
    const abstractId = ctx.numbering.numToAbstract.get(numId);
    const format = String(ctx.numbering.formats.get(abstractId)?.[level] ?? "bullet").toLowerCase();
    list = { level, ordered: !["bullet", "none"].includes(format) };
  }
  const runs = [];
  for (const match of pXml.matchAll(/<w:r\b[^>]*(?:\/>|[\s\S]*?<\/w:r>)/g)) runs.push(...await parseDocxRun(match[0], ctx));
  if (!runs.length) runs.push({ text: "" });
  return { type: "paragraph", ...(heading ? { heading } : {}), ...(align ? { align } : {}), ...(list ? { list } : {}), runs };
}

async function parseDocxTable(tblXml, ctx) {
  const rows = [];
  for (const rowMatch of tblXml.matchAll(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g)) {
    const cells = [];
    for (const cellMatch of rowMatch[0].matchAll(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g)) {
      const runs = [];
      const paragraphs = [...cellMatch[0].matchAll(/<w:p\b[^>]*(?:\/>|[\s\S]*?<\/w:p>)/g)];
      for (let index = 0; index < paragraphs.length; index += 1) {
        if (index > 0) runs.push({ text: "\n" });
        const parsed = await parseDocxParagraph(paragraphs[index][0], ctx);
        runs.push(...parsed.runs);
      }
      if (!runs.length) runs.push({ text: "" });
      cells.push(runs);
    }
    if (cells.length) rows.push(cells);
  }
  return { type: "table", header: false, rows };
}

async function docx2rich(input) {
  const file = await ensureReadable(input.path);
  const buf = await fsp.readFile(file);
  const zip = await JSZip.loadAsync(buf);
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("不是有效的 docx（缺少 word/document.xml）");
  const stylesXml = await zip.file("word/styles.xml")?.async("string");
  const numberingXml = await zip.file("word/numbering.xml")?.async("string");
  const rels = docxRelationships(await zip.file("word/_rels/document.xml.rels")?.async("string"));
  const ctx = { zip, rels, styleHeadings: parseStylesHeadings(stylesXml), numbering: parseNumberingXml(numberingXml), imageCount: 0, imageBytes: 0, skippedImages: 0 };
  const bodyMatch = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(docXml);
  const bodyXml = bodyMatch?.[1] ?? docXml;
  const blocks = [];
  const tokenRe = /<w:p\b[^>]*(?:\/>|[\s\S]*?<\/w:p>)|<w:tbl\b[^>]*(?:\/>|[\s\S]*?<\/w:tbl>)/g;
  for (const match of bodyXml.matchAll(tokenRe)) {
    if (match[0].startsWith("<w:tbl")) blocks.push(await parseDocxTable(match[0], ctx));
    else if (match[0].trim() !== "<w:p/>") blocks.push(await parseDocxParagraph(match[0], ctx));
    else blocks.push({ type: "paragraph", runs: [{ text: "" }] });
    if (blocks.length >= 6000) break;
  }
  const headings = blocks.filter((block) => block.type === "paragraph" && block.heading).length;
  const lists = blocks.filter((block) => block.type === "paragraph" && block.list).length;
  const tables = blocks.filter((block) => block.type === "table").length;
  const chars = blocks.reduce((sum, block) => sum + (block.type === "table" ? block.rows.flat().reduce((cellSum, cell) => cellSum + cell.reduce((runSum, run) => runSum + String(run.text ?? "").length, 0), 0) : block.runs.reduce((runSum, run) => runSum + String(run.text ?? "").length, 0)), 0);
  const title = blocks.find((block) => block.type === "paragraph" && block.heading)?.runs?.map((run) => run.text).join("") ?? "";
  return { file, title, blocks, stats: { blocks: blocks.length, headings, lists, tables, images: ctx.imageCount, skippedImages: ctx.skippedImages, chars } };
}

const RICH_HEADING_SIZES = [20, 16, 13];
const RICH_HEADING_PX = [24, 19, 16];

function richRunProps(run, heading) {
  const parts = [];
  let bold = Boolean(run?.b);
  let size = Number(run?.size) || 0;
  if (heading) {
    bold = true;
    if (!size) size = RICH_HEADING_PX[Math.min(heading, 3) - 1];
  }
  const fonts = run?.font ? ` w:ascii="${escapeXmlText(run.font)}" w:hAnsi="${escapeXmlText(run.font)}" w:eastAsia="${escapeXmlText(run.font)}"` : ` w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="微软雅黑"`;
  parts.push(`<w:rFonts${fonts}/>`);
  if (bold) parts.push("<w:b/>");
  if (run?.i) parts.push("<w:i/>");
  if (run?.s) parts.push("<w:strike/>");
  if (run?.u) parts.push('<w:u w:val="single"/>');
  if (run?.color && /^#[0-9a-f]{6}$/i.test(run.color)) parts.push(`<w:color w:val="${run.color.slice(1).toUpperCase()}"/>`);
  if (run?.highlight && /^#[0-9a-f]{6}$/i.test(run.highlight)) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${run.highlight.slice(1).toUpperCase()}"/>`);
  if (size > 0) parts.push(`<w:sz w:val="${Math.round(size * 1.5)}"/>`, `<w:szCs w:val="${Math.round(size * 1.5)}"/>`);
  return `<w:rPr>${parts.join("")}</w:rPr>`;
}

function richImageXml(run, ctx) {
  const image = ctx.addImage(run?.img);
  if (!image) return "";
  const width = Math.max(24, Math.min(Number(run.img.width) || 480, 640));
  const height = Math.max(18, Math.round(width * (Number(run.img.height) || 320) / (Number(run.img.width) || 480)));
  const cx = Math.round(width * 9525);
  const cy = Math.round(height * 9525);
  const docPrId = ctx.docPrId++;
  ctx.warnings.imageDowngraded = false;
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${docPrId}" name="图片 ${docPrId}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="ruodao-image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${image.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function richRunXml(run, ctx, heading) {
  if (run?.img?.dataUrl) return richImageXml(run, ctx);
  const text = String(run?.text ?? "");
  if (!text) return "";
  const props = richRunProps(run, heading);
  const pieces = text.split(/(\n|\t)/).filter((piece) => piece !== "");
  const inner = pieces.map((piece) => {
    if (piece === "\n") return "<w:br/>";
    if (piece === "\t") return "<w:tab/>";
    return `<w:t xml:space="preserve">${escapeXmlText(piece)}</w:t>`;
  }).join("");
  return `<w:r>${props}${inner}</w:r>`;
}

function nextRelationshipId(relsXml) {
  let max = 0;
  for (const match of String(relsXml ?? "").matchAll(/Id="(?:rId)?(\d+)"/g)) max = Math.max(max, Number(match[1]));
  return max + 1;
}

function ensureContentDefaults(xml, extensions) {
  let next = String(xml ?? "");
  for (const ext of extensions) {
    const mime = IMAGE_MIME[ext];
    if (!mime || new RegExp(`Extension="${ext}"`, "i").test(next)) continue;
    next = next.replace(/<\/Types>/i, `<Default Extension="${ext}" ContentType="${mime}"/></Types>`);
  }
  return next;
}

function numberingLevelsXml(abstractId, ordered) {
  const levels = [];
  for (let index = 0; index < 5; index += 1) {
    levels.push(`<w:lvl w:ilvl="${index}"><w:start w:val="1"/><w:numFmt w:val="${ordered ? "decimal" : "bullet"}"/><w:lvlText w:val="${ordered ? `%${index + 1}.` : "•"}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${(index + 1) * 420}" w:hanging="420"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="微软雅黑" w:hint="default"/></w:rPr></w:lvl>`);
  }
  return `<w:abstractNum w:abstractNumId="${abstractId}"><w:multiLevelType w:val="hybridMultilevel"/>${levels.join("")}</w:abstractNum>`;
}

function ensureNumbering(xml) {
  const base = String(xml ?? "").trim() || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`;
  let maxAbstract = 0;
  let maxNum = 0;
  for (const match of base.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"/g)) maxAbstract = Math.max(maxAbstract, Number(match[1]));
  for (const match of base.matchAll(/<w:num\b[^>]*w:numId="(\d+)"/g)) maxNum = Math.max(maxNum, Number(match[1]));
  const bulletAbstract = maxAbstract + 1;
  const orderedAbstract = maxAbstract + 2;
  const bulletNum = maxNum + 1;
  const orderedNum = maxNum + 2;
  const addition = `${numberingLevelsXml(bulletAbstract, false)}${numberingLevelsXml(orderedAbstract, true)}<w:num w:numId="${bulletNum}"><w:abstractNumId w:val="${bulletAbstract}"/></w:num><w:num w:numId="${orderedNum}"><w:abstractNumId w:val="${orderedAbstract}"/></w:num>`;
  return { xml: base.replace(/<\/w:numbering>/i, `${addition}</w:numbering>`), bulletNum, orderedNum };
}

class RichDocxWriter {
  constructor({ zip, contentTypes, relsXml, nextRid }) {
    this.zip = zip;
    this.contentTypes = contentTypes;
    this.rels = [];
    this.nextRid = nextRid ?? nextRelationshipId(relsXml);
    this.mediaCounter = 0;
    this.docPrId = 1;
    this.imageCount = 0;
    this.imageBytes = 0;
    this.skippedImages = 0;
  }

  addImage(img) {
    const match = /^data:image\/([a-zA-Z0-9.+-]+);base64,([\s\S]+)$/.exec(String(img?.dataUrl ?? ""));
    if (!match) { this.skippedImages += 1; return null; }
    let ext = match[1].toLowerCase();
    if (ext === "jpeg") ext = "jpg";
    if (!IMAGE_MIME[ext]) { this.skippedImages += 1; return null; }
    const buffer = Buffer.from(match[2].replace(/\s+/g, ""), "base64");
    if (!buffer.length || buffer.length > 12 * 1024 * 1024) { this.skippedImages += 1; return null; }
    if (this.imageCount >= MAX_EMBED_IMAGES || this.imageBytes + buffer.length > 32 * 1024 * 1024) { this.skippedImages += 1; return null; }
    this.imageBytes += buffer.length;
    this.imageCount += 1;
    const name = `image-ruodao-${Date.now()}-${++this.mediaCounter}.${ext}`;
    this.zip.file(`word/media/${name}`, buffer);
    const rid = `rId${this.nextRid++}`;
    this.rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);
    this.contentTypes = ensureContentDefaults(this.contentTypes, [ext]);
    return { rid };
  }

  imageXml(run) {
    const image = this.addImage(run?.img);
    if (!image) return "";
    const rawWidth = Number(run.img.width) || 480;
    const width = Math.max(24, Math.min(rawWidth, 640));
    const height = Math.max(18, Math.round(width * ((Number(run.img.height) || 320) / rawWidth)));
    const cx = Math.round(width * 9525);
    const cy = Math.round(height * 9525);
    const docPrId = this.docPrId++;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${docPrId}" name="图片 ${docPrId}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="ruodao-image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${image.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }

  runXml(run, heading) {
    if (run?.img?.dataUrl) return this.imageXml(run);
    const text = String(run?.text ?? "");
    if (!text) return "";
    const props = richRunProps(run, heading);
    const pieces = text.split(/(\n|\t)/).filter((piece) => piece !== "");
    const inner = pieces.map((piece) => {
      if (piece === "\n") return "<w:br/>";
      if (piece === "\t") return "<w:tab/>";
      return `<w:t xml:space="preserve">${escapeXmlText(piece)}</w:t>`;
    }).join("");
    return `<w:r>${props}${inner}</w:r>`;
  }

  paragraphXml(block, numbering) {
    const heading = block?.heading ? Math.min(Math.max(Number(block.heading) || 1, 1), 3) : 0;
    const pPr = [];
    if (heading) pPr.push(`<w:pStyle w:val="Heading${heading}"/>`);
    if (["left", "center", "right", "justify"].includes(String(block?.align))) pPr.push(`<w:jc w:val="${block.align === "justify" ? "both" : block.align}"/>`);
    if (block?.list) {
      const numId = block.list.ordered ? numbering.orderedNum : numbering.bulletNum;
      const level = Math.min(Math.max(Number(block.list.level) || 0, 0), 4);
      pPr.push(`<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`);
    }
    const runs = (Array.isArray(block?.runs) ? block.runs : []).map((run) => this.runXml(run, heading)).join("");
    return `<w:p>${pPr.length ? `<w:pPr>${pPr.join("")}</w:pPr>` : ""}${runs}</w:p>`;
  }

  tableXml(block, numbering) {
    const rows = Array.isArray(block?.rows) ? block.rows.slice(0, 200) : [];
    if (!rows.length) return "";
    const border = `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="D7DCE5"/><w:left w:val="single" w:sz="4" w:space="0" w:color="D7DCE5"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="D7DCE5"/><w:right w:val="single" w:sz="4" w:space="0" w:color="D7DCE5"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="D7DCE5"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D7DCE5"/></w:tblBorders>`;
    const body = rows.map((row, rowIndex) => {
      const cells = (Array.isArray(row) ? row : []).map((cell) => {
        let runs = Array.isArray(cell) ? cell : [{ text: String(cell ?? "") }];
        if (block.header && rowIndex === 0) runs = runs.map((run) => ({ ...run, b: true }));
        const inner = runs.length ? runs.map((run) => this.runXml(run, 0)).join("") : "";
        return `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p>${inner}</w:p></w:tc>`;
      }).join("");
      return `<w:tr>${cells}</w:tr>`;
    }).join("");
    return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>${border}<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>${body}</w:tbl>`;
  }
}

function defaultSectPr() {
  return `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;
}

function docxStylesXml() {
  const headingStyle = (level, size, before) => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="${level - 1}"/><w:spacing w:before="${before}" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="${size}"/></w:rPr></w:style>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="微软雅黑" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>${headingStyle(1, 40, 280)}${headingStyle(2, 32, 240)}${headingStyle(3, 26, 200)}</w:styles>`;
}

async function docxSaveRich(input) {
  const blocks = (Array.isArray(input.blocks) ? input.blocks : []).slice(0, 6000);
  if (!blocks.length) throw new Error("blocks 不能为空：请至少保留一个段落");
  let zip;
  let docXml = "";
  let relsXml = "";
  let contentTypes = "";
  let numberingSource = "";
  let originalFile = "";
  if (input.path) {
    originalFile = await ensureReadable(input.path);
    zip = await JSZip.loadAsync(await fsp.readFile(originalFile));
    docXml = await zip.file("word/document.xml")?.async("string") ?? "";
    if (!docXml) throw new Error("不是有效的 docx（缺少 word/document.xml）");
    relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string") ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    contentTypes = await zip.file("[Content_Types].xml")?.async("string") ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`;
    numberingSource = await zip.file("word/numbering.xml")?.async("string") ?? "";
  } else {
    zip = new JSZip();
    relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`;
    contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXmlText(input.title ?? "若道文档")}</dc:title><dc:creator>Ruodao office-suite</dc:creator></cp:coreProperties>`);
    zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Ruodao office-suite</Application></Properties>`);
    zip.file("word/styles.xml", docxStylesXml());
    docXml = "";
  }
  const numbering = ensureNumbering(numberingSource);
  const writer = new RichDocxWriter({ zip, contentTypes, relsXml, nextRid: input.path ? undefined : 3 });
  const body = blocks.map((block) => block?.type === "table" ? writer.tableXml(block, numbering) : writer.paragraphXml(block, numbering)).join("");
  if (input.path) {
    const sectMatch = /(<w:sectPr\b[\s\S]*?<\/w:sectPr>)\s*<\/w:body>/.exec(docXml);
    const sectPr = sectMatch?.[1] ?? defaultSectPr();
    const nextDoc = docXml.replace(/<w:body\b[^>]*>[\s\S]*<\/w:body>/, `<w:body>${body}${sectPr}</w:body>`);
    zip.file("word/document.xml", nextDoc);
  } else {
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${DOCX_XML_NS}><w:body>${body}${defaultSectPr()}</w:body></w:document>`);
  }
  zip.file("word/numbering.xml", numbering.xml);
  let nextRels = String(relsXml).replace(/<\/Relationships>/i, `${writer.rels.join("")}</Relationships>`);
  if (!input.path && !/relationships\/numbering/.test(nextRels)) nextRels = nextRels.replace(/<\/Relationships>/i, `<Relationship Id="rId${writer.nextRid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`);
  if (input.path && !numberingSource) {
    nextRels = nextRels.replace(/<\/Relationships>/i, `<Relationship Id="rId${writer.nextRid++}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`);
    if (!/PartName="\/word\/numbering.xml"/.test(writer.contentTypes)) writer.contentTypes = writer.contentTypes.replace(/<\/Types>/i, `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`);
  }
  zip.file("word/_rels/document.xml.rels", nextRels);
  zip.file("[Content_Types].xml", writer.contentTypes);
  const base = originalFile ? path.basename(originalFile, path.extname(originalFile)) : `ruodao-${Date.now()}`;
  const output = path.resolve(String(input.output ?? (input.inPlace && originalFile ? originalFile : path.join(originalFile ? path.dirname(originalFile) : process.cwd(), `${base}-edited.docx`))));
  await fsp.mkdir(path.dirname(output), { recursive: true });
  await fsp.writeFile(output, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return { file: originalFile || null, output, stats: { blocks: blocks.length, images: writer.imageCount, skippedImages: writer.skippedImages, imageBytes: writer.imageBytes } };
}

/* ---------------- xlsx / csv ---------------- */

async function readWorkbook(inputPath) {
  const file = await ensureReadable(inputPath);
  const buf = await fsp.readFile(file);
  const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  return { file, wb };
}

async function xlsxGrid(input) {
  const { file, wb } = await readWorkbook(input.path);
  const sheets = wb.SheetNames.map((name) => {
    const sheet = wb.Sheets[name] ?? {};
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    const formulas = {};
    for (const [address, cell] of Object.entries(sheet)) {
      if (address.startsWith("!") || !cell || typeof cell !== "object" || !cell.f) continue;
      formulas[address] = `=${cell.f}`;
    }
    return {
      name,
      rows: matrix.map((row) => (row ?? []).map((cell) => (cell === null || cell === undefined ? "" : String(cell)))),
      formulas,
    };
  });
  return { file, sheets };
}

// 公式缓存值：SheetJS 读回时缺 <v> 会直接丢掉公式单元格，所以写入前先按常见语义求值。
function formulaColumnIndex(letters) {
  return letters.toUpperCase().split("").reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
}

function evaluateFormulaGrid(rows, formulas) {
  const cache = new Map();
  const visiting = new Set();
  const byCell = new Map(formulas.map((formula) => [`${formula.r}:${formula.c}`, formula.f]));
  const roundValue = (value) => (Number.isFinite(value) ? Math.round(value * 1e8) / 1e8 : value);
  const isNumeric = (value) => String(value ?? "").trim() !== "" && Number.isFinite(Number(value));
  const toNumber = (value) => {
    if (typeof value === "number") return value;
    const text = String(value ?? "").trim();
    if (text === "") return 0;
    const parsed = Number(text);
    if (Number.isFinite(parsed)) return parsed;
    if (text.startsWith("#")) throw new Error(text);
    return 0;
  };

  const tokenize = (input) => {
    const tokens = [];
    const text = String(input ?? "").replace(/\s+/g, "");
    let index = 0;
    while (index < text.length) {
      const char = text[index];
      if (/[0-9.]/.test(char)) {
        let number = char;
        index += 1;
        while (index < text.length && /[0-9.]/.test(text[index])) { number += text[index]; index += 1; }
        tokens.push(number);
        continue;
      }
      if (char === '"') {
        let value = "";
        index += 1;
        while (index < text.length && text[index] !== '"') { value += text[index]; index += 1; }
        index += 1;
        tokens.push(`"${value}`);
        continue;
      }
      if (/[A-Za-z_$]/.test(char)) {
        let word = char;
        index += 1;
        while (index < text.length && /[A-Za-z0-9_.$:]/.test(text[index])) { word += text[index]; index += 1; }
        tokens.push(word.toUpperCase());
        continue;
      }
      if ("+-*/^%(),".includes(char)) { tokens.push(char); index += 1; continue; }
      throw new Error("#ERR!");
    }
    return tokens;
  };

  const valueAt = (r, c) => {
    const key = `${r}:${c}`;
    if (cache.has(key)) return cache.get(key);
    const formula = byCell.get(key);
    const raw = rows[r]?.[c];
    if (formula === undefined) return raw === undefined || raw === null ? "" : raw;
    if (visiting.has(key)) throw new Error("#CIRC!");
    visiting.add(key);
    let value;
    try { value = evaluate(formula); }
    catch (error) { value = /^#/.test(String(error?.message ?? "")) ? String(error.message) : "#ERR!"; }
    finally { visiting.delete(key); }
    cache.set(key, value);
    return value;
  };

  const cellValue = (address) => {
    const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(address.trim());
    if (!match) throw new Error("#REF!");
    const r = Number(match[2]) - 1;
    const c = formulaColumnIndex(match[1]);
    if (r < 0 || c < 0) throw new Error("#REF!");
    return valueAt(r, c);
  };

  const rangeNumbers = (token) => {
    const match = /^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/.exec(token.trim());
    if (!match) throw new Error("#REF!");
    const c1 = formulaColumnIndex(match[1]);
    const r1 = Number(match[2]) - 1;
    const c2 = formulaColumnIndex(match[3]);
    const r2 = Number(match[4]) - 1;
    const values = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r += 1) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c += 1) {
        const value = valueAt(r, c);
        if (typeof value === "number") values.push(value);
        else if (isNumeric(value)) values.push(Number(value));
      }
    }
    return values;
  };

  const evaluate = (source) => {
    const tokens = tokenize(source);
    let position = 0;
    const peek = () => tokens[position];
    const next = () => tokens[position++];
    const expect = (token) => { if (next() !== token) throw new Error("#ERR!"); };

    const parseExpression = () => {
      let left = parseTerm();
      while (peek() === "+" || peek() === "-") {
        const operator = next();
        const right = parseTerm();
        const a = toNumber(left);
        const b = toNumber(right);
        left = operator === "+" ? a + b : a - b;
      }
      return left;
    };
    const parseTerm = () => {
      let left = parseUnary();
      while (peek() === "*" || peek() === "/") {
        const operator = next();
        const right = parseUnary();
        const a = toNumber(left);
        const b = toNumber(right);
        if (operator === "/" && b === 0) throw new Error("#DIV/0!");
        left = operator === "*" ? a * b : a / b;
      }
      return left;
    };
    const parseUnary = () => {
      if (peek() === "-") { next(); return -toNumber(parseUnary()); }
      if (peek() === "+") { next(); return parseUnary(); }
      return parsePower();
    };
    const parsePower = () => {
      const base = parsePrimary();
      if (peek() === "^") { next(); return toNumber(base) ** toNumber(parseUnary()); }
      return base;
    };
    const parsePrimary = () => {
      const token = next();
      if (token === undefined) throw new Error("#ERR!");
      if (token === "(") { const value = parseExpression(); expect(")"); return value; }
      if (token.startsWith('"')) return token.slice(1);
      if (/^[0-9.]+$/.test(token)) {
        if (peek() === "%") { next(); return Number(token) / 100; }
        const value = Number(token);
        if (!Number.isFinite(value)) throw new Error("#NUM!");
        return value;
      }
      if (/^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/.test(token)) {
        return rangeNumbers(token).reduce((sum, value) => sum + value, 0);
      }
      if (/^\$?[A-Z]+\$?\d+$/.test(token)) return cellValue(token);
      if (peek() === "(") return callFunction(token);
      if (token === "TRUE") return 1;
      if (token === "FALSE") return 0;
      throw new Error("#NAME?");
    };
    const callFunction = (name) => {
      expect("(");
      const args = [];
      if (peek() !== ")") {
        do {
          const token = peek();
          if (token && /^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/.test(token)) { next(); args.push(rangeNumbers(token)); }
          else args.push(parseExpression());
        } while (peek() === "," && next());
      }
      expect(")");
      const flat = args.flatMap((arg) => (Array.isArray(arg) ? arg : [arg]));
      const numbers = flat.map(toNumber).filter((value) => Number.isFinite(value));
      const sum = () => numbers.reduce((total, value) => total + value, 0);
      switch (name) {
        case "SUM": return sum();
        case "AVERAGE": return numbers.length ? sum() / numbers.length : 0;
        case "COUNT": return numbers.length;
        case "COUNTA": return flat.filter((value) => String(value ?? "").trim() !== "").length;
        case "MIN": return numbers.length ? Math.min(...numbers) : 0;
        case "MAX": return numbers.length ? Math.max(...numbers) : 0;
        case "ABS": return Math.abs(numbers[0] ?? 0);
        case "ROUND": { const digits = Math.max(0, Math.min(8, Number(numbers[1] ?? 0))); return Number((numbers[0] ?? 0).toFixed(digits)); }
        case "INT": return Math.floor(numbers[0] ?? 0);
        case "SQRT": return Math.sqrt(Math.max(0, numbers[0] ?? 0));
        case "IF": {
          const [condition, yes, no] = flat;
          const truthy = typeof condition === "number" ? condition !== 0 : Boolean(String(condition ?? "").trim());
          return truthy ? yes ?? "" : no ?? "";
        }
        default: throw new Error("#NAME?");
      }
    };

    const result = parseExpression();
    if (position < tokens.length) throw new Error("#ERR!");
    return typeof result === "number" ? roundValue(result) : result;
  };

  const values = new Map();
  for (const formula of formulas) values.set(`${formula.r}:${formula.c}`, valueAt(formula.r, formula.c));
  return values;
}

async function xlsxCreate(input) {
  const sheetInputs = Array.isArray(input.sheets) ? input.sheets : [];
  if (!sheetInputs.length) throw new Error("sheets 不能为空：{ name, rows: [[...]] }");
  const wb = XLSX.utils.book_new();
  const usedNames = new Set();
  sheetInputs.slice(0, 20).forEach((sheet, index) => {
    let name = String(sheet?.name ?? `Sheet${index + 1}`).replace(/[\\/?*\[\]:]/g, "_").slice(0, 31) || `Sheet${index + 1}`;
    while (usedNames.has(name)) name = `${name.slice(0, 28)}_${index}`;
    usedNames.add(name);
    const rows = Array.isArray(sheet?.rows) ? sheet.rows : [];
    const formulas = [];
    const clean = rows.slice(0, 20000).map((row, rowIndex) => (Array.isArray(row) ? row.slice(0, 200).map((cell, colIndex) => {
      const text = String(cell ?? "");
      if (text === "") return "";
      if (text.startsWith("=") && text.length > 1) { formulas.push({ r: rowIndex, c: colIndex, f: text.slice(1) }); return ""; }
      const num = Number(text);
      return Number.isFinite(num) && text.trim() !== "" && !/^0\d/.test(text) ? num : text;
    }) : []));
    const ws = XLSX.utils.aoa_to_sheet(clean.length ? clean : [[""]]);
    const cached = evaluateFormulaGrid(clean, formulas);
    for (const formula of formulas) {
      const address = XLSX.utils.encode_cell({ r: formula.r, c: formula.c });
      const value = cached.get(`${formula.r}:${formula.c}`);
      if (typeof value === "number" && Number.isFinite(value)) ws[address] = { t: "n", f: formula.f, v: value };
      else if (typeof value === "string" && value && !value.startsWith("#")) ws[address] = { t: "s", f: formula.f, v: value };
      // 无法求值的公式保留公式本身，缓存留空让 Excel/WPS 打开时自行重算。
      else ws[address] = { t: "s", f: formula.f, v: "" };
    }
    const widths = (clean[0] ?? []).map((_, col) => {
      const longest = Math.max(...clean.slice(0, 200).map((row) => String(row[col] ?? "").length), 4);
      return { wch: Math.min(Math.max(longest + 2, 6), 48) };
    });
    ws["!cols"] = widths;
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  const output = path.resolve(String(input.output ?? path.join(process.cwd(), `ruodao-${Date.now()}.xlsx`)));
  await fsp.mkdir(path.dirname(output), { recursive: true });
  await fsp.writeFile(output, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return { output, sheets: wb.SheetNames };
}

async function xlsx2json(input) {
  const { file, wb } = await readWorkbook(input.path);
  const limit = Math.max(1, Math.min(Number(input.rowLimit ?? 500) || 500, 5000));
  const sheets = wb.SheetNames.map((name) => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null });
    return { name, totalRows: rows.length, rows: rows.slice(0, limit) };
  });
  return { file, sheets };
}

async function xlsx2csv(input) {
  const { file, wb } = await readWorkbook(input.path);
  const sheetName = input.sheet && wb.SheetNames.includes(input.sheet) ? input.sheet : wb.SheetNames[0];
  const csv = XLSX.utils.sheet_to_csv(wb.Sheets[sheetName]);
  const out = path.resolve(String(input.output ?? path.join(path.dirname(file), `${path.basename(file, path.extname(file))}.csv`)));
  await fsp.writeFile(out, "\uFEFF" + csv, "utf8");
  return { file, sheet: sheetName, output: out, rows: csv.split(/\r?\n/).filter(Boolean).length };
}

async function xlsxEdit(input) {
  const { file, wb } = await readWorkbook(input.path);
  const sheetName = input.sheet && wb.SheetNames.includes(input.sheet) ? input.sheet : wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const edits = Array.isArray(input.edits) ? input.edits : [];
  if (!edits.length) throw new Error("edits 不能为空：[{ cell: 'A1', value: ... }]");
  for (const { cell, value } of edits) {
    if (!/^[A-Z]{1,3}[1-9]\d*$/.test(String(cell ?? ""))) throw new Error(`非法单元格地址：${cell}`);
    sheet[cell] = { t: typeof value === "number" ? "n" : "s", v: value };
  }
  const out = path.resolve(String(input.output ?? path.join(path.dirname(file), `${path.basename(file, path.extname(file))}-edited${path.extname(file).toLowerCase() === ".csv" ? ".xlsx" : path.extname(file)}`)));
  await fsp.mkdir(path.dirname(out), { recursive: true });
  await fsp.writeFile(out, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return { file, sheet: sheetName, output: out, applied: edits.length };
}

async function csv2xlsx(input) {
  const file = await ensureReadable(input.path);
  const text = await fsp.readFile(file, "utf8");
  const wb = XLSX.read(text.replace(/^\uFEFF/, ""), { type: "string", raw: false });
  const out = path.resolve(String(input.output ?? path.join(path.dirname(file), `${path.basename(file, path.extname(file))}.xlsx`)));
  await fsp.writeFile(out, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  return { file, output: out, sheets: wb.SheetNames };
}

/* ---------------- pptx ---------------- */

async function pptx2json(input) {
  const file = await ensureReadable(input.path);
  const buf = await fsp.readFile(file);
  const zip = await JSZip.loadAsync(buf);
  const slideNames = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => (Number(/(\d+)\.xml$/.exec(a)?.[1]) - Number(/(\d+)\.xml$/.exec(b)?.[1])));
  if (!slideNames.length) throw new Error("不是有效的 pptx（缺少 ppt/slides/）");
  const stripTags = (xml) => xml.replace(/<[^>]+>/g, "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const slides = [];
  for (const name of slideNames) {
    const xml = await zip.file(name)?.async("string");
    if (!xml) continue;
    const paragraphs = [...xml.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)].map((m) => stripTags(m[0]).trim()).filter(Boolean);
    const tableCount = (xml.match(/<a:tbl>/g) ?? []).length;
    const imageCount = (xml.match(/<p:pic\b/g) ?? []).length;
    slides.push({ index: slides.length + 1, title: paragraphs[0] ?? "", paragraphs, tableCount, imageCount });
  }
  const notesCount = Object.keys(zip.files).filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n)).length;
  return { file, stats: { slides: slides.length, notesCount }, slides };
}

// 大纲 → pptx：shape = { title, bullets[], notes?, layout? }，支持主题配色与版式
const PPTX_ML = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
const PPTX_W = 12192000;
const PPTX_H = 6858000;

function pptxHex(value, fallback) {
  const text = String(value ?? "").trim().replace(/^#/, "");
  return /^[0-9a-fA-F]{6}$/.test(text) ? text.toUpperCase() : fallback;
}

function normalizePptxTheme(theme) {
  const source = theme && typeof theme === "object" ? theme : {};
  return {
    bg: pptxHex(source.bg, "FFFFFF"),
    bgAlt: pptxHex(source.bgAlt, "F6F7F9"),
    primary: pptxHex(source.primary, "EA580C"),
    accent: pptxHex(source.accent, "2563EB"),
    text: pptxHex(source.text, "111827"),
    muted: pptxHex(source.muted, "6B7280"),
    font: String(source.font ?? "微软雅黑").slice(0, 40) || "微软雅黑",
    decor: ["circle", "bar", "grid"].includes(source.decor) ? source.decor : "bar",
    gradient: Array.isArray(source.gradient) && source.gradient.length >= 2 ? source.gradient.map((color) => pptxHex(color, "FFFFFF")) : null,
  };
}

function pptxSolid(color, alpha = 100000) {
  return `<a:solidFill><a:srgbClr val="${color}"><a:alpha val="${Math.max(0, Math.min(100000, alpha))}"/></a:srgbClr></a:solidFill>`;
}

function pptxRun(text, theme, { color = theme.text, size = 18, bold = false, alpha = 100000 } = {}) {
  return `<a:r><a:rPr lang="zh-CN" sz="${Math.round(size * 100)}"${bold ? ' b="1"' : ""} dirty="0">${pptxSolid(color, alpha)}<a:latin typeface="${PPTX_ML(theme.font)}"/><a:ea typeface="${PPTX_ML(theme.font)}"/></a:rPr><a:t>${PPTX_ML(text)}</a:t></a:r>`;
}

function pptxParagraph(runs, { align = "l", marL = 0, indent = 0, bullet = "", bulletColor = "EA580C", spaceAfter = 0, line = 100000 } = {}) {
  const bulletXml = bullet
    ? `<a:buClr><a:srgbClr val="${bulletColor}"/></a:buClr><a:buSzPct val="110000"/><a:buFont typeface="Arial"/><a:buChar char="${PPTX_ML(bullet)}"/>`
    : `<a:buNone/>`;
  return `<a:p><a:pPr algn="${align}" marL="${marL}" indent="${indent}"><a:lnSpc><a:spcPct val="${line}"/></a:lnSpc>${spaceAfter ? `<a:spcAft><a:spcPts val="${spaceAfter}"/></a:spcAft>` : ""}${bulletXml}</a:pPr>${runs.join("")}</a:p>`;
}

function pptxTextBox({ id, name, x, y, w, h, children, anchor = "t", inset = 0 }) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${PPTX_ML(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${anchor}" lIns="${inset}" tIns="${inset}" rIns="${inset}" bIns="${inset}"/><a:lstStyle/>${children}</p:txBody></p:sp>`;
}

function pptxShape({ id, name, x, y, w, h, fill, alpha = 100000, geom = "rect", lineColor = "", lineWidth = 0 }) {
  const fillXml = fill ? pptxSolid(fill, alpha) : "<a:noFill/>";
  const lineXml = lineColor ? `<a:ln w="${Math.round(lineWidth)}"><a:solidFill><a:srgbClr val="${lineColor}"/></a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${PPTX_ML(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr></p:sp>`;
}

function pptxBackground(theme, layout) {
  if (layout === "cover" && theme.gradient) {
    return `<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="${theme.gradient[0]}"/></a:gs><a:gs pos="100000"><a:srgbClr val="${theme.gradient[1]}"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>`;
  }
  return `<p:bg><p:bgPr>${pptxSolid(theme.bg, 100000)}<a:effectLst/></p:bgPr></p:bg>`;
}

function pptxDecorShapes(theme, idStart) {
  const shapes = [];
  let id = idStart;
  if (theme.decor === "circle") {
    shapes.push(pptxShape({ id: id++, name: "装饰圆", x: PPTX_W - 3350000, y: -1300000, w: 4900000, h: 4900000, fill: theme.primary, alpha: 12000, geom: "ellipse" }));
    shapes.push(pptxShape({ id: id++, name: "装饰环", x: PPTX_W - 2300000, y: 3800000, w: 2200000, h: 2200000, fill: "", lineColor: theme.accent, lineWidth: 19050, geom: "ellipse" }));
  } else if (theme.decor === "grid") {
    for (let index = 0; index < 4; index += 1) shapes.push(pptxShape({ id: id++, name: `网格线${index + 1}`, x: PPTX_W - 1200000 - index * 560000, y: 0, w: 16000, h: PPTX_H, fill: theme.primary, alpha: 9000 }));
    shapes.push(pptxShape({ id: id++, name: "装饰条", x: 0, y: 0, w: 160000, h: PPTX_H, fill: theme.primary, alpha: 80000 }));
  } else {
    shapes.push(pptxShape({ id: id++, name: "装饰条", x: 0, y: 0, w: 150000, h: PPTX_H, fill: theme.primary, alpha: 100000 }));
    shapes.push(pptxShape({ id: id++, name: "底部分隔线", x: 838200, y: PPTX_H - 520000, w: PPTX_W - 1676400, h: 12000, fill: theme.accent, alpha: 80000 }));
  }
  return { xml: shapes.join(""), nextId: id };
}

function pptxBulletLines(bullets) {
  return (Array.isArray(bullets) ? bullets : []).slice(0, 12).map((line) => {
    if (typeof line === "string") return { text: line, level: 0 };
    const text = String(line?.text ?? "");
    const level = Math.min(Math.max(Number(line?.level) || 0, 0), 4);
    return { text, level };
  }).filter((line) => line.text.trim());
}

function pptxSlideXml(slide, slideId, themeInput, total) {
  const theme = normalizePptxTheme(themeInput);
  const layout = ["cover", "section", "content", "end"].includes(slide?.layout) ? slide.layout : "content";
  const title = String(slide?.title ?? "").trim() || (layout === "end" ? "谢谢观看" : "未命名页面");
  const lines = pptxBulletLines(slide?.bullets);
  let id = 2;
  const parts = [];
  const decor = pptxDecorShapes(theme, id);
  parts.push(decor.xml);
  id = decor.nextId;

  const footer = layout === "cover" || layout === "end" ? "" : `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="页脚"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="${PPTX_H - 480000}"/><a:ext cx="7000000" cy="260000"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${pptxParagraph([pptxRun("若道办公套件 · 演示文稿", theme, { color: theme.muted, size: 9 })], { align: "l" })}</p:txBody></p:sp>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="页码"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${PPTX_W - 1500000}" y="${PPTX_H - 480000}"/><a:ext cx="660000" cy="260000"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${pptxParagraph([pptxRun(`${slideId} / ${total}`, theme, { color: theme.muted, size: 9 })], { align: "r" })}</p:txBody></p:sp>`;

  if (layout === "cover") {
    const subtitle = lines.length ? lines.map((line) => line.text).slice(0, 3).join(" · ") : "汇报人 · 日期";
    parts.push(pptxTextBox({ id: id++, name: "封面小标题", x: 838200, y: 1260000, w: 9000000, h: 320000, children: pptxParagraph([pptxRun("RUODAO PRESENTATION", theme, { color: theme.primary, size: 11, bold: true })], { align: "l" }) }));
    parts.push(pptxTextBox({ id: id++, name: "封面标题", x: 838200, y: 1980000, w: 9700000, h: 1900000, children: pptxParagraph([pptxRun(title, theme, { color: theme.text, size: title.length > 20 ? 40 : 50, bold: true })], { align: "l", line: 95000 }) }));
    parts.push(pptxTextBox({ id: id++, name: "封面副标题", x: 838200, y: 4080000, w: 9400000, h: 700000, children: pptxParagraph([pptxRun(subtitle, theme, { color: theme.muted, size: 17 })], { align: "l" }) }));
    parts.push(pptxShape({ id: id++, name: "封面强调线", x: 838200, y: 3780000, w: 1080000, h: 54000, fill: theme.primary }));
  } else if (layout === "section") {
    const number = String(slideId).padStart(2, "0");
    parts.push(pptxTextBox({ id: id++, name: "章节序号", x: 838200, y: 1250000, w: 3000000, h: 1600000, children: pptxParagraph([pptxRun(number, theme, { color: theme.primary, size: 88, bold: true, alpha: 26000 })], { align: "l" }) }));
    parts.push(pptxTextBox({ id: id++, name: "章节标题", x: 838200, y: 3100000, w: 10000000, h: 1100000, children: pptxParagraph([pptxRun(title, theme, { color: theme.text, size: 40, bold: true })], { align: "l" }) }));
    parts.push(pptxShape({ id: id++, name: "章节强调线", x: 838200, y: 4380000, w: 1320000, h: 48000, fill: theme.accent }));
    if (lines.length) parts.push(pptxTextBox({ id: id++, name: "章节说明", x: 838200, y: 4560000, w: 9800000, h: 700000, children: lines.slice(0, 2).map((line) => pptxParagraph([pptxRun(line.text, theme, { color: theme.muted, size: 14 })], { align: "l" })).join("") }));
  } else if (layout === "end") {
    const subtitle = lines.length ? lines.map((line) => line.text).slice(0, 2).join(" · ") : "期待与您继续交流";
    parts.push(pptxTextBox({ id: id++, name: "结束标题", x: 838200, y: 2450000, w: 10500000, h: 1300000, children: pptxParagraph([pptxRun(title || "谢谢观看", theme, { color: theme.text, size: 48, bold: true })], { align: "ctr" }) }));
    parts.push(pptxTextBox({ id: id++, name: "结束副标题", x: 838200, y: 3900000, w: 10500000, h: 600000, children: pptxParagraph([pptxRun(subtitle, theme, { color: theme.muted, size: 16 })], { align: "ctr" }) }));
    parts.push(pptxShape({ id: id++, name: "结束强调线", x: (PPTX_W - 1000000) / 2, y: 3720000, w: 1000000, h: 46000, fill: theme.primary }));
  } else {
    parts.push(pptxTextBox({ id: id++, name: "页标题", x: 838200, y: 640000, w: 10500000, h: 820000, children: pptxParagraph([pptxRun(title, theme, { color: theme.text, size: title.length > 26 ? 24 : 29, bold: true })], { align: "l" }) }));
    parts.push(pptxShape({ id: id++, name: "标题强调线", x: 838200, y: 1530000, w: 880000, h: 50000, fill: theme.primary }));
    const bulletParagraphs = lines.map((line) => pptxParagraph(
      [pptxRun(line.text, theme, { color: line.level === 0 ? theme.text : theme.muted, size: line.level === 0 ? 17 : 14.5 })],
      { align: "l", marL: 838200 + line.level * 380000, indent: -260000, bullet: line.level === 0 ? "•" : "–", bulletColor: line.level === 0 ? theme.primary : theme.muted, spaceAfter: 900, line: 106000 },
    )).join("");
    parts.push(pptxTextBox({ id: id++, name: "正文", x: 300000, y: 1750000, w: PPTX_W - 1000000, h: 4300000, children: bulletParagraphs || pptxParagraph([pptxRun("在左侧编辑页标题与要点，保存后即可在 WPS / Office 中继续编辑。", theme, { color: theme.muted, size: 15 })], { align: "l" }) }));
    parts.push(footer);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>${pptxBackground(theme, layout)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${parts.join("")}</p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function pptxNotesXml(text) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" dirty="0"/><a:t>${PPTX_ML(text)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`;

function pptxMasterXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`;
}

function pptxLayoutXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj" preserve="1"><p:cSld name="标题和内容"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function pptxThemeXml(input) {
  const theme = normalizePptxTheme(input);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Ruodao Office"><a:themeElements><a:clrScheme name="Ruodao"><a:dk1><a:srgbClr val="${theme.text}"/></a:dk1><a:lt1><a:srgbClr val="${theme.bg}"/></a:lt1><a:dk2><a:srgbClr val="${theme.muted}"/></a:dk2><a:lt2><a:srgbClr val="${theme.bgAlt}"/></a:lt2><a:accent1><a:srgbClr val="${theme.primary}"/></a:accent1><a:accent2><a:srgbClr val="${theme.accent}"/></a:accent2><a:accent3><a:srgbClr val="059669"/></a:accent3><a:accent4><a:srgbClr val="D97706"/></a:accent4><a:accent5><a:srgbClr val="DC2626"/></a:accent5><a:accent6><a:srgbClr val="0891B2"/></a:accent6><a:hlink><a:srgbClr val="${theme.accent}"/></a:hlink><a:folHlink><a:srgbClr val="${theme.primary}"/></a:folHlink></a:clrScheme><a:fontScheme name="Ruodao"><a:majorFont><a:latin typeface="${PPTX_ML(theme.font)}"/><a:ea typeface="${PPTX_ML(theme.font)}"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${PPTX_ML(theme.font)}"/><a:ea typeface="${PPTX_ML(theme.font)}"/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Ruodao"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

async function pptxBuild(input) {
  const shapes = Array.isArray(input.slides) ? input.slides : [];
  if (!shapes.length) throw new Error("slides 不能为空：[{ title, bullets: ['...'] , notes? }]");
  if (shapes.length > 100) throw new Error("单次最多生成 100 页");
  const theme = normalizePptxTheme(input.theme);
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${PPTX_ML(input.title ?? "Ruodao 演示文稿")}</dc:title><dc:creator>Ruodao office-suite</dc:creator></cp:coreProperties>`);
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Ruodao office-suite</Application></Properties>`);
  const slideIds = shapes.map((_, i) => 256 + i).join("");
  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${shapes.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${PPTX_W}" cy="${PPTX_H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>${shapes.map((_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join("")}<Relationship Id="rId${shapes.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`);
  zip.file("ppt/slideMasters/slideMaster1.xml", pptxMasterXml());
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>`);
  zip.file("ppt/slideLayouts/slideLayout1.xml", pptxLayoutXml());
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file("ppt/theme/theme1.xml", pptxThemeXml(theme));
  shapes.forEach((shape, i) => {
    zip.file(`ppt/slides/slide${i + 1}.xml`, pptxSlideXml(shape, i + 1, theme, shapes.length));
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  });
  const notesSlides = shapes.filter((s) => s.notes && String(s.notes).trim());
  if (notesSlides.length) {
    // 仅当至少一页带备注时写入备注部件，避免空引用
    shapes.forEach((shape, i) => {
      if (!String(shape.notes ?? "").trim()) return;
      zip.file(`ppt/notesSlides/notesSlide${i + 1}.xml`, pptxNotesXml(shape.notes));
      zip.file(`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${i + 1}.xml"/></Relationships>`);
      zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${i + 1}.xml"/></Relationships>`);
    });
    // 备注部件需要补进 ContentTypes
    let ct = await zip.file("[Content_Types].xml").async("string");
    ct = ct.replace("</Types>", `${shapes.map((shape, i) => (String(shape.notes ?? "").trim() ? `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` : "")).join("")}</Types>`);
    zip.file("[Content_Types].xml", ct);
  }
  const output = path.resolve(String(input.output ?? path.join(process.cwd(), `ruodao-${Date.now()}.pptx`)));
  await fsp.mkdir(path.dirname(output), { recursive: true });
  await fsp.writeFile(output, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  return { output, slides: shapes.length };
}

/* ---------------- pdf（V1 仅探测） ---------------- */

async function pdfNote(input) {
  const file = await ensureReadable(input.path);
  const buf = await fsp.readFile(file);
  const head = buf.subarray(0, 8).toString("latin1");
  if (!head.startsWith("%PDF")) throw new Error("不是有效的 PDF 文件");
  const matches = [...buf.toString("latin1").matchAll(/\/Type\s*\/Page[^s]/g)];
  const pages = matches.length || (() => { const m = /\/Count\s+(\d+)/.exec(buf.toString("latin1")); return m ? Number(m[1]) : 0; })();
  return { file, pages, note: "页数与元信息探测；查看、缩略图、缩放旋转、导出 PNG/文本/Word、拆分/合并/压缩由工作台在本机完成" };
}

/* ---------------- AI 处理（复用若道模型库凭据，由宿主经 modelBinding 注入） ---------------- */

const AI_TASKS = {
  summarize: "请对以下文档内容输出简洁的中文摘要，先用一句话概括主旨，再列出 3-8 条要点。",
  rewrite: "请改写以下文本，使其更通顺专业，保持原意与信息量，直接输出改写结果。",
  translate: "请将以下文本翻译为中文（若原文已是中文则翻译为英文），直接输出译文。",
  points: "请从以下内容提取要点清单（每条一行，不超过 20 字），保留顺序。",
  continue: "请根据以下内容自然续写一段（150-300 字），风格保持一致，直接输出续写内容。",
  free: "",
};

async function aiProcess(input) {
  const baseUrl = env("RUODAO_LLM_BASE_URL").replace(/\/+$/, "");
  const apiKey = env("RUODAO_LLM_API_KEY");
  const model = env("RUODAO_LLM_MODEL");
  if (!baseUrl || !model) throw new Error("未配置 AI 模型：请先在模型管理中添加模型，并在工作台选择文本模型");
  const preset = AI_TASKS[input.task ?? "summarize"] ?? "";
  const instruction = String(input.instruction ?? "").trim();
  const text = String(input.text ?? "").trim();
  if (!text) throw new Error("text 不能为空");
  if (text.length > 60000) throw new Error("文本过长（>6 万字符），请分段处理");
  const system = "你是专业的办公文档助手。基于用户提供的文档内容完成任务，不要编造内容。";
  const user = [instruction || preset, "——以下为文档内容——", text].filter(Boolean).join("\n\n");
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(new Error("AI 请求超时")), Number(input.timeoutMs ?? 120000) || 120000);
  let response;
  try {
    response = await fetch(chatEndpoint(baseUrl), {
      method: "POST",
      headers: { "content-type": "application/json", ...(apiKey ? { authorization: `Bearer ${apiKey}` } : {}) },
      body: JSON.stringify({ model, temperature: 0.4, messages: [{ role: "system", content: system }, { role: "user", content: user }] }),
      signal: ac.signal,
    });
  } finally { clearTimeout(timer); }
  if (!response.ok) {
    const detail = (await response.text().catch(() => "")).slice(0, 300);
    throw new Error(`AI 请求失败（HTTP ${response.status}）：${detail}`);
  }
  const data = await response.json();
  const result = String(data?.choices?.[0]?.message?.content ?? "").trim();
  if (!result) throw new Error("AI 未返回内容");
  return { task: input.task ?? "free", model, result };
}

/* ---------------- doctor ---------------- */

async function doctor() {
  const checks = { jszip: false, xlsx: false };
  try { checks.jszip = typeof JSZip?.loadAsync === "function"; } catch { /* ignore */ }
  try { checks.xlsx = typeof XLSX?.read === "function"; } catch { /* ignore */ }
  const ok = checks.jszip && checks.xlsx;
  return { ok, version: VERSION, checks, node: process.version };
}

/* ---------------- AI 管家：对话式调度全部插件能力 ---------------- */

async function agentProcess(input) {
  return runAgent(input, {
    binPath: fileURLToPath(import.meta.url),
    cwd: PLUGIN_DIR,
    tools: agentTools,
    systemPrompt,
    contextText,
    preparePayload,
    compact,
  });
}

/* ---------------- main ---------------- */

const COMMANDS = { doctor, docx2text, docx2rich, "docx-save-rich": docxSaveRich, "docx-edit": docxEdit, "docx-build": docxBuild, xlsx2json, "xlsx-grid": xlsxGrid, "xlsx-create": xlsxCreate, xlsx2csv, "xlsx-edit": xlsxEdit, csv2xlsx, pptx2json, "pptx-build": pptxBuild, "pdf-note": pdfNote, ai: aiProcess, agent: agentProcess };

const input = await readStdin();
// 宿主执行插件命令时会把 payload.mode 覆写成完整命令 ID（com.ruodao.office.xxx），
// 而每个命令的 action.args 里已经写死了短命令名；短名优先，再兼容完整 ID。
const argMode = String(process.argv[2] ?? "").trim();
const declaredMode = COMMANDS[argMode] ? argMode : String(input.mode ?? input.command ?? "").trim();
const mode = declaredMode.replace(/^com\.ruodao\.office\./i, "");
const handler = COMMANDS[mode];
if (!handler) fail(`未知命令：${declaredMode}（可用：${Object.keys(COMMANDS).join(" / ")}）`);
try {
  emit({ state: "succeeded", ...(await handler(input)) });
} catch (error) {
  fail(String(error?.message ?? error));
}
