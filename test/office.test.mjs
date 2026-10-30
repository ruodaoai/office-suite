import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import fsp from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const PLUGIN_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const JSZip = require(path.join(PLUGIN_DIR, "vendor", "jszip.min.js"));

const CLI = path.join(PLUGIN_DIR, "bin", "office.mjs");

function run(mode, payload = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [CLI, mode], { cwd: PLUGIN_DIR, windowsHide: true });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (c) => (stdout += c));
    child.stderr.on("data", (c) => (stderr += c));
    child.on("error", reject);
    child.on("close", () => {
      try { resolve(JSON.parse(stdout.trim().split("\n").pop() || "{}")); }
      catch { reject(new Error(`非 JSON 输出：${stdout || stderr}`)); }
    });
    child.stdin.end(JSON.stringify({ mode, ...payload }));
  });
}

async function makeDocx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>项目计划</w:t></w:r></w:p><w:p><w:r><w:t>第一段：目标与范围说明。</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`);
  const out = path.join(os.tmpdir(), `office-suite-test-${Date.now()}.docx`);
  await fsp.writeFile(out, await zip.generateAsync({ type: "nodebuffer" }));
  return out;
}

test("doctor 自检通过", async () => {
  const result = await run("doctor");
  assert.equal(result.state, "succeeded");
  assert.equal(result.ok, true);
});

test("宿主用完整命令 ID 覆写 mode 时仍命中短命令", async () => {
  const result = await run("doctor", { mode: "com.ruodao.office.doctor" });
  assert.equal(result.state, "succeeded");
  assert.equal(result.ok, true);
});

test("docx2text 提取标题/段落/表格", async () => {
  const file = await makeDocx();
  const result = await run("docx2text", { path: file });
  assert.equal(result.state, "succeeded");
  assert.equal(result.blocks[0].type, "heading");
  assert.equal(result.blocks[0].text, "项目计划");
  assert.ok(result.blocks.some((b) => b.text.includes("第一段")));
  assert.equal(result.stats.tables, 1);
  await fsp.unlink(file);
});

test("docx-edit 文本替换", async () => {
  const file = await makeDocx();
  const result = await run("docx-edit", { path: file, replacements: [{ from: "目标与范围", to: "目标与边界" }] });
  assert.equal(result.state, "succeeded");
  assert.ok(fs.existsSync(result.output));
  const check = await run("docx2text", { path: result.output });
  assert.ok(check.blocks.some((b) => b.text.includes("目标与边界")));
  await fsp.unlink(file); await fsp.unlink(result.output);
});

test("xlsx2json / xlsx-edit / csv2xlsx", async () => {
  const XLSX = require(path.join(PLUGIN_DIR, "vendor", "xlsx.full.min.js"));
  const src = path.join(os.tmpdir(), `office-suite-test-${Date.now()}.xlsx`);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ 名称: "A", 数量: 3 }, { 名称: "B", 数量: 7 }]), "数据");
  fs.writeFileSync(src, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  const parsed = await run("xlsx2json", { path: src });
  assert.equal(parsed.state, "succeeded");
  assert.equal(parsed.sheets[0].rows[1].数量, 7);
  const edited = await run("xlsx-edit", { path: src, edits: [{ cell: "B3", value: 99 }] });
  assert.equal(edited.state, "succeeded");
  const after = await run("xlsx2json", { path: edited.output });
  assert.equal(after.sheets[0].rows[1].数量, 99);
  const csv = await run("xlsx2csv", { path: src });
  assert.equal(csv.state, "succeeded");
  const back = await run("csv2xlsx", { path: csv.output });
  assert.equal(back.state, "succeeded");
  assert.ok(back.sheets.length >= 1);
  await Promise.all([src, edited.output, csv.output, back.output].map((p) => fsp.unlink(p).catch(() => undefined)));
});

test("pptx-build 生成可解析的 pptx", async () => {
  const out = path.join(os.tmpdir(), `office-suite-test-${Date.now()}.pptx`);
  const built = await run("pptx-build", { title: "测试演示", slides: [{ title: "封面", bullets: [{ text: "要点一" }, { text: "要点二", level: 1 }], notes: "备注内容" }, { title: "第二页", bullets: [] }] });
  assert.equal(built.state, "succeeded");
  assert.ok(fs.existsSync(built.output));
  const parsed = await run("pptx2json", { path: built.output });
  assert.equal(parsed.state, "succeeded");
  assert.equal(parsed.stats.slides, 2);
  assert.equal(parsed.slides[0].title, "封面");
  assert.ok(parsed.slides[0].paragraphs.some((t) => t.includes("要点一")));
  await fsp.unlink(built.output);
});

test("docx-build 生成可回读的 Word 文档", async () => {
  const out = path.join(os.tmpdir(), `office-suite-build-${Date.now()}.docx`);
  const built = await run("docx-build", { title: "测试文档", blocks: [{ type: "heading", level: 1, text: "季度总结" }, { type: "paragraph", text: "本季度完成 <A&B> 项目。" }, { type: "heading", level: 2, text: "下一步" }], output: out });
  assert.equal(built.state, "succeeded");
  const parsed = await run("docx2text", { path: built.output });
  assert.equal(parsed.state, "succeeded");
  assert.equal(parsed.blocks[0].type, "heading");
  assert.equal(parsed.blocks[0].text, "季度总结");
  assert.ok(parsed.blocks.some((b) => b.text.includes("<A&B>")));
  await fsp.unlink(built.output);
});

test("xlsx-grid 读取 / xlsx-create 多工作表生成", async () => {
  const XLSX = require(path.join(PLUGIN_DIR, "vendor", "xlsx.full.min.js"));
  const src = path.join(os.tmpdir(), `office-suite-grid-${Date.now()}.xlsx`);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["姓名", "部门"], ["张三", "研发"]]), "名单");
  fs.writeFileSync(src, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
  const grid = await run("xlsx-grid", { path: src });
  assert.equal(grid.state, "succeeded");
  assert.deepEqual(grid.sheets[0].rows[1], ["张三", "研发"]);
  const built = await run("xlsx-create", { sheets: [{ name: "预算", rows: [["项目", "金额"], ["服务器", "12000"], ["合计", "12000"]] }, { name: "备注", rows: [["说明"], ["<测试>"]] }] });
  assert.equal(built.state, "succeeded");
  const check = await run("xlsx-grid", { path: built.output });
  assert.deepEqual(check.sheets.map((s) => s.name), ["预算", "备注"]);
  assert.equal(check.sheets[0].rows[1][1], "12000");
  await Promise.all([src, built.output].map((p) => fsp.unlink(p).catch(() => undefined)));
});

test("非法输入报错", async () => {
  const bad = await run("docx2text", { path: path.join(os.tmpdir(), "not-exist.docx") });
  assert.equal(bad.state, "failed");
  const unknown = await run("no-such-mode");
  assert.equal(unknown.state, "failed");
});

test("docx-save-rich 富文本往返（标题/样式/列表/表格）与原地保存", async () => {
  const out = path.join(os.tmpdir(), `office-suite-rich-${Date.now()}.docx`);
  const blocks = [
    { type: "paragraph", heading: 1, runs: [{ text: "季度总结" }] },
    { type: "paragraph", runs: [{ text: "普通 " }, { text: "重点", b: true, color: "#ff0000" }] },
    { type: "paragraph", list: { level: 0, ordered: false }, runs: [{ text: "列表项" }] },
    { type: "table", header: true, rows: [
      [[{ text: "名称" }], [{ text: "数值" }]],
      [[{ text: "服务器" }], [{ text: "12000" }]],
    ] },
  ];
  const built = await run("docx-save-rich", { blocks, output: out, title: "富文本" });
  assert.equal(built.state, "succeeded");
  assert.ok(fs.existsSync(out));

  const rich = await run("docx2rich", { path: out });
  assert.equal(rich.state, "succeeded");
  const heading = rich.blocks.find((block) => block.type === "paragraph" && block.heading === 1);
  assert.ok(heading, "应保留标题层级");
  assert.equal(heading.runs.map((run) => run.text).join(""), "季度总结");
  const styled = rich.blocks.find((block) => block.type === "paragraph" && block.runs?.some((run) => run.text === "重点"));
  const boldRun = styled.runs.find((run) => run.text === "重点");
  assert.equal(boldRun.b, true);
  assert.equal(String(boldRun.color).toLowerCase(), "#ff0000");
  assert.ok(rich.blocks.some((block) => block.type === "paragraph" && block.list), "应保留项目符号列表");
  assert.equal(rich.stats.tables, 1);
  const table = rich.blocks.find((block) => block.type === "table");
  assert.equal(table.rows[1][0].map((run) => run.text).join(""), "服务器");

  const replaced = await run("docx-save-rich", { blocks: [{ type: "paragraph", runs: [{ text: "替换后" }] }], path: out, inPlace: true });
  assert.equal(replaced.state, "succeeded");
  assert.equal(path.resolve(replaced.output), path.resolve(out));
  const after = await run("docx2rich", { path: out });
  assert.equal(after.blocks[0].runs.map((run) => run.text).join(""), "替换后");
  await fsp.unlink(out);
});

test("pptx-build 应用主题配色、版式与 16:9 尺寸", async () => {
  const out = path.join(os.tmpdir(), `office-suite-theme-${Date.now()}.pptx`);
  const built = await run("pptx-build", {
    title: "主题演示",
    theme: { bg: "0F172A", bgAlt: "1E293B", primary: "818CF8", accent: "22D3EE", text: "F8FAFC", muted: "CBD5E1", font: "微软雅黑", decor: "grid" },
    output: out,
    slides: [
      { title: "封面标题", bullets: ["副标题"], layout: "cover" },
      { title: "内容页面", bullets: ["要点一", { text: "二级要点", level: 1 }], layout: "content", notes: "演讲备注" },
      { title: "谢谢观看", bullets: ["欢迎交流"], layout: "end" },
    ],
  });
  assert.equal(built.state, "succeeded");
  const zip = await JSZip.loadAsync(await fsp.readFile(out));
  const themeXml = await zip.file("ppt/theme/theme1.xml").async("string");
  assert.match(themeXml, /818CF8/i);
  assert.match(themeXml, /22D3EE/i);
  const presentation = await zip.file("ppt/presentation.xml").async("string");
  assert.ok(presentation.includes('cx="12192000"') && presentation.includes('cy="6858000"'), "应为 16:9");
  const slide1 = await zip.file("ppt/slides/slide1.xml").async("string");
  const slide2 = await zip.file("ppt/slides/slide2.xml").async("string");
  assert.ok(slide1.includes("封面标题"));
  assert.ok(slide2.includes("内容页面"));
  const parsed = await run("pptx2json", { path: out });
  assert.equal(parsed.stats.slides, 3);
  await fsp.unlink(out);
});

test("xlsx-create 写入真实公式单元格", async () => {
  const XLSX = require(path.join(PLUGIN_DIR, "vendor", "xlsx.full.min.js"));
  const built = await run("xlsx-create", { sheets: [{ name: "预算", rows: [["项目", "预算", "实际", "差额"], ["人力", "100", "80", "=B2-C2"], ["合计", "=SUM(B2:B2)", "=SUM(C2:C2)", "=B3-C3"]] }] });
  assert.equal(built.state, "succeeded");
  const wb = XLSX.read(await fsp.readFile(built.output), { type: "buffer" });
  const sheet = wb.Sheets["预算"];
  assert.equal(sheet.D2.f, "B2-C2");
  assert.equal(sheet.B3.f, "SUM(B2:B2)");
  const grid = await run("xlsx-grid", { path: built.output });
  assert.deepEqual(grid.sheets[0].rows[1].slice(0, 3), ["人力", "100", "80"]);
  assert.equal(grid.sheets[0].formulas?.D2, "=B2-C2");
  assert.equal(grid.sheets[0].formulas?.B3, "=SUM(B2:B2)");
  await fsp.unlink(built.output);
});
