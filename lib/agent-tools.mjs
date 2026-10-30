/**
 * Office 办公套件的 AI 管家能力声明：工具集 = 插件全部命令 + 模型提示词 + 结果压缩。
 */
import os from "node:os";
import path from "node:path";
import { RISK_LABEL, truncate } from "./agent-core.mjs";

const OUTPUT_DIR = String(process.env.RUODAO_OFFICE_OUTPUT_DIR ?? "").trim()
  || path.join(os.homedir(), "Documents", "若道办公");

function safeName(value, fallback = "文档") {
  const text = String(value ?? "").replace(/[\\/:*?"<>|\r\n]+/g, " ").replace(/\s+/g, " ").trim();
  return (text || fallback).slice(0, 60);
}

function stamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
}

function withOutput(args, baseName, extension) {
  const next = { ...args };
  if (!String(next.output ?? "").trim()) {
    next.output = path.join(OUTPUT_DIR, `${safeName(baseName)}-${stamp()}.${extension}`);
  }
  return next;
}

const RISK_TEXT = (risk) => RISK_LABEL[risk] ?? risk;

export const tools = {
  doctor: {
    command: "doctor",
    title: "环境自检",
    risk: "read",
    timeoutMs: 30000,
    description: "检查办公套件内嵌解析库是否可用。遇到“无法读取文件”类问题时可先自检。",
    parameters: { type: "object", properties: {} },
    summarize: (args, data) => data?.ok ? "引擎正常" : "引擎异常",
  },
  docx2text: {
    command: "docx2text",
    title: "读取 Word",
    risk: "read",
    timeoutMs: 120000,
    description: "读取 .docx 的标题、段落与表格统计，返回块结构。需要文档内容时先调用它。",
    parameters: {
      type: "object",
      properties: { path: { type: "string", description: "docx 文件的绝对路径" } },
      required: ["path"],
    },
    summarize: (args, data) => `读取 ${data?.blocks?.length ?? 0} 个段落块`,
  },
  "docx-edit": {
    command: "docx-edit",
    title: "编辑 Word",
    risk: "write",
    timeoutMs: 120000,
    description: "对 docx 做批量文本替换并另存，保留原排版。适合改错别字、替换占位内容；不提供 output 时自动保存到输出目录。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string", description: "源 docx 绝对路径" },
        replacements: {
          type: "array",
          description: "替换列表，按顺序执行",
          items: {
            type: "object",
            properties: { from: { type: "string" }, to: { type: "string" } },
            required: ["from", "to"],
          },
        },
        output: { type: "string", description: "另存路径（可选，绝对路径）" },
      },
      required: ["path", "replacements"],
    },
    summarize: (args, data) => `已生成 ${data?.output ?? "文档"}（替换 ${data?.applied ?? args?.replacements?.length ?? 0} 处）`,
  },
  "docx-build": {
    command: "docx-build",
    title: "生成 Word",
    risk: "write",
    timeoutMs: 120000,
    description: "由标题/段落块生成新的 docx。适合写报告、纪要、方案。blocks 中 type 为 heading 或 paragraph，heading 可带 level 1-3。",
    parameters: {
      type: "object",
      properties: {
        title: { type: "string", description: "文档标题（也用于文件名）" },
        blocks: {
          type: "array",
          items: {
            type: "object",
            properties: {
              type: { type: "string", enum: ["heading", "paragraph"] },
              level: { type: "number", description: "heading 级别 1-3" },
              text: { type: "string" },
            },
            required: ["type", "text"],
          },
        },
        output: { type: "string", description: "保存路径（可选，绝对路径）" },
      },
      required: ["blocks"],
    },
    summarize: (args, data) => `已生成 ${data?.output ?? "Word 文档"}（${args?.blocks?.length ?? 0} 个段落块）`,
  },
  "xlsx-grid": {
    command: "xlsx-grid",
    title: "读取表格网格",
    risk: "read",
    timeoutMs: 120000,
    description: "把 xlsx/xls/csv 读成每个工作表的二维数组，适合查看和改写单元格。大表会截断。",
    parameters: {
      type: "object",
      properties: { path: { type: "string", description: "表格文件绝对路径" } },
      required: ["path"],
    },
    summarize: (args, data) => `读取 ${data?.sheets?.length ?? 0} 个工作表`,
  },
  "xlsx2json": {
    command: "xlsx2json",
    title: "读取表格行",
    risk: "read",
    timeoutMs: 120000,
    description: "把 xlsx/xls 读成以首行为表头的 JSON 行，适合做数据分析与汇总。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string" },
        sheet: { type: "string", description: "工作表名（可选）" },
        rowLimit: { type: "number", description: "最多读取行数（可选）" },
      },
      required: ["path"],
    },
    summarize: (args, data) => `读取 ${data?.sheets?.reduce((sum, sheet) => sum + (sheet.rows?.length ?? 0), 0) ?? 0} 行数据`,
  },
  "xlsx-create": {
    command: "xlsx-create",
    title: "生成 Excel",
    risk: "write",
    timeoutMs: 120000,
    description: "由二维数组生成 xlsx，可包含多个工作表。首行通常是表头；数字会自动识别。",
    parameters: {
      type: "object",
      properties: {
        sheets: {
          type: "array",
          items: {
            type: "object",
            properties: {
              name: { type: "string" },
              rows: { type: "array", items: { type: "array", items: {} } },
            },
            required: ["name", "rows"],
          },
        },
        output: { type: "string", description: "保存路径（可选，绝对路径）" },
      },
      required: ["sheets"],
    },
    summarize: (args, data) => `已生成 ${data?.output ?? "Excel"}（${args?.sheets?.length ?? 0} 个工作表）`,
  },
  "xlsx-edit": {
    command: "xlsx-edit",
    title: "编辑表格单元格",
    risk: "write",
    timeoutMs: 120000,
    description: "修改指定单元格后另存。cell 用 Excel 坐标，如 A1、C5。",
    parameters: {
      type: "object",
      properties: {
        path: { type: "string" },
        sheet: { type: "string", description: "工作表名（可选）" },
        edits: {
          type: "array",
          items: {
            type: "object",
            properties: { cell: { type: "string" }, value: {} },
            required: ["cell", "value"],
          },
        },
        output: { type: "string", description: "另存路径（可选）" },
      },
      required: ["path", "edits"],
    },
    summarize: (args, data) => `已生成 ${data?.output ?? "表格"}（修改 ${args?.edits?.length ?? 0} 个单元格）`,
  },
  "xlsx2csv": {
    command: "xlsx2csv",
    title: "表格转 CSV",
    risk: "write",
    timeoutMs: 120000,
    description: "把 xlsx/xls 的某个工作表导出为 CSV（UTF-8 BOM，Excel 友好）。",
    parameters: {
      type: "object",
      properties: { path: { type: "string" }, sheet: { type: "string" }, output: { type: "string" } },
      required: ["path"],
    },
    summarize: (args, data) => `已导出 ${data?.output ?? "CSV"}`,
  },
  csv2xlsx: {
    command: "csv2xlsx",
    title: "CSV 转 Excel",
    risk: "write",
    timeoutMs: 120000,
    description: "把 CSV 转为 xlsx。",
    parameters: {
      type: "object",
      properties: { path: { type: "string" }, output: { type: "string" } },
      required: ["path"],
    },
    summarize: (args, data) => `已生成 ${data?.output ?? "Excel"}`,
  },
  pptx2json: {
    command: "pptx2json",
    title: "解析 PPT",
    risk: "read",
    timeoutMs: 120000,
    description: "读取 .pptx 每一页的标题、要点与备注。",
    parameters: {
      type: "object",
      properties: { path: { type: "string" } },
      required: ["path"],
    },
    summarize: (args, data) => `解析 ${data?.slides?.length ?? 0} 页`,
  },
  "pptx-build": {
    command: "pptx-build",
    title: "生成 PPT",
    risk: "write",
    timeoutMs: 180000,
    description: "由大纲生成 pptx。slides 每页含 title、bullets（字符串数组）、notes（可选）。需要 AI 先设计大纲时直接组织内容即可。",
    parameters: {
      type: "object",
      properties: {
        title: { type: "string", description: "演示标题（也用于文件名）" },
        slides: {
          type: "array",
          items: {
            type: "object",
            properties: {
              title: { type: "string" },
              bullets: { type: "array", items: { type: "string" } },
              notes: { type: "string" },
            },
            required: ["title", "bullets"],
          },
        },
        output: { type: "string", description: "保存路径（可选，绝对路径）" },
      },
      required: ["slides"],
    },
    summarize: (args, data) => `已生成 ${data?.output ?? "PPT"}（${args?.slides?.length ?? 0} 页）`,
  },
  "pdf-note": {
    command: "pdf-note",
    title: "PDF 信息",
    risk: "read",
    timeoutMs: 60000,
    description: "读取 PDF 页数等元信息（不支持 PDF 内容编辑）。",
    parameters: {
      type: "object",
      properties: { path: { type: "string" } },
      required: ["path"],
    },
    summarize: (args, data) => `${data?.pages ?? 0} 页 PDF`,
  },
  ai: {
    command: "ai",
    title: "文档 AI 处理",
    risk: "read",
    timeoutMs: 150000,
    description: "对一段文本做摘要、要点、改写、翻译、续写或自由指令。适合处理长文档片段；短文本可直接自己处理。",
    parameters: {
      type: "object",
      properties: {
        task: { type: "string", enum: ["summarize", "points", "rewrite", "translate", "continue", "free"] },
        text: { type: "string", description: "待处理文本，最长 6 万字符" },
        instruction: { type: "string", description: "task=free 时的具体指令" },
      },
      required: ["task", "text"],
    },
    summarize: (args, data) => `完成「${args?.task ?? "AI 处理"}」，输出 ${String(data?.result ?? "").length} 字`,
  },
};

export function preparePayload(name, args) {
  const input = { ...(args ?? {}) };
  if (name === "docx-build") return withOutput(input, input.title, "docx");
  if (name === "docx-edit") {
    const base = path.basename(String(input.path ?? "文档")).replace(/\.docx$/i, "");
    return withOutput(input, `${base}-修改`, "docx");
  }
  if (name === "xlsx-create") return withOutput(input, input.sheets?.[0]?.name, "xlsx");
  if (name === "xlsx-edit") {
    const base = path.basename(String(input.path ?? "表格")).replace(/\.(xlsx|xls|csv)$/i, "");
    return withOutput(input, `${base}-修改`, "xlsx");
  }
  if (name === "xlsx2csv") {
    const base = path.basename(String(input.path ?? "表格")).replace(/\.(xlsx|xls)$/i, "");
    return withOutput(input, base, "csv");
  }
  if (name === "csv2xlsx") {
    const base = path.basename(String(input.path ?? "表格")).replace(/\.csv$/i, "");
    return withOutput(input, base, "xlsx");
  }
  if (name === "pptx-build") return withOutput(input, input.title, "pptx");
  return input;
}

export function compact(name, data) {
  if (!data) return null;
  if (name === "docx2text") {
    return {
      blocks: (data.blocks ?? []).slice(0, 400),
      stats: data.stats ?? null,
      text: truncate((data.blocks ?? []).map((block) => block.text).join("\n"), 12000),
    };
  }
  if (name === "xlsx-grid") {
    return {
      sheets: (data.sheets ?? []).map((sheet) => ({
        name: sheet.name,
        rows: (sheet.rows ?? []).slice(0, 40).map((row) => row.slice(0, 16)),
        totalRows: sheet.rows?.length ?? 0,
      })),
    };
  }
  if (name === "xlsx2json") {
    return {
      sheets: (data.sheets ?? []).map((sheet) => ({
        name: sheet.name,
        rows: (sheet.rows ?? []).slice(0, 40),
        totalRows: sheet.rows?.length ?? 0,
      })),
    };
  }
  if (name === "pptx2json") {
    return { slides: (data.slides ?? []).slice(0, 60) };
  }
  if (name === "ai") {
    return { task: data.task, result: truncate(data.result, 12000) };
  }
  if (name === "docx-edit") {
    return { output: data.output, applied: data.applied ?? null };
  }
  return data;
}

export function summarize(name, args, data) {
  const spec = tools[name];
  if (!spec?.summarize) return `${name} 完成`;
  return spec.summarize(args, data);
}

export function contextText(context) {
  const file = context?.activeFile;
  if (!file || !file.name) return "";
  const lines = [
    "【当前工作台上下文】",
    `用户正在查看：${file.name}${file.kind ? `（${file.kind}）` : ""}`,
    file.path ? `文件路径：${file.path}` : "该文件尚未保存到磁盘",
  ];
  if (file.text) lines.push("当前内容片段：", truncate(file.text, 16000));
  lines.push("用户说“当前文档/这份表格/这个 PPT”时，指的就是它；需要修改时优先使用它的路径，未保存文件请让用户先保存。");
  return lines.join("\n");
}

export const systemPrompt = [
  "你是若道 Office 办公套件的 AI 管家，负责代替用户完成 Word、Excel、PPT、CSV 与 PDF 的读取、生成、编辑与转换。",
  `默认输出目录：${OUTPUT_DIR}。用户未指定保存路径时，生成文件会放在这里；回答中必须给出真实产物路径。`,
  "",
  "工作规则：",
  "1. 需要文件内容时先调用读取工具，不要凭空编造文档内容或文件路径。",
  "2. 能用一个工具完成就不要拆成多个；确实需要多步时按顺序连续调用（如：读取 → 分析 → 生成）。",
  "3. 写入类工具会请求用户确认，确认前不要声称文件已生成。",
  "4. 长文本处理优先用 ai 工具；短文本可直接在回答里处理。",
  "5. 用户要求“翻译/总结/改写当前文档”时，先读取当前文档或使用上下文中的内容。",
  "6. 汇报用简洁中文：先给结论，再给产物路径、页数/段数等关键结果，最后给一句下一步建议。",
  "7. 不要调用 agent 自身，不要假设用户看不到的工具或能力。",
].join("\n");
