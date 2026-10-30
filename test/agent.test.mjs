import test from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import fsp from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chatEndpoint } from "../lib/agent-core.mjs";

const PLUGIN_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const CLI = path.join(PLUGIN_DIR, "bin", "office.mjs");
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), "office-agent-test-"));

const llmEnv = (baseUrl) => ({ RUODAO_LLM_BASE_URL: baseUrl, RUODAO_LLM_MODEL: "stub-model", RUODAO_LLM_API_KEY: "test-key" });

function toolCall(id, name, args) {
  return {
    role: "assistant",
    content: "",
    tool_calls: [{ id, type: "function", function: { name, arguments: JSON.stringify(args ?? {}) } }],
  };
}

function reply(content) {
  return { role: "assistant", content };
}

/** 起一个只按脚本顺序应答的 OpenAI 兼容端点，并记录收到的请求体。 */
async function stubServer(script) {
  const requests = [];
  const server = http.createServer((req, res) => {
    let body = "";
    req.on("data", (chunk) => { body += chunk; });
    req.on("end", () => {
      let parsed = null;
      try { parsed = JSON.parse(body); } catch { parsed = null; }
      requests.push({ url: req.url, body: parsed });
      const message = script[Math.min(requests.length - 1, script.length - 1)];
      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ choices: [{ message }], usage: { total_tokens: 1 } }));
    });
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const { port } = server.address();
  return {
    baseUrl: `http://127.0.0.1:${port}/v1`,
    requests,
    close: () => new Promise((resolve) => server.close(resolve)),
  };
}

function runAgent(payload, llmEnv = {}) {
  return new Promise((resolve, reject) => {
    const env = { ...process.env, ...llmEnv };
    if (!llmEnv.RUODAO_LLM_BASE_URL) {
      delete env.RUODAO_LLM_BASE_URL;
      delete env.RUODAO_LLM_MODEL;
    }
    const child = spawn(process.execPath, [CLI, "agent"], { cwd: PLUGIN_DIR, windowsHide: true, env });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", reject);
    child.on("close", () => {
      const line = String(stdout).trim().split(/\r?\n/).filter(Boolean).at(-1) ?? "";
      let data = null;
      try { data = JSON.parse(line); } catch { data = null; }
      if (!data) reject(new Error(`非 JSON 输出：${stdout || stderr}`));
      else resolve(data);
    });
    child.stdin.end(JSON.stringify(payload));
  });
}

test("AI 管家工具覆盖清单中的全部命令（agent 自身除外）", async () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(PLUGIN_DIR, "ruodao-plugin.json"), "utf8"));
  const commands = manifest.contributes.commands.map((command) => String(command.id).split(".").pop());
  const { tools } = await import("../lib/agent-tools.mjs");
  for (const name of commands) {
    if (name === "agent") continue;
    assert.ok(tools[name], `缺少工具：${name}`);
  }
});

test("chatEndpoint 兼容带与不带 /chat/completions 的基地址", () => {
  assert.equal(chatEndpoint("https://api.example.com/v1"), "https://api.example.com/v1/chat/completions");
  assert.equal(chatEndpoint("https://api.example.com/v1/"), "https://api.example.com/v1/chat/completions");
  assert.equal(chatEndpoint("https://api.example.com/v1/chat/completions"), "https://api.example.com/v1/chat/completions");
  assert.equal(chatEndpoint(""), "");
});

test("只读工具自动执行，写入工具等待审批，批准后继续", async () => {
  const output = path.join(TMP, `agent-${Date.now()}.docx`);
  const stub = await stubServer([
    toolCall("call_doctor", "doctor", {}),
    toolCall("call_build", "docx-build", { title: "AI 测试", blocks: [{ type: "heading", level: 1, text: "AI 生成" }], output }),
    reply("已生成文档。"),
  ]);
  try {
    const first = await runAgent({ messages: [{ role: "user", content: "生成一份测试文档" }] }, llmEnv(stub.baseUrl));
    assert.equal(first.state, "succeeded");
    assert.equal(first.status, "awaiting_approval");
    assert.equal(first.pending.command, "docx-build");
    assert.equal(first.pending.risk, "write");
    assert.equal(first.steps.length, 1);
    assert.equal(first.steps[0].command, "doctor");
    assert.equal(first.steps[0].ok, true);
    assert.equal(fs.existsSync(output), false, "审批前不应写文件");
    assert.equal(stub.requests[0].url, "/v1/chat/completions");
    assert.equal(stub.requests[0].body.tool_choice, "auto");
    assert.ok(stub.requests[0].body.tools.some((tool) => tool.function.name === "docx-build"));

    const second = await runAgent({
      messages: first.messages,
      approvals: [first.pending.id],
    }, llmEnv(stub.baseUrl));
    assert.equal(second.status, "done");
    assert.equal(second.reply, "已生成文档。");
    assert.ok(second.steps.some((step) => step.command === "docx-build" && step.ok));
    assert.equal(fs.existsSync(output), true);
    const toolMessage = stub.requests[1].body.messages.find((message) => message.role === "tool");
    assert.ok(toolMessage.content.includes("引擎正常"));
  } finally {
    await stub.close();
    await fsp.unlink(output).catch(() => undefined);
  }
});

test("拒绝写入工具时不执行，并把拒绝结果回传给模型", async () => {
  const output = path.join(TMP, `agent-reject-${Date.now()}.docx`);
  const stub = await stubServer([
    toolCall("call_build", "docx-build", { title: "不该生成", blocks: [], output }),
    reply("已按你的要求取消。"),
  ]);
  try {
    const first = await runAgent({ messages: [{ role: "user", content: "生成文档" }] }, llmEnv(stub.baseUrl));
    assert.equal(first.status, "awaiting_approval");
    assert.equal(fs.existsSync(output), false);
    const second = await runAgent({
      messages: first.messages,
      rejections: [first.pending.id],
    }, llmEnv(stub.baseUrl));
    assert.equal(second.status, "done");
    assert.equal(fs.existsSync(output), false);
    const rejectedStep = second.steps.find((step) => step.command === "docx-build");
    assert.equal(rejectedStep.rejected, true);
    assert.equal(rejectedStep.ok, false);
    const toolMessage = stub.requests[1].body.messages.find((message) => message.role === "tool");
    assert.ok(toolMessage.content.includes("拒绝"));
  } finally {
    await stub.close();
    await fsp.unlink(output).catch(() => undefined);
  }
});

test("工具执行失败以错误结果回传，不中断本轮对话", async () => {
  const missing = path.join(TMP, "not-exist.docx");
  const stub = await stubServer([
    toolCall("call_read", "docx2text", { path: missing }),
    reply("文件不存在，请确认路径。"),
  ]);
  try {
    const result = await runAgent({ messages: [{ role: "user", content: "读取这个文件" }] }, llmEnv(stub.baseUrl));
    assert.equal(result.status, "done");
    assert.equal(result.reply, "文件不存在，请确认路径。");
    assert.equal(result.steps[0].ok, false);
    assert.ok(result.steps[0].error);
    const toolMessage = stub.requests[1].body.messages.find((message) => message.role === "tool");
    const payload = JSON.parse(toolMessage.content.split("\n")[0]);
    assert.equal(payload.ok, false);
  } finally {
    await stub.close();
  }
});

test("缺少模型配置时返回明确错误", async () => {
  const result = await runAgent({ messages: [{ role: "user", content: "你好" }] });
  assert.equal(result.state, "failed");
  assert.ok(String(result.error).includes("未配置 AI 模型"));
});
