import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { once } from "../outlook.mjs";
import { localTools } from "../local.mjs";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
test("draft retry reuses the receipt and interrupted requests require readback", async (t) => {
  const dir = await mkdtemp(join(tmpdir(), "staff-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  let calls = 0;
  const args = { request_id: "draft", subject: "Example" };
  const action = async () => ({ count: ++calls });
  assert.deepEqual(await once(dir, args, action), await once(dir, args, action));
  assert.equal(calls, 1);
  await assert.rejects(once(dir, { ...args, subject: "different" }, action), /another draft/);
  const interrupted = { request_id: "interrupted" };
  await assert.rejects(
    once(dir, interrupted, async () => {
      throw Error("lost browser");
    }),
  );
  await assert.rejects(once(dir, interrupted, action), /Check Outlook Drafts/);
});
test("workstation MCP lists draft and EIP tools and rejects paths outside the sync root", async (t) => {
  const c = new Client({ name: "test", version: "1" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await localTools({ email: "staff@example.test", syncRoot: tmpdir() }).connect(a);
  await c.connect(b);
  t.after(() => c.close());
  const names = (await c.listTools()).tools.map((t) => t.name);
  for (const n of [
    "eip_inspect",
    "eip_prepare",
    "eip_draft_step",
    "outlook_draft",
    "work_file_download",
    "work_file_publish",
  ])
    assert.ok(names.includes(n));
  const bad = await c.callTool({
    name: "work_file_download",
    arguments: { source: "..", filename: "example.pdf" },
  });
  assert.ok(bad.isError);
});
