import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { once } from "node:events";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { createApp } from "../host.mjs";

test("HTTP enrollment, employee rejection, authenticated MCP initialization, tool execution and revocation", async (t) => {
  const root = mkdtempSync(join(tmpdir(), "staff-http-"));
  const email = "pilot@example.test";
  const cfg = {
    employees: [email],
    netask: { [email]: "u123" },
    origin: "https://example.test",
    engine: "test.whl",
    engineHash: "sha",
  };
  writeFileSync(join(root, "config.json"), JSON.stringify(cfg));
  let who = email;
  const { app } = createApp(root, {
    exchange: async (_tenant, action) =>
      action === "devicecode"
        ? {
            device_code: "fake",
            verification_uri: "https://example.test",
            user_code: "fake",
            expires_in: 300,
            interval: 0.001,
          }
        : { access_token: "test", refresh_token: "test" },
    identify: async () => ({ id: "pilot-oid", userPrincipalName: who }),
  });
  const listener = app.listen(0, "127.0.0.1");
  await once(listener, "listening");
  t.after(async () => {
    await new Promise((r) => listener.close(r));
    rmSync(root, { recursive: true, force: true });
  });
  const base = "http://127.0.0.1:" + listener.address().port;
  const post = (path, body) =>
    fetch(base + path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  assert.equal((await post("/mcp", {})).status, 401);
  assert.equal((await fetch(base + "/engine/test.whl")).status, 401);
  assert.equal(
    (await post("/enroll", { email: "other@example.test" })).status,
    403,
  );
  const enroll = async () => {
    const p = await (await post("/enroll", { email })).json();
    await new Promise((r) => setTimeout(r, 10));
    return post("/enroll/finish", { key: p.key });
  };
  who = "other@example.test";
  assert.equal((await enroll()).status, 400);
  who = email;
  const token = (await (await enroll()).json()).token;
  assert.match(token, /^[\w-]{43}$/);
  const headers = { Authorization: "Bearer " + token };
  assert.equal(
    (await (await fetch(base + "/info", { headers })).json()).eip,
    "u123",
  );
  const m = new Client({ name: "http-acceptance", version: "1" });
  await m.connect(
    new StreamableHTTPClientTransport(new URL(base + "/mcp"), {
      requestInit: { headers },
    }),
  );
  const tools = (await m.listTools()).tools;
  assert.ok(tools.some((t) => t.name === "meeting_read"));
  assert.ok(!tools.some((t) => t.name === "portal_add_teammate"));
  const result = await m.callTool({ name: "whoami", arguments: {} });
  assert.equal(JSON.parse(result.content[0].text).userPrincipalName, email);
  await m.close();
  assert.equal(
    (await fetch(base + "/disconnect", { method: "POST", headers })).status,
    204,
  );
  assert.equal((await fetch(base + "/info", { headers })).status, 401);
});
