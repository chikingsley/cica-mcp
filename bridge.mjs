// Codex stdio transport for the hosted MCP. Keeps the per-user bearer out of
// command arguments, shell history and the Codex config on Windows and Unix.
import { createInterface } from "node:readline";
import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
const cfg = JSON.parse(
  await readFile(join(homedir(), ".cica-mcp", "account.json")),
);
for await (const line of createInterface({
  input: process.stdin,
  crlfDelay: Infinity,
})) {
  let msg;
  try {
    msg = JSON.parse(line);
    const r = await fetch("https://cica-mcp.grassinside.com/mcp", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + cfg.token,
        "Content-Type": "application/json",
        Accept: "application/json, text/event-stream",
      },
      body: line,
      signal: AbortSignal.timeout(120000),
    });
    if (!r.ok)
      throw Error(
        "Staff connection returned " +
          r.status +
          ". Run setup again if sign-in expired.",
      );
    if (r.status !== 202 && msg.id !== undefined)
      console.log(JSON.stringify(await r.json()));
  } catch (e) {
    if (msg?.id !== undefined)
      console.log(
        JSON.stringify({
          jsonrpc: "2.0",
          id: msg.id,
          error: { code: -32000, message: e.message },
        }),
      );
  }
}
