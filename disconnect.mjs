import { readFile, unlink } from "node:fs/promises";
import spawn from "cross-spawn";
import { cfgPath, api, browser, browserArgs } from "./local.mjs";
const cfg = JSON.parse(await readFile(cfgPath));
await api(cfg, "/disconnect", { method: "POST" });
await unlink(cfgPath);
for (const name of ["cica", "cica-eip"])
  spawn.sync("codex", ["mcp", "remove", name], { stdio: "inherit" });
spawn.sync(browser, ["--session", browserArgs(cfg).owner, "close"], {
  stdio: "inherit",
});
console.log(
  "Staff bearer revoked. To clear the retained local EIP login, remove the dedicated browser profile: " +
    browserArgs(cfg).profile,
);
