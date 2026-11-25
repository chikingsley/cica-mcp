import { readFile, unlink } from "node:fs/promises";
import spawn from "cross-spawn";
import { cfgPath, browser, browserArgs, accountDir } from "./local.mjs";
const cfg = JSON.parse(await readFile(cfgPath));
for (const name of ["cica", "cica-portal"])
  spawn.sync("codex", ["mcp", "remove", name], { stdio: "inherit" });
spawn.sync(browser, ["--session", browserArgs(cfg).owner, "close"], { stdio: "inherit" });
await unlink(cfgPath);
console.log("Local profiles and documents remain available at " + accountDir(cfg.email));
