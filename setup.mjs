import { mkdir, readFile, writeFile, realpath, rename, stat } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import spawn from "cross-spawn";
import { unzipSync } from "fflate";
import { accountDir, browser, browserArgs, cfgPath } from "./local.mjs";
import { mailbox } from "./outlook.mjs";
const email = (process.argv[2] ?? "").toLowerCase().trim();
const [, , , eip, archive, syncFolder] = process.argv;
if (!/^[\w.+-]+@cicahuntek\.us$/.test(email) || !/^u\d+$/.test(eip) || !archive)
  throw Error(
    "Provide your company email, Netask user ID, downloaded Netask ZIP and optional synced company folder.",
  );
const suffix = email.match(/\.(\d+)@/);
if (suffix && suffix[1] !== eip.slice(1))
  throw Error("Verify your email and Netask employee number.");
const dir = accountDir(email);
const engine = "netask_eip-0.1.0-py3-none-any.whl";
const raw = await readFile(archive);
const data = archive.toLowerCase().endsWith(".zip")
  ? unzipSync(raw, { filter: (f) => f.name === engine })[engine]
  : raw;
if (
  !data ||
  createHash("sha256").update(data).digest("hex") !==
    "6fb3894dbb6606f1114a37f18eb06837186038dd515f5edf81d250b4c9504b2f"
)
  throw Error("Download the current package from the portal.");
const syncRoot = syncFolder ? await realpath(syncFolder) : undefined;
if (syncRoot && !(await stat(syncRoot)).isDirectory()) throw Error("Choose a synced folder");
const cfg = { email, eip, engine, syncRoot };
function run(cmd, args) {
  const r = spawn.sync(cmd, args, { stdio: "inherit" });
  if (r.status !== 0) throw Error("Setup command failed: " + cmd);
}
run("codex", ["--version"]);
run("uv", ["--version"]);
await mkdir(join(dir, "documents"), { recursive: true, mode: 0o700 });
await writeFile(join(dir, engine), data, { mode: 0o600 });
run("uv", [
  "run",
  "--no-project",
  "--python",
  "3.14",
  "--with",
  join(dir, engine),
  "netask-eip",
  "--help",
]);
run(process.execPath, [
  fileURLToPath(new URL("./node_modules/playwright/cli.js", import.meta.url)),
  "install",
  "chromium",
]);
run(browser, ["install"]);
console.log("Sign in to Outlook in the browser window that opens.");
await mailbox(dir, email, async () => console.log("Verified your Outlook mailbox."));
const { owner, profile } = browserArgs(cfg);
run(browser, [
  "--session",
  owner,
  "--profile",
  profile,
  "--headed",
  "open",
  "https://eip.cicahuntek.us/tw/findex.php",
]);
const tmp = cfgPath + "." + randomUUID();
await writeFile(tmp, JSON.stringify(cfg), { mode: 0o600 });
await rename(tmp, cfgPath);
run("codex", ["mcp", "add", "cica-portal", "--url", "https://cica.grassinside.com/mcp"]);
run("codex", [
  "mcp",
  "add",
  "cica",
  "--",
  process.execPath,
  fileURLToPath(new URL("./local.mjs", import.meta.url)),
]);
console.log(
  "Finish signing in to Netask. In a new Codex chat, ask it to verify EIP status. Your documents folder is " +
    join(dir, "documents"),
);
