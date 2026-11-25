import { mkdir, readFile, writeFile, realpath, rename, stat } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import spawn from "cross-spawn";
import { unzipSync } from "fflate";
import { accountDir, browser, browserArgs, cfgPath, runNetask } from "./local.mjs";
import { powershell, verifyMailbox } from "./outlook.mjs";
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
const engine = "netask_eip-0.1.1-py3-none-any.whl";
const raw = await readFile(archive);
const data = archive.toLowerCase().endsWith(".zip")
  ? unzipSync(raw, { filter: (f) => f.name === engine })[engine]
  : raw;
if (
  !data ||
  createHash("sha256").update(data).digest("hex") !==
    "b83205297c7a209c5cd1c8a58905db704e966bf93e8a3db28acafc78d024ca72"
)
  throw Error("Download the current package from the portal.");
const syncRoot = syncFolder ? await realpath(syncFolder) : undefined;
if (syncRoot && !(await stat(syncRoot)).isDirectory()) throw Error("Choose a synced folder");
const cfg = {
  email,
  eip,
  engine,
  syncRoot,
  outlookCli: process.env.CICA_OUTLOOK_CLI
    ? await realpath(process.env.CICA_OUTLOOK_CLI)
    : undefined,
};
if (!cfg.outlookCli && process.platform !== "win32") {
  const candidate = spawn.sync("which", ["outlook"], { encoding: "utf8" });
  if (candidate.status === 0) {
    const path = await realpath(candidate.stdout.trim());
    if (
      await verifyMailbox({ ...cfg, outlookCli: path }).then(
        () => true,
        () => false,
      )
    )
      cfg.outlookCli = path;
  }
}
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
let mailReady = false;
try {
  if (!cfg.outlookCli) {
    // Microsoft manages the delegated session in the employee's local user profile.
    run(powershell, [
      "-NoProfile",
      "-Command",
      `
    $ErrorActionPreference = 'Stop'
    if (!(Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
      Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force
    }
    $context = Get-MgContext
    if (!$context -or $context.Account -ne '${email}' -or $context.Scopes -notcontains 'Mail.ReadWrite' -or $context.Scopes -notcontains 'User.Read') {
      Connect-MgGraph -Scopes User.Read,Mail.ReadWrite -ContextScope CurrentUser -NoWelcome
    }`,
    ]);
  }
  await verifyMailbox(cfg);
  mailReady = true;
} catch (error) {
  console.error("Outlook setup needs attention:", error.message);
  process.exitCode = 1;
}
run(browser, ["install"]);
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
console.log(
  "Complete your Netask sign-in in the opened window. Setup will verify the applicant and close the sign-in window.",
);
let identity;
const deadline = Date.now() + 600000;
while (Date.now() < deadline) {
  const status = await runNetask(cfg, ["status"], false).catch(() => null);
  if (status?.identity?.account === eip) {
    identity = status.identity;
    break;
  }
  await new Promise((resolve) => setTimeout(resolve, 2500));
}
if (!identity) throw Error("Finish signing in to the configured Netask account and rerun setup.");
run(browser, ["--session", owner, "close"]);
try {
  await runNetask(cfg, ["open"]);
  const status = await runNetask(cfg, ["status"]);
  if (status.identity?.account !== eip)
    throw Error("Netask headless session could not verify the applicant.");
} finally {
  run(browser, ["--session", owner, "close"]);
}
console.log(
  (mailReady
    ? "Mailbox and EIP identity verified."
    : "EIP identity verified. Complete Microsoft sign-in for mail.") +
    " Your documents folder is " +
    join(dir, "documents"),
);
