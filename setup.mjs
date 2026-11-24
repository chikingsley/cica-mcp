import { readFile, writeFile, mkdir, rename } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import spawn from "cross-spawn";
import { setTimeout as delay } from "node:timers/promises";
import {
  base,
  origin,
  cfgPath,
  api,
  accountDir,
  browser,
  browserArgs,
} from "./local.mjs";
process.umask(0o077);
function run(cmd, args) {
  const child = spawn.sync(cmd, args, { stdio: "inherit" });
  if (child.status !== 0) throw Error("Setup command failed: " + cmd);
}
async function post(path, body) {
  const r = await fetch(origin + path, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body),
  });
  if (r.status === 202) return null;
  if (!r.ok)
    throw Error(
      "Enrollment returned " + r.status + " " + (await r.text()).slice(0, 200),
    );
  return r.json();
}
const email = (process.argv[2] || "").toLowerCase().trim();
if (!email.includes("@"))
  throw Error("Run node setup.mjs with your work email address.");
run("uv", ["--version"]);
run("codex", ["--version"]);
let cfg = await readFile(cfgPath, "utf8")
  .then(JSON.parse)
  .catch(() => null);
if (
  cfg?.email !== email ||
  !(await api(cfg, "/info")
    .then(() => true)
    .catch(() => false))
) {
  const p = await post("/enroll", { email });
  console.log(
    "Open " +
      p.url +
      " yourself. Enter " +
      p.code +
      " and sign in to your work account. This code expires shortly.",
  );
  let authorized;
  const end = Date.now() + p.expires_in * 1000;
  while (!authorized && Date.now() < end) {
    await delay((p.interval + 1) * 1000);
    authorized = await post("/enroll/finish", { key: p.key });
  }
  if (!authorized) throw Error("Microsoft sign-in timed out");
  cfg = { email, token: authorized.token };
}
const info = await (await api(cfg, "/info")).json();
const dir = accountDir(email);
if (
  info.email !== email ||
  !/^u\d+$/.test(info.eip) ||
  !/^netask_eip-[\w.-]+\.whl$/.test(info.engine)
)
  throw Error("Invalid staff account mapping");
await mkdir(dir, { recursive: true, mode: 0o700 });
const bytes = Buffer.from(
  await (await api(cfg, "/engine/" + info.engine)).arrayBuffer(),
);
if (createHash("sha256").update(bytes).digest("hex") !== info.sha256)
  throw Error("Netask package verification failed");
await writeFile(join(dir, info.engine), bytes, { mode: 0o600 });
cfg = { ...cfg, eip: info.eip, engine: info.engine };
const tmp = cfgPath + "." + randomUUID();
await writeFile(tmp, JSON.stringify(cfg), { mode: 0o600 });
await rename(tmp, cfgPath);
await mkdir(join(dir, "documents"), { recursive: true, mode: 0o700 });
run("uv", [
  "run",
  "--no-project",
  "--python",
  "3.14",
  "--with",
  join(dir, cfg.engine),
  "netask-eip",
  "--help",
]);
for (const [name, file] of [
  ["cica", "bridge.mjs"],
  ["cica-eip", "local.mjs"],
]) {
  run("codex", [
    "mcp",
    "add",
    name,
    "--",
    process.execPath,
    fileURLToPath(new URL("./" + file, import.meta.url)),
  ]);
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
  "Finish your own Netask sign-in in the opened browser. Start a fresh Codex chat and ask it to check cica whoami and cica-eip status. Work documents folder: " +
    join(dir, "documents"),
);
