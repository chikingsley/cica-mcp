import { join, basename, resolve, relative, isAbsolute, sep } from "node:path";
import { homedir } from "node:os";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { randomUUID, createHash } from "node:crypto";
import { mkdir, readFile, writeFile, realpath, copyFile, constants } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { z } from "zod";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { server, tool } from "./common.mjs";
import { findNativeBinary, nativeTarget } from "./node_modules/agent-browser/scripts/native.js";
const exec = promisify(execFile);
import { prepareAttachments, verifyMailbox, compose, once } from "./outlook.mjs";
export const base = join(homedir(), ".cica-mcp");
export const browser = findNativeBinary(
  fileURLToPath(new URL("./node_modules/agent-browser/", import.meta.url)),
  nativeTarget(),
);
export const cfgPath = join(base, "account.json");
export const accountDir = (email) =>
  join(base, createHash("sha256").update(email).digest("hex").slice(0, 24));
export function browserArgs(cfg) {
  const dir = accountDir(cfg.email),
    owner = "cica-" + basename(dir);
  return { profile: join(dir, "browser"), owner };
}
export async function runNetask(cfg, args, headless = true) {
  const dir = accountDir(cfg.email);
  const { owner, profile } = browserArgs(cfg);

  const { stdout } = await exec(
    "uv",
    [
      "run",
      "--no-project",
      "--python",
      "3.14",
      "--with",
      join(dir, cfg.engine),
      "netask-eip",
      "--wrapper",
      browser,
      "--owner",
      owner,
      "--local-profile",
      profile,
      ...(headless && cfg.engine !== "netask_eip-0.1.0-py3-none-any.whl"
        ? ["--local-headless"]
        : []),
      "--account",
      cfg.eip,
      ...args,
    ],
    { maxBuffer: 12e6, timeout: 120000, windowsHide: true },
  );
  return JSON.parse(stdout);
}
export function localTools(cfg) {
  const m = server(),
    dir = accountDir(cfg.email),
    work = join(dir, "documents");
  const cli = (...args) => runNetask(cfg, args);
  const docname = z
    .string()
    .regex(/^[\p{L}\p{N} (._)-]{1,160}$/u)
    .refine((s) => !s.startsWith(".") && !s.endsWith(".") && !s.endsWith(" "));
  const document = async (name) => {
    const path = await realpath(join(work, docname.parse(name)));
    if (resolve(path) !== resolve(work, name))
      throw Error("Use a regular file in the employee workspace");
    return path;
  };
  tool(
    m,
    "eip_inspect",
    "Start a session with open, then verify identity with status. Consult the portal workflow_guide before preparing an NDA or EIP.",
    {
      view: z.enum([
        "open",
        "status",
        "request-schema",
        "directory",
        "drafts",
        "workflows",
        "templates",
        "examples",
        "form",
        "workflow",
      ]),
      id: z.int().positive().optional(),
    },
    async ({ view, id }) => {
      if (view === "drafts") return cli("workflows", "--folder", "draft");
      if (["form", "examples", "workflow"].includes(view)) {
        if (!id) throw Error("Choose a form or workflow ID");
        return cli(view, ...(view === "workflow" ? [] : ["--template"]), String(id));
      }
      return cli(view);
    },
  );
  const synced = async (subpath) => {
    if (!cfg.syncRoot) throw Error("Choose a synced company folder during setup");
    const root = await realpath(cfg.syncRoot);
    const path = await realpath(resolve(root, subpath));
    const rel = relative(root, path);
    if (isAbsolute(rel) || rel === ".." || rel.startsWith(".." + sep))
      throw Error("Choose a path within the configured synced folder");
    return path;
  };
  tool(
    m,
    "work_file_download",
    "Copy a file from the employee's synced company folder to the private documents folder.",
    { source: z.string(), filename: docname },
    async ({ source, filename }) => {
      const path = join(work, filename);
      await copyFile(await synced(source), path, constants.COPYFILE_EXCL);
      return { path };
    },
  );
  tool(
    m,
    "work_file_publish",
    "Copy a new document into a company folder. OneDrive handles upload.",
    { folder: z.string(), filename: docname },
    async ({ folder, filename }) => {
      const path = join(await synced(folder), filename);
      await copyFile(await document(filename), path, constants.COPYFILE_EXCL);
      return { localPath: path };
    },
  );
  tool(
    m,
    "outlook_draft",
    "Save and read back an unsent Outlook API draft in your own mailbox. Reuse the same request ID on retry.",
    {
      request_id: z.uuid(),
      to: z.array(z.email()).min(1),
      subject: z.string().trim().min(1).max(240),
      body: z.string().trim().min(1),
      attachments: z.array(docname).default([]),
    },
    async (args) => {
      const paths = await Promise.all(args.attachments.map(document));
      const files = await prepareAttachments(paths);
      await verifyMailbox(cfg);
      return once(dir, args, () => compose(cfg, args, files));
    },
  );
  tool(
    m,
    "eip_prepare",
    "Prepare a typed EIP draft from request_schema. Employee identity is taken from Netask. Supply evidence files by name from the documents folder. Inspect prior requests and confirm company procedure and routing first.",
    {
      request: z.record(z.string(), z.unknown()),
      attachments: z
        .array(z.object({ filename: docname, roles: z.array(z.string()).min(1) }))
        .min(1),
    },
    async ({ request, attachments }) => {
      if (!Array.isArray(request.route) || !request.route.length)
        throw Error(
          "Supply the employee-specific route from current form examples or the approved company procedure",
        );
      const status = await cli("status");
      if (status.identity?.account !== cfg.eip) throw Error("Sign in to your Netask account");
      const evidence = await Promise.all(
        attachments.map(async (a) => ({
          path: await document(a.filename),
          roles: a.roles,
        })),
      );
      const id = randomUUID(),
        folder = join(dir, "packets", id);
      await mkdir(folder, { recursive: true, mode: 0o700 });
      const path = join(folder, "request.json");
      await writeFile(
        path,
        JSON.stringify({
          ...request,
          applicant: status.identity.name,
          attachments: evidence,
        }),
        { mode: 0o600 },
      );
      const prepared = await cli("prepare", path);
      await writeFile(join(folder, "packet.json"), JSON.stringify(prepared), {
        mode: 0o600,
      });
      return { id, prepared };
    },
  );
  tool(
    m,
    "eip_draft_step",
    "Check, open, fill, attach or save a prepared EIP draft. Use editor for its fingerprint and recover interrupted operations by attempt ID. Never submits.",
    {
      id: z.uuid(),
      step: z.enum([
        "draft-check",
        "draft-open",
        "editor",
        "draft-fill",
        "draft-attachments",
        "draft-upload",
        "draft-upload-status",
        "draft-save",
        "draft-save-status",
      ]),
      fingerprint: z.string().optional(),
      attachment: z.int().nonnegative().optional(),
      attempt: z.string().optional(),
    },
    async ({ id, step, fingerprint, attachment, attempt }) => {
      const path = join(dir, "packets", id, "packet.json");
      if (step === "editor") return cli(step);
      if (step === "draft-upload-status" || step === "draft-save-status") {
        if (!attempt) throw Error("Provide attempt ID");
        return cli(
          step,
          ...(step === "draft-save-status" ? [path, "--attempt", attempt] : [attempt]),
        );
      }
      const args = [step, path];
      if (["draft-fill", "draft-attachments", "draft-upload", "draft-save"].includes(step)) {
        if (!fingerprint) throw Error("Inspect editor first");
        args.push("--expected-state", fingerprint);
      }
      if (step === "draft-upload") {
        if (attachment === undefined) throw Error("Select an attachment index");
        args.push("--attachment", String(attachment));
      }
      return cli(...args);
    },
  );
  return m;
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const cfg = JSON.parse(await readFile(cfgPath));
  await localTools(cfg).connect(new StdioServerTransport());
}
