import { join, basename, resolve } from "node:path";
import { homedir } from "node:os";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { randomUUID, createHash } from "node:crypto";
import { mkdir, readFile, writeFile, realpath } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { z } from "zod";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { server, tool } from "./common.mjs";
import {
  findNativeBinary,
  nativeTarget,
} from "./node_modules/agent-browser/scripts/native.js";
const exec = promisify(execFile);
export const origin = "https://cica-mcp.grassinside.com";
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
export async function api(cfg, path, options = {}) {
  const r = await fetch(origin + path, {
    ...options,
    headers: { ...options.headers, Authorization: "Bearer " + cfg.token },
  });
  if (!r.ok) throw Error("CICA service " + r.status);
  return r;
}
export function localTools(cfg) {
  const m = server(),
    dir = accountDir(cfg.email),
    work = join(dir, "documents");
  const { owner, profile } = browserArgs(cfg);
  const cli = async (...args) => {
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
        "--account",
        cfg.eip,
        ...args,
      ],
      { maxBuffer: 12e6, timeout: 120000, windowsHide: true },
    );
    return JSON.parse(stdout);
  };
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
    "Verify identity, discover form schemas, reviewers, existing requests or drafts.",
    {
      view: z.enum([
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
        return cli(
          view,
          ...(view === "workflow" ? [] : ["--template"]),
          String(id),
        );
      }
      return cli(view);
    },
  );
  tool(
    m,
    "work_file_download",
    "Download an employee-accessible work document to the private local documents folder for editing, NDA generation or EIP attachment.",
    {
      drive: z.string(),
      item: z.string(),
      filename: docname,
      pdf: z.boolean().default(false),
    },
    async ({ drive, item, filename, pdf }) => {
      await mkdir(work, { recursive: true, mode: 0o700 });
      const r = await api(
        cfg,
        "/files/" +
          encodeURIComponent(drive) +
          "/" +
          encodeURIComponent(item) +
          (pdf ? "?format=pdf" : ""),
      );
      const path = join(work, filename);
      await writeFile(path, Buffer.from(await r.arrayBuffer()), {
        flag: "wx",
        mode: 0o600,
      });
      return { path };
    },
  );
  tool(
    m,
    "work_file_publish",
    "Create a new file in an employee-accessible folder using a prepared local document.",
    { drive: z.string(), folder: z.string(), filename: docname },
    async ({ drive, folder, filename }) =>
      (
        await api(
          cfg,
          "/files/" +
            encodeURIComponent(drive) +
            "/" +
            encodeURIComponent(folder),
          {
            method: "POST",
            headers: {
              "Content-Type": "application/octet-stream",
              "X-File-Name": encodeURIComponent(filename),
            },
            body: await readFile(await document(filename)),
          },
        )
      ).json(),
  );
  tool(
    m,
    "eip_prepare",
    "Prepare a typed EIP draft from request_schema. Employee identity is taken from Netask. Supply evidence files by name from the documents folder. Inspect prior requests and confirm company procedure and routing first.",
    {
      request: z.record(z.string(), z.unknown()),
      attachments: z
        .array(
          z.object({ filename: docname, roles: z.array(z.string()).min(1) }),
        )
        .min(1),
    },
    async ({ request, attachments }) => {
      if (!Array.isArray(request.route) || !request.route.length)
        throw Error(
          "Supply the employee-specific route from current form examples or the approved company procedure",
        );
      const status = await cli("status");
      if (status.identity?.account !== cfg.eip)
        throw Error("Sign in to your Netask account");
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
          ...(step === "draft-save-status"
            ? [path, "--attempt", attempt]
            : [attempt]),
        );
      }
      const args = [step, path];
      if (
        [
          "draft-fill",
          "draft-attachments",
          "draft-upload",
          "draft-save",
        ].includes(step)
      ) {
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
if (
  process.argv[1] &&
  resolve(process.argv[1]) === fileURLToPath(import.meta.url)
) {
  const cfg = JSON.parse(await readFile(cfgPath));
  await localTools(cfg).connect(new StdioServerTransport());
}
