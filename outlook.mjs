import { join, basename, extname } from "node:path";
import { fileURLToPath } from "node:url";
import { mkdir, readFile, writeFile, rename, stat } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";

// The Outlook command is an employee-configured API CLI. A .mjs entry point
// works on Windows without relying on cmd.exe shims.
export const powershell = process.platform === "win32" ? "powershell.exe" : "pwsh";
function outlook(cfg, args, input) {
  const js = extname(cfg.outlookCli) === ".mjs";
  return command(
    js ? process.execPath : cfg.outlookCli,
    js ? [cfg.outlookCli, ...args] : args,
    input,
    { ...process.env, OUTLOOK_ACCOUNT_EMAIL: cfg.email },
  );
}
function command(file, args, input, env = process.env) {
  return new Promise((resolve, reject) => {
    const child = execFile(
      file,
      args,
      { env, timeout: 120000, maxBuffer: 16e6, windowsHide: true },
      (error, stdout, stderr) => {
        if (error) return reject(new Error(stderr.trim() || error.message));
        try {
          resolve(JSON.parse(stdout.replace(/^\uFEFF/, "")));
        } catch {
          reject(new Error("Mailbox command returned invalid JSON"));
        }
      },
    );
    child.stdin.on("error", () => {}); // execFile reports early process exits.
    child.stdin.end(input === undefined ? undefined : JSON.stringify(input));
  });
}
const script = fileURLToPath(new URL("./graph-mail.ps1", import.meta.url));
function graph(input) {
  return command(powershell, ["-NoProfile", "-NonInteractive", "-File", script], input);
}
export async function verifyMailbox(cfg) {
  const actual = cfg.outlookCli
    ? (await outlook(cfg, ["context"])).user?.email
    : (await graph({ account: cfg.email })).account;
  if (actual?.toLowerCase() !== cfg.email)
    throw Error("Sign in to the configured employee mailbox: " + cfg.email);
}
export async function prepareAttachments(paths) {
  const files = [];
  let size = 0;
  for (const path of paths) {
    const info = await stat(path);
    size += info.size;
    if (!info.isFile() || size > 2800000) throw Error("Use regular files totaling at most 2.8 MB.");
    files.push({ path, name: basename(path), bytes: (await readFile(path)).toString("base64") });
  }
  return files;
}
export async function compose(cfg, args, files) {
  let id, saved;
  if (cfg.outlookCli) {
    await verifyMailbox(cfg);
    const result = await outlook(cfg, ["draft", ...files.flatMap((f) => ["--attach", f.path])], {
      Subject: args.subject,
      ToRecipients: args.to.map((Address) => ({ EmailAddress: { Address } })),
      Body: { ContentType: "Text", Content: args.body },
    });
    id = result.DraftId;
    const draft = await outlook(cfg, ["draft-read", id]);
    const attachments = files.length ? (await outlook(cfg, ["draft-attachments", id])).value : [];
    saved = {
      id: draft.Id,
      isDraft: draft.IsDraft,
      subject: draft.Subject,
      to: draft.ToRecipients.map((r) => r.EmailAddress.Address),
      webLink: draft.WebLink,
      attachments: attachments.map((f) => ({ name: f.Name, contentBytes: f.ContentBytes })),
    };
  } else {
    const attachments = files.map((f) => ({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: f.name,
      contentBytes: f.bytes,
    }));
    const message = {
      subject: args.subject,
      body: { contentType: "Text", content: args.body },
      toRecipients: args.to.map((address) => ({ emailAddress: { address } })),
      attachments,
    };
    const graphDraft = await graph({ account: cfg.email, message });
    id = graphDraft.createdId;
    const d = graphDraft.saved;
    if (d.body.content.replace(/\r\n/g, "\n").trim() !== args.body.replace(/\r\n/g, "\n").trim())
      throw Error("Saved draft body differs; check Drafts.");
    saved = {
      id: d.id,
      isDraft: d.isDraft,
      subject: d.subject,
      to: d.toRecipients.map((r) => r.emailAddress.address),
      webLink: d.webLink,
      attachments: d.attachments,
    };
  }
  const normalized = (a) =>
    a
      .map((s) => s.toLowerCase())
      .sort()
      .join("\n");
  if (
    !id ||
    saved.id !== id ||
    saved.isDraft !== true ||
    saved.subject !== args.subject ||
    normalized(saved.to) !== normalized(args.to) ||
    files.some(
      (f) => !saved.attachments.some((a) => a.name === f.name && a.contentBytes === f.bytes),
    )
  )
    throw Error("Draft readback differs; check Outlook Drafts.");
  return { id, folder: "Drafts", subject: saved.subject, isDraft: true, webLink: saved.webLink };
}
export async function once(dir, args, action) {
  const folder = join(dir, "draft-receipts");
  await mkdir(folder, { recursive: true, mode: 0o700 });
  const path = join(folder, args.request_id);
  const digest = createHash("sha256").update(JSON.stringify(args)).digest("hex");
  try {
    await writeFile(path, JSON.stringify({ digest }), { flag: "wx", mode: 0o600 });
  } catch (e) {
    if (e.code !== "EEXIST") throw e;
    const saved = JSON.parse(await readFile(path, "utf8"));
    if (saved.digest !== digest) throw Error("This request ID belongs to another draft");
    if (saved.result) return saved.result;
    throw Error("Check Outlook Drafts for this interrupted request before creating another");
  }
  const result = await action();
  await writeFile(path + ".complete", JSON.stringify({ digest, result }), { mode: 0o600 });
  await rename(path + ".complete", path);
  return result;
}
