import { chromium } from "playwright";
import { join, basename } from "node:path";
import { mkdir, readFile, writeFile, rename } from "node:fs/promises";
import { createHash } from "node:crypto";

export async function mailbox(dir, email, action) {
  const context = await chromium.launchPersistentContext(join(dir, "outlook"), { headless: false });
  try {
    const page = context.pages()[0] ?? (await context.newPage());
    page.setDefaultTimeout(20000);
    await page.goto("https://outlook.office.com/mail/", { waitUntil: "domcontentloaded" });
    await page
      .getByRole("treeitem", { name: email, exact: true })
      .first()
      .waitFor({ timeout: 180000 });
    return await action(page);
  } finally {
    await context.close();
  }
}

export async function compose(page, args, attachments) {
  await page.getByRole("button", { name: "New mail", exact: true }).first().click();
  const to = page.locator('[contenteditable="true"][aria-label="To"]');
  for (const address of args.to) {
    await to.fill(address);
    await to.press("Enter");
  }
  await page.getByRole("textbox", { name: "Subject", exact: true }).fill(args.subject);
  const body = page.getByRole("textbox", { name: "Message body" });
  await body.press("Control+Home");
  await page.keyboard.insertText(args.body + "\n\n");
  if (attachments.length) {
    const chooser = page.waitForEvent("filechooser");
    await page.getByRole("button", { name: "Attach file", exact: true }).first().click();
    const browse = page.getByText(/browse this computer/i).first();
    await browse.click();
    await (await chooser).setFiles(attachments);
    for (const path of attachments)
      await page.getByText(basename(path), { exact: true }).first().waitFor();
  }
  await page
    .getByRole("button", { name: "Editing " + args.subject + " Close", exact: true })
    .click();
  await page
    .getByRole("treeitem", { name: /^Drafts\b/i })
    .last()
    .click();
  const entry = page.getByRole("option").filter({ hasText: args.subject }).first();
  await entry.waitFor();
  await entry.click();
  if (
    (await page.getByRole("textbox", { name: "Subject", exact: true }).inputValue()) !==
      args.subject ||
    !(await body.innerText()).includes(args.body.trim())
  )
    throw Error("Check the draft in Outlook before retrying");
  return { folder: "Drafts", subject: args.subject, isDraft: true };
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
