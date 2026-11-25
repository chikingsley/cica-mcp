import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, rm } from "node:fs/promises";
import { join, delimiter } from "node:path";
import { tmpdir } from "node:os";
import { spawnSync } from "node:child_process";
import { compose, verifyMailbox, prepareAttachments, powershell } from "../outlook.mjs";
const available = spawnSync(powershell, ["-NoProfile", "-Command", "exit 0"]).status === 0;
test(
  "Graph PowerShell verifies delegated identity, creates a draft and checks its saved bytes",
  { skip: !available },
  async (t) => {
    const root = await mkdtemp(join(tmpdir(), "staff-graph-"));
    t.after(() => rm(root, { recursive: true, force: true }));
    const dir = join(root, "Microsoft.Graph.Authentication");
    await mkdir(dir);
    await writeFile(
      join(dir, "Microsoft.Graph.Authentication.psm1"),
      `
function Get-MgContext { @{AuthType='Delegated';Scopes=@('Mail.ReadWrite','User.Read')} }
function Invoke-MgGraphRequest($Method,$Uri,$Body,$Headers,$ContentType) {
  if ($Uri -eq 'https://graph.microsoft.com/v1.0/me?$select=id,mail,userPrincipalName' -and $Method -eq 'GET') {return @{id='employee';mail='employee@cicahuntek.us'}}
  if ($Uri -eq 'https://graph.microsoft.com/v1.0/users/employee/mailFolders/drafts/messages' -and $Method -eq 'POST') {
    $script:d = $Body | ConvertFrom-Json
    return @{id='saved/1'}
  }
  if ($Uri -like 'https://graph.microsoft.com/v1.0/users/employee/messages/saved%2F1?*$select=*' -and $Method -eq 'GET' -and $Headers.Prefer -eq 'outlook.body-content-type="text"') {
    return @{id='saved/1';subject=$script:d.subject;isDraft=$true;body=$script:d.body;toRecipients=$script:d.toRecipients;webLink='review'}
  }
  if ($Uri -eq 'https://graph.microsoft.com/v1.0/users/employee/messages/saved%2F1/attachments' -and $Method -eq 'GET') {
    return @{value=$script:d.attachments}
  }
  throw "Unexpected Graph call: $Method $Uri"
}
`,
    );
    const previous = process.env.PSModulePath;
    process.env.PSModulePath = [root, previous].filter(Boolean).join(delimiter);
    t.after(() => {
      if (previous === undefined) delete process.env.PSModulePath;
      else process.env.PSModulePath = previous;
    });
    const cfg = { email: "employee@cicahuntek.us" };
    await verifyMailbox(cfg);
    await assert.rejects(verifyMailbox({ email: "else@cicahuntek.us" }), /differs/);
    const file = join(root, "文件.txt");
    await writeFile(file, "证据");
    const r = await compose(
      cfg,
      { to: [cfg.email], subject: "NDA 草稿", body: "附件\n合同" },
      await prepareAttachments([file]),
    );
    assert.equal(r.id, "saved/1");
  },
);
