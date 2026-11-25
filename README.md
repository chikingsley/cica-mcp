# CICA MCP

Open **MCP** in the [CICA portal](https://cica.grassinside.com/mcp). Its meetings, minutes, transcript and availability tools run in the existing portal. An employee can connect those tools independently:

```sh
codex mcp add cica-portal --url https://cica.grassinside.com/mcp
```

## Employee setup

Give your Codex agent the portal link and ask it to install the workstation tools for your own account. Use a workstation with a desktop browser, Git, Node 22.13 or newer, npm, uv and Codex CLI. Download the Netask ZIP from the company SharePoint link on the portal. Keep this checkout in a permanent location.

```sh
git clone https://github.com/chikingsley/cica-mcp.git
cd cica-mcp
npm ci
node setup.mjs YOUR_WORK_EMAIL YOUR_NETASK_ID "PATH_TO_DOWNLOADED_ZIP" "PATH_TO_SYNCED_COMPANY_FOLDER"
```

The synced folder argument is optional. Sign in to Outlook and Netask in the windows setup opens. Use your own work identity and complete your usual MFA. Setup checks your Outlook mailbox name. Your browser profiles, prepared packets, documents and draft receipts live under your workstation account in `~/.cica-mcp`. Codex starts the workstation tools when it needs them.

In a fresh Codex chat, ask: “Check my EIP status and show me the applicant. Find the NDA form and reviewers. Draft an email to me with the subject ‘CICA setup test’ and verify it in my Outlook Drafts.” The mail tool uses Outlook on the web and leaves the message in Drafts. Use the same request ID to retry a draft. For an interrupted request, check Drafts first. You can place prepared files in the private documents folder printed by setup, or copy files from a locally synced company library.

The portal endpoint offers meeting reads. The workstation provides Outlook drafts, EIP inspection and draft preparation, and copies to or from your chosen synced folder. Portal roster administration stays with the portal operator. Microsoft and Netask enforce their own sign-in policies. Workstation setup uses interactive sign-in and requests no Graph application permissions.

To remove the Codex registrations, run `node disconnect.mjs`. It prints the location of your local profiles and documents.
