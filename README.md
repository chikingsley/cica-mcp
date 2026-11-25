# CICA MCP

Open **MCP** in the [CICA portal](https://cica.grassinside.com/mcp). Give your agent that link and ask it to connect your own employee account. The portal serves meeting records, availability and the shared EIP and NDA workflow guide. To connect those tools independently:

```sh
codex mcp add cica-portal --url https://cica.grassinside.com/mcp
```

## Employee setup

Use your work computer with Git, Node 22.13 or newer, npm, uv and Codex CLI. Download the Netask ZIP from the portal's company SharePoint link. Choose a permanent location for the checkout.

```sh
git clone https://github.com/chikingsley/cica-mcp.git
cd cica-mcp
npm ci
node setup.mjs YOUR_WORK_EMAIL YOUR_NETASK_ID "PATH_TO_DOWNLOADED_ZIP" "PATH_TO_SYNCED_COMPANY_FOLDER"
```

The synced folder is optional. If you already have the CICA account-bound Outlook API CLI, set `CICA_OUTLOOK_CLI` to its absolute executable or JavaScript .mjs entry point before setup. On Linux, setup also checks the installed `outlook` command for your identity. The command must support context, draft, draft-read and draft-attachments.

On a fresh Windows workstation, setup uses Windows PowerShell and Microsoft's Graph Authentication module. Other platforms can use PowerShell 7. Setup installs the module for your Windows or operating-system user when needed. Sign in to Microsoft with your own company account. Your tenant controls consent to the delegated User.Read and Mail.ReadWrite permissions. The tool checks the mailbox identity, creates an unsent draft by API, then reads the subject, recipients and attachments back. The Graph path also checks the saved body. Microsoft manages the local delegated sign-in. [Microsoft sign-in documentation](https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands)

Setup opens Netask for sign-in, verifies the employee again in a headless session, then closes both sessions. EIP tools use a headless browser on demand. Start a fresh EIP session with eip_inspect open, then request status. Codex starts its local MCP tool when needed. Your documents, packets and draft receipts live under your workstation account in `~/.cica-mcp`. Link your locally synced company library to copy documents in either direction. OneDrive handles the cloud transfer.

Start a new Codex chat. Try “Read the NDA workflow guide, verify my EIP identity and find the correct form. Save a test email to me with the subject ‘CICA setup test’ and verify it in Drafts.” Or ask in Chinese: “先查 NDA 流程，确认我的 EIP 账号，再给我自己建一封测试邮件草稿。” The employee reviews and sends the saved email in Outlook. A repeated draft request reuses its request ID. Check Drafts after an interrupted write.

## Updates

The portal serves its current workflow guide on every call. For a workstation tool update, close the Codex session using the tools, run `git pull --ff-only` and `npm ci` in this checkout, then restart Codex. For the upgrade from browser-based Outlook drafting, also rerun setup to verify your API connection. When the portal lists a new Netask ZIP, download it and rerun setup with the new file. Setup verifies its pinned digest. To remove the Codex registrations, run `node disconnect.mjs`. Microsoft sign-out is available through `Disconnect-MgGraph`.
