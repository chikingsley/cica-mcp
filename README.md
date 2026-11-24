# CICA staff MCP

Give your local Codex agent access to your work Outlook mailbox, accessible Microsoft files, the company meetings and availability portal, and your Netask EIP drafts.

## Tell your Codex agent

> Install CICA staff MCP from https://github.com/chikingsley/cica-mcp in a permanent local folder. Follow the employee setup in the README using my CICA work email. Pause for my Microsoft and Netask sign-ins, then verify both identities. Show me my Outlook mailbox and EIP applicant before preparing a draft.

## Employee setup

Use the desktop or CLI version of Codex on your workstation. Install Node.js 22.13 or newer, npm, Git, and uv. The administrator enables your work email and Netask account for the pilot. Run the following in your permanent checkout, replacing the last argument with your work email.

```sh
git clone https://github.com/chikingsley/cica-mcp.git
cd cica-mcp
npm ci
node setup.mjs your-work-email
```

Setup displays a Microsoft device code. Enter it on Microsoft's sign-in page for the setup you just started. Review the requested delegated permissions. If your organization requests administrator consent, contact your Microsoft administrator. A separate local browser then opens the company Netask site for your own login. Keep credentials and enrollment codes inside the sign-in flow.

Open a fresh Codex chat. Ask it to run `cica.whoami` and `cica-eip.eip_inspect` with view `status`. Check both returned identities. A first pilot can search for a known contact, locate a work file, save a test email to your own Drafts, and verify it in Outlook. For EIP, start with a real authorized request and review the packet before saving. Saved drafts and final submissions are separate actions.

## Daily use

- “Find the supplier and our previous correspondence. Prepare the NDA from the current approved company template and save an Outlook draft with the Word and PDF attached.”
- “Prepare an EIP draft for this signed agreement. Check existing workflows for duplicates, use my applicant identity, and show me the route, attachments and saved draft.”
- “Find last week's meeting minutes and a time the team can meet next Tuesday.”
- “Add my work calendar to the availability portal.” Authorized portal editors can also add a named coworker. The portal displays free/busy information subject to Exchange calendar sharing.

Codex uses the employee's own Microsoft permissions to locate documents. Use the local work-file tools to download the controlled template, prepare and review the document, publish a new file in the intended folder, and optionally download its Microsoft-rendered PDF. The local documents folder is printed during setup. Confirm the official agreement number and update the controlling register through an authorized existing workflow. The connector supports new-file publication and mailbox drafts. It exposes no send, signature or final EIP submission tool.

## How it connects

`cica` connects to the hosted MCP at https://cica-mcp.grassinside.com/mcp through a local standard-input/output relay. The relay reads the employee's private connection file. `cica-eip` runs on the workstation and uses a dedicated, persistent browser profile. The two registrations have one setup command. Reopen Codex after installation.

The host stores delegated refresh tokens in a private employee-keyed database. Each refreshed Graph session verifies the employee's Microsoft object ID and principal name. Employee Netask sign-in, evidence and EIP draft packets stay on the workstation. The authenticated installer downloads a pinned private Netask wheel and verifies its digest. The public repository contains the connector and tests. An employee can propose changes through a GitHub pull request. Production releases remain under the maintainer's control.

To disconnect, run `node disconnect.mjs` from the checkout for the current workstation account. To reinstall, run setup again. Provider permission grants can also be reviewed in your Microsoft account.

## Development

Run `npm ci` and `npm test`. Service deployment uses `CICA_STAFF_STATE` for the private configuration and employee credential database. Staff identities, tenant configuration, company templates and the private EIP engine are supplied by the administrator. The portal adapter uses the company's existing calendar and meeting service. On the deployment host, `./stack up`, `./stack status` and `./stack down` manage the service and route.
