import { z } from "zod";
import { readFileSync, writeFileSync, renameSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { graph } from "./auth.mjs";
import { server, tool } from "./common.mjs";
const id = z.string().min(1).max(500);
const esc = encodeURIComponent;
const file = (drive, item) => "/drives/" + esc(drive) + "/items/" + esc(item);
const jsonBody = (obj) => ({
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(obj),
});
export function addPerson(path, person) {
  const config = JSON.parse(readFileSync(path, "utf8"));
  const people = config.people;
  const existing = people.find((p) => p.email.toLowerCase() === person.email);
  if (existing) return { status: "already enrolled", name: existing.name };
  if (
    people.length >= 20 ||
    [config.name, ...people.map((p) => p.name)].some(
      (n) => n?.toLowerCase() === person.name.toLowerCase(),
    )
  )
    throw Error("Roster is full or the label is already in use");
  people.push({ ...person, selected: false });
  const tmp = path + "." + process.pid + ".tmp";
  writeFileSync(tmp, JSON.stringify(config, null, 2) + "\n", {
    mode: 0o600,
    flag: "wx",
  });
  renameSync(tmp, path);
  return {
    status: "enrolled",
    name: person.name,
    coverage: "Await the calendar sync and verify availability",
  };
}
export function employeeTools(me, token, config) {
  const m = server();
  const get = async (path) => (await graph(token, path)).json();
  tool(
    m,
    "whoami",
    "Verify the work mailbox and account for this connection.",
    {},
    () => me,
  );
  tool(
    m,
    "find_person",
    "Search people available to the signed-in employee.",
    { name: z.string().min(2).max(120) },
    async ({ name }) => {
      const query = new URLSearchParams({
        $search: '"' + name.replace(/["\\]/g, "") + '"',
        $top: "20",
        $select: "displayName,scoredEmailAddresses,companyName",
      });
      const prefix = name.replaceAll("'", "''");
      const contacts = new URLSearchParams({
        $filter: "startswith(displayName,'" + prefix + "')",
        $select: "displayName,emailAddresses,companyName",
        $top: "30",
      });
      const [suggested, saved] = await Promise.all([
        get("/me/people?" + query),
        get("/me/contacts?" + contacts),
      ]);
      return { suggested: suggested.value, contacts: saved.value };
    },
  );
  tool(
    m,
    "find_mail",
    "Search this employee’s own mailbox for prior correspondence.",
    { query: z.string().min(2).max(200) },
    async ({ query }) => {
      const params = new URLSearchParams({
        $search: '"' + query.replace(/["\\]/g, "") + '"',
        $top: "25",
        $select:
          "id,subject,from,toRecipients,receivedDateTime,isDraft,webLink",
      });
      return (await get("/me/messages?" + params)).value;
    },
  );
  tool(
    m,
    "read_mail",
    "Read the full text of a selected message in this mailbox.",
    { id },
    async ({ id }) =>
      (
        await graph(
          token,
          "/me/messages/" + esc(id) + "?$select=subject,body,isDraft,webLink",
          { headers: { Prefer: 'outlook.body-content-type="text"' } },
        )
      ).json(),
  );
  tool(
    m,
    "find_work_file",
    "Search the employee’s OneDrive or an accessible SharePoint drive. Returns file and drive IDs.",
    { query: z.string().min(2).max(160), drive: id.optional() },
    async ({ query, drive }) => {
      const path =
        (drive ? "/drives/" + esc(drive) : "/me/drive") +
        "/root/search(q='" +
        esc(query.replaceAll("'", "''")) +
        "')?$top=40&$select=id,name,webUrl,parentReference,file";
      return (await get(path)).value;
    },
  );
  tool(
    m,
    "document_libraries",
    "List document libraries on the configured company SharePoint site using the employee’s permissions.",
    {},
    async () =>
      (
        await get(
          "/sites/" + esc(config.site) + "/drives?$select=id,name,webUrl",
        )
      ).value,
  );
  tool(
    m,
    "outlook_draft",
    "Save an unsent email in your own Drafts. Review existing mail first. Supply a unique request ID and reuse it on retries. Optional file attachments come from employee-accessible Microsoft drives.",
    {
      request_id: z.uuid(),
      to: z.array(z.email()).min(1).max(25),
      cc: z.array(z.email()).max(25).default([]),
      subject: z.string().min(1).max(240),
      body: z.string().min(1).max(24000),
      attachments: z
        .array(z.object({ drive: id, item: id }))
        .max(5)
        .default([]),
    },
    async ({ request_id, to, cc, subject, body, attachments }) => {
      const marker =
        "String {1e3a381f-25a7-4b04-8b20-3949ac1c2cb1} Name StaffDraftId";
      const filter =
        "singleValueExtendedProperties/any(p:p/id eq '" +
        marker +
        "' and p/value eq '" +
        request_id +
        "')";
      const previous = await get(
        "/me/messages?" +
          new URLSearchParams({
            $filter: filter,
            $select: "id,isDraft,webLink,subject",
          }),
      );
      if (previous.value.length)
        return { existing: true, ...previous.value[0] };
      const files = [];
      let total = 0;
      for (const a of attachments) {
        const path = file(a.drive, a.item),
          meta = await get(path + "?$select=name,size,file");
        if (!meta.file || (total += meta.size) > 2700000)
          throw Error("Draft attachments exceed 2.7 MB or are folders");
        const bytes = Buffer.from(
          await (await graph(token, path + "/content")).arrayBuffer(),
        );
        files.push({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: meta.name,
          contentBytes: bytes.toString("base64"),
        });
      }
      const recipients = (a) =>
        a.map((address) => ({ emailAddress: { address } }));
      const draft = await (
        await graph(token, "/me/messages", {
          method: "POST",
          ...jsonBody({
            subject,
            body: { contentType: "Text", content: body },
            toRecipients: recipients(to),
            ccRecipients: recipients(cc),
            attachments: files,
            singleValueExtendedProperties: [{ id: marker, value: request_id }],
          }),
        })
      ).json();
      const saved = await get(
        "/me/messages/" +
          esc(draft.id) +
          "?$select=id,subject,isDraft,webLink,hasAttachments",
      );
      if (
        !saved.isDraft ||
        saved.subject !== subject ||
        Boolean(files.length) !== saved.hasAttachments
      )
        throw Error("Draft readback mismatch");
      return saved;
    },
  );
  tool(
    m,
    "portal_add_my_availability",
    "Opt your own work mailbox into the team availability portal. Use only when the employee asks to join.",
    { display_name: z.string().trim().min(2).max(70) },
    ({ display_name }) =>
      addPerson(config.roster, {
        name: display_name,
        email: me.userPrincipalName.toLowerCase(),
      }),
  );
  if (config.rosterManagers?.includes(me.userPrincipalName.toLowerCase())) {
    tool(
      m,
      "portal_add_teammate",
      "Enroll an explicitly requested company colleague in portal availability. Resolves the mailbox in the company directory first. Calendar sharing still follows Exchange permissions.",
      { work_email: z.email(), display_name: z.string().trim().min(2).max(70) },
      async ({ work_email, display_name }) => {
        const colleague = await get(
          "/users/" +
            esc(work_email) +
            "?$select=id,displayName,mail,userPrincipalName",
        );
        const email = colleague.userPrincipalName.toLowerCase();
        if (
          email !== work_email.toLowerCase() ||
          !email.endsWith("@" + config.domain)
        )
          throw Error("Select a verified company mailbox");
        return addPerson(config.roster, { name: display_name, email });
      },
    );
  }
  tool(
    m,
    "nda_workflow",
    "Get the sequence for a new NDA or a signed return.",
    {},
    () => ({
      new: "Find the counterparty and prior mail. Locate the current approved template and NDA register in SharePoint. Confirm the assigned document number. Prepare Word and PDF from the approved template and review both. Save the draft in your own mailbox with both attachments.",
      returned:
        "Find the signed agreement and verify any legal-review outcome. Use the local EIP companion to inspect request_schema, resolve reviewers, check for duplicates, prepare the appropriate signature or executed-copy packet, and save a draft. Verify applicant, attachments and saved record.",
    }),
  );
  const portal = () => import(pathToFileURL(config.portalAdapter).href);
  tool(
    m,
    "meeting_search",
    "Find CICA portal meetings by keyword or date.",
    { query: z.string().max(180), offset: z.int().nonnegative().default(0) },
    async ({ query, offset }) =>
      (await portal()).search({ query, offset, limit: 25 }),
  );
  tool(
    m,
    "meeting_read",
    "Read minutes, action items or a page of speaker-labelled transcript turns.",
    {
      meeting_id: z.string().regex(/^[a-zA-Z0-9-]+$/),
      section: z.enum(["minutes", "transcript"]),
      offset: z.int().nonnegative().default(0),
    },
    async (args) => (await portal()).meeting({ ...args, limit: 60 }),
  );
  tool(
    m,
    "meeting_availability",
    "Find common meeting windows for named teammates using portal calendar coverage.",
    {
      week: z.string().date(),
      people: z.array(z.string()).min(1),
      duration_minutes: z.int().min(15).max(240),
    },
    async (args) =>
      (await portal()).availability({
        ...args,
        buffer_minutes: 30,
        travel_minutes: 0,
      }),
  );
  return m;
}
