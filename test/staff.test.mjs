import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { Accounts, verifyAccount } from "../auth.mjs";
import { employeeTools, addPerson } from "../staff.mjs";
import { localTools, accountDir } from "../local.mjs";

async function connect(t, server) {
  const client = new Client({ name: "acceptance", version: "1" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(a);
  await client.connect(b);
  t.after(() => client.close());
  return client;
}
test("tokens are isolated, revocable and bound to verified Microsoft object and email", () => {
  const accounts = new Accounts(":memory:");
  const me = { id: "one", userPrincipalName: "one@example.test" };
  const other = { id: "two", userPrincipalName: "two@example.test" };
  const one = accounts.issue(me, "refresh1"),
    two = accounts.issue(other, "refresh2");
  assert.equal(accounts.lookup(one).refresh, "refresh1");
  assert.equal(accounts.lookup(two).refresh, "refresh2");
  assert.throws(() =>
    verifyAccount(
      accounts.lookup(one),
      other,
      new Set([other.userPrincipalName]),
    ),
  );
  assert.throws(() => verifyAccount(accounts.lookup(one), me, new Set()));
  verifyAccount(accounts.lookup(one), me, new Set([me.userPrincipalName]));
  accounts.revoke(accounts.lookup(one).key);
  assert.equal(accounts.lookup(one), undefined);
  assert.ok(accounts.lookup(two));
});
test("portal membership is additive, case-insensitive and idempotent", () => {
  const dir = mkdtempSync(join(tmpdir(), "staff-test-")),
    path = join(dir, "roster");
  try {
    writeFileSync(
      path,
      JSON.stringify({ name: "Owner", people: [], site: "example" }),
    );
    const person = { name: "Pilot", email: "pilot@example.test" };
    assert.equal(addPerson(path, person).status, "enrolled");
    assert.equal(addPerson(path, person).status, "already enrolled");
    assert.throws(() =>
      addPerson(path, { name: "PILOT", email: "two@example.test" }),
    );
    const saved = JSON.parse(readFileSync(path));
    assert.equal(saved.site, "example");
    assert.equal(saved.people.length, 1);
    assert.equal(saved.people[0].selected, false);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
test("MCP round trip drafts and verifies in the selected employee mailbox; retry returns existing draft", async (t) => {
  let stored,
    created = 0;
  const me = { id: "user-a", userPrincipalName: "a@example.test" };
  t.mock.method(globalThis, "fetch", async (url, options) => {
    assert.equal(options.headers.Authorization, "Bearer employee-a");
    const send = (o) => Response.json(o);
    if (new URL(url).searchParams.has("$filter"))
      return send({ value: stored ? [stored] : [] });
    if (options.method === "POST") {
      assert.equal(url, "https://graph.microsoft.com/v1.0/me/messages");
      ++created;
      const body = JSON.parse(options.body);
      stored = {
        id: "draft-1",
        isDraft: true,
        hasAttachments: false,
        subject: body.subject,
      };
      assert.equal(
        body.toRecipients[0].emailAddress.address,
        "supplier@example.test",
      );
      return send(stored);
    }
    assert.match(url, /me\/messages\/draft-1/);
    return send(stored);
  });
  const client = await connect(t, employeeTools(me, "employee-a", {}));
  assert.ok(
    (await client.listTools()).tools.some((t) => t.name === "find_person"),
  );
  const args = {
    request_id: "27e0a4f8-fdc4-4b22-9c39-c69d636be844",
    to: ["supplier@example.test"],
    subject: "NDA review",
    body: "Please review the attached agreement.",
  };
  for (let i = 0; i < 2; i++) {
    const r = await client.callTool({ name: "outlook_draft", arguments: args });
    assert.equal(JSON.parse(r.content[0].text).isDraft, true);
  }
  assert.equal(created, 1);
});
test("Netask companion advertises draft steps and uses distinct directories per employee", async (t) => {
  const client = await connect(t, localTools({ email: "pilot@example.test" }));
  assert.notEqual(
    accountDir("pilot@example.test"),
    accountDir("b@example.test"),
  );
  assert.deepEqual(
    (await client.listTools()).tools.map((t) => t.name).sort(),
    [
      "eip_inspect",
      "work_file_download",
      "work_file_publish",
      "eip_prepare",
      "eip_draft_step",
    ].sort(),
  );
});
