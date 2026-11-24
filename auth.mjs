import { createHash, randomBytes } from "node:crypto";
import { DatabaseSync } from "node:sqlite";

const digest = (s) => createHash("sha256").update(s).digest("hex");
export const secret = () => randomBytes(32).toString("base64url");
export class Accounts {
  constructor(file) {
    this.db = new DatabaseSync(file);
    this.db.exec(
      "CREATE TABLE IF NOT EXISTS accounts (key TEXT PRIMARY KEY, oid TEXT NOT NULL, email TEXT NOT NULL, refresh TEXT NOT NULL)",
    );
  }
  issue(me, refresh) {
    const token = secret();
    this.db
      .prepare("INSERT INTO accounts VALUES (?,?,?,?)")
      .run(digest(token), me.id, me.userPrincipalName.toLowerCase(), refresh);
    return token;
  }
  lookup(token) {
    return (
      token &&
      this.db.prepare("SELECT * FROM accounts WHERE key=?").get(digest(token))
    );
  }
  rotate(key, refresh) {
    this.db
      .prepare("UPDATE accounts SET refresh=? WHERE key=?")
      .run(refresh, key);
  }
  revoke(key) {
    this.db.prepare("DELETE FROM accounts WHERE key=?").run(key);
  }
}
export const scopes =
  "offline_access User.Read User.ReadBasic.All Mail.ReadWrite Contacts.Read People.Read Files.ReadWrite.All";
export async function microsoft(tenant, operation, fields) {
  const r = await fetch(
    "https://login.microsoftonline.com/" +
      encodeURIComponent(tenant) +
      "/oauth2/v2.0/" +
      operation,
    {
      method: "POST",
      body: new URLSearchParams(fields),
      signal: AbortSignal.timeout(15000),
    },
  );
  const body = await r.json();
  if (!r.ok) throw new Error(body.error || "Microsoft sign-in failed");
  return body;
}
export async function graph(token, path, options = {}) {
  const r = await fetch("https://graph.microsoft.com/v1.0" + path, {
    ...options,
    headers: { Authorization: "Bearer " + token, ...options.headers },
    signal: AbortSignal.timeout(30000),
  });
  if (!r.ok)
    throw new Error(
      "Microsoft Graph " +
        r.status +
        " " +
        ((await r.json().catch(() => ({}))).error?.code || ""),
    );
  return r;
}
export async function identity(token) {
  return (
    await graph(token, "/me?$select=id,displayName,userPrincipalName,mail")
  ).json();
}
export function verifyAccount(account, me, allowed) {
  if (
    me.id !== account.oid ||
    me.userPrincipalName.toLowerCase() !== account.email ||
    !allowed.has(account.email)
  )
    throw Error("Employee identity mismatch");
}
