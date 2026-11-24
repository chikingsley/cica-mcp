import express from "express";
import { existsSync, readFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import {
  Accounts,
  secret,
  scopes,
  microsoft,
  graph,
  identity,
  verifyAccount,
} from "./auth.mjs";
import { employeeTools } from "./staff.mjs";

process.umask(0o077);
export function createApp(
  root,
  { exchange = microsoft, identify = identity } = {},
) {
  const cfg = JSON.parse(readFileSync(join(root, "config.json"), "utf8"));
  const allowed = new Set(cfg.employees.map((s) => s.toLowerCase()));
  const accounts = new Accounts(join(root, "staff.sqlite"));
  const pending = new Map(),
    cache = new Map(),
    locks = new Map();
  const app = express();
  app.disable("x-powered-by");
  app.use((req, res, next) => {
    res.set("Cache-Control", "no-store");
    if (req.headers.origin && req.headers.origin !== cfg.origin)
      return res.sendStatus(403);
    next();
  });
  const readJson = express.json({ limit: "150kb" });
  app.get("/", (_, res) =>
    res
      .type("text")
      .send(
        "CICA staff MCP\nSetup: " +
          cfg.repo +
          "\nEndpoint: " +
          cfg.origin +
          "/mcp\nSign in using your own company account.",
      ),
  );
  app.get("/health", (_, res) =>
    res.json({ service: "cica-mcp", ready: true }),
  );
  app.post("/enroll", readJson, async (req, res) => {
    const now = Date.now();
    for (const [k, v] of pending) if (now >= v.expires) pending.delete(k);
    if (pending.size >= 16) return res.sendStatus(429);
    const email = String(req.body?.email || "").toLowerCase();
    // Identical response for unknown addresses and invalid requests. Microsoft
    // verifies the actual account on completion before any capability is issued.
    if (!allowed.has(email))
      return res.status(403).json({ error: "Employee enrollment is required" });
    const key = secret();
    pending.set(key, { expires: now + 60000 });
    try {
      const dev = await exchange(cfg.tenant, "devicecode", {
        client_id: cfg.client,
        scope: scopes,
      });
      pending.set(key, {
        email,
        device: dev.device_code,
        expires: now + dev.expires_in * 1000,
        next: now + (dev.interval || 5) * 1000,
        interval: (dev.interval || 5) * 1000,
      });
      res.json({
        key,
        code: dev.user_code,
        url: dev.verification_uri,
        interval: dev.interval || 5,
        expires_in: dev.expires_in,
      });
    } catch (e) {
      pending.delete(key);
      throw e;
    }
  });
  app.post("/enroll/finish", readJson, async (req, res) => {
    const key = req.body?.key,
      p = pending.get(key);
    if (!p?.device || Date.now() >= p.expires) return res.sendStatus(410);
    if (p.next > Date.now()) return res.sendStatus(202);
    p.next = Infinity;
    try {
      const t = await exchange(cfg.tenant, "token", {
        client_id: cfg.client,
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        device_code: p.device,
      });
      const me = await identify(t.access_token);
      if (me.userPrincipalName.toLowerCase() !== p.email || !t.refresh_token)
        throw Error("Employee identity mismatch");
      res.json({ token: accounts.issue(me, t.refresh_token), email: p.email });
    } catch (e) {
      if (["authorization_pending", "slow_down"].includes(e.message)) {
        if (e.message === "slow_down") p.interval += 5000;
        p.next = Date.now() + p.interval;
        return res.sendStatus(202);
      }
      throw e;
    } finally {
      if (p.next === Infinity) pending.delete(key);
    }
  });
  app.use(async (req, res, next) => {
    const header = req.get("Authorization") || "";
    const a = /^Bearer ([\w-]{43})$/.exec(header);
    const account = a && accounts.lookup(a[1]);
    if (!account || !allowed.has(account.email))
      return res.status(401).set("WWW-Authenticate", "Bearer").end();
    res.locals.account = account;
    next();
  });
  async function session(account) {
    let entry = cache.get(account.key);
    if (entry && entry.until > Date.now()) return entry.p;
    const p = (async () => {
      const t = await exchange(cfg.tenant, "token", {
        client_id: cfg.client,
        grant_type: "refresh_token",
        refresh_token: account.refresh,
        scope: scopes,
      });
      const me = await identify(t.access_token);
      verifyAccount(account, me, allowed);
      if (t.refresh_token) accounts.rotate(account.key, t.refresh_token);
      return { me, token: t.access_token };
    })();
    cache.set(account.key, { until: Date.now() + 120000, p });
    try {
      return await p;
    } catch (e) {
      cache.delete(account.key);
      throw e;
    }
  }
  app.post("/disconnect", (req, res) => {
    const { key } = res.locals.account;
    accounts.revoke(key);
    cache.delete(key);
    res.sendStatus(204);
  });
  app.get("/engine/:name", (req, res) => {
    const file = cfg.engine;
    if (req.params.name !== file || !existsSync(join(root, file)))
      return res.sendStatus(404);
    res.sendFile(join(root, file));
  });
  app.get("/info", (req, res) =>
    res.json({
      email: res.locals.account.email,
      engine: cfg.engine,
      sha256: cfg.engineHash,
      eip: cfg.netask[res.locals.account.email],
    }),
  );
  const itemPath = (d, i) =>
    "/drives/" + encodeURIComponent(d) + "/items/" + encodeURIComponent(i);
  app.get("/files/:drive/:item", async (req, res) => {
    const { token } = await session(res.locals.account),
      path = itemPath(req.params.drive, req.params.item);
    const meta = await (await graph(token, path + "?$select=size,file")).json();
    if (!meta.file || meta.size > 20000000) return res.sendStatus(413);
    res
      .type("application/octet-stream")
      .send(
        Buffer.from(
          await (
            await graph(
              token,
              path +
                "/content" +
                (req.query.format === "pdf" ? "?format=pdf" : ""),
            )
          ).arrayBuffer(),
        ),
      );
  });
  app.post(
    "/files/:drive/:folder",
    express.raw({ type: "application/octet-stream", limit: "20mb" }),
    async (req, res) => {
      const { token } = await session(res.locals.account),
        name = decodeURIComponent(req.get("X-File-Name") || "");
      if (
        !name ||
        !/^[\p{L}\p{N} (._)-]{1,160}$/u.test(name) ||
        !req.body?.length
      )
        return res.sendStatus(400);
      const path =
        itemPath(req.params.drive, req.params.folder) +
        ":/" +
        encodeURIComponent(name) +
        ":/createUploadSession";
      const sessionData = await (
        await graph(token, path, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            item: { "@microsoft.graph.conflictBehavior": "fail", name },
          }),
        })
      ).json();
      const bytes = req.body;
      let result;
      for (let start = 0; start < bytes.length; start += 3276800) {
        const part = bytes.subarray(start, start + 3276800);
        result = await fetch(sessionData.uploadUrl, {
          method: "PUT",
          headers: {
            "Content-Range": `bytes ${start}-${start + part.length - 1}/${bytes.length}`,
          },
          body: part,
          signal: AbortSignal.timeout(30000),
        });
        if (!result.ok)
          throw Error("SharePoint upload failed: " + result.status);
      }
      const item = await result.json();
      const check = await (
        await graph(
          token,
          itemPath(req.params.drive, item.id) + "?$select=id,name,size,webUrl",
        )
      ).json();
      if (check.size !== bytes.length || check.name !== name)
        throw Error("File verification failed");
      res.json(check);
    },
  );
  app.post("/mcp", readJson, async (req, res) => {
    const { me, token } = await session(res.locals.account);
    // Serialize complete write calls per employee across concurrent MCP requests.
    const key = me.id,
      prior = locks.get(key) || Promise.resolve();
    let release;
    const done = new Promise((r) => {
      release = r;
    });
    locks.set(key, done);
    await prior;
    try {
      const mcp = employeeTools(me, token, cfg);
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
        enableJsonResponse: true,
      });
      res.on("close", () => {
        void mcp.close();
      });
      await mcp.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } finally {
      release();
      if (locks.get(key) === done) locks.delete(key);
    }
  });
  app.use((err, req, res, next) => {
    if (!res.headersSent)
      res.status(400).json({ error: String(err.message).slice(0, 200) });
  });
  return { app, port: cfg.port };
}
if (
  process.argv[1] &&
  resolve(process.argv[1]) === fileURLToPath(import.meta.url)
) {
  const { app, port } = createApp(process.env.CICA_STAFF_STATE);
  app.listen(port, "127.0.0.1");
}
