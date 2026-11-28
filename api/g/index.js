// Verkpunktar – deiling teikningar til viðskiptavinar (án innskráningar).
//   GET  /api/g/{token}       -> titill, stærð og punktar (gluggar) á teikningunni
//   GET  /api/g/{token}/img   -> mynd af teikningunni
//   POST /api/g/{token}       -> viðskiptavinur sendir merkta glugga
// Gögn eru í SharePoint (sama safn og appið): Verkpunktar/_gogn/...
// Stillingar (Azure Static Web App → Environment variables):
//   VP_TENANT_ID, VP_CLIENT_ID, VP_CLIENT_SECRET, VP_DRIVE_ID, (VP_ROOT=Verkpunktar)
//   valfrjálst: VP_WIX_CLIENT_ID, VP_WIX_FORM_ID – tilkynning í Wix-form „Gluggaval“
"use strict";

const GRAPH = "https://graph.microsoft.com/v1.0";
const ROOT = process.env.VP_ROOT || "Verkpunktar";
const DATA = "_gogn";
const SHARE_DIR = "deilingar";
const TOKEN_RE = /^[A-Za-z0-9]{20,40}$/;

let appTok = null;
async function graphToken() {
  if (appTok && appTok.exp > Date.now() + 60000) return appTok.t;
  const body = new URLSearchParams({
    client_id: process.env.VP_CLIENT_ID,
    client_secret: process.env.VP_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const r = await fetch(`https://login.microsoftonline.com/${process.env.VP_TENANT_ID}/oauth2/v2.0/token`, { method: "POST", body });
  if (!r.ok) throw new Error("Auðkenning við Microsoft mistókst (" + r.status + ")");
  const j = await r.json();
  appTok = { t: j.access_token, exp: Date.now() + j.expires_in * 1000 };
  return appTok.t;
}
async function g(path, opt = {}) {
  const t = await graphToken();
  return fetch(path.startsWith("http") ? path : GRAPH + path, { ...opt, headers: { Authorization: "Bearer " + t, ...(opt.headers || {}) } });
}
const drive = () => `/drives/${process.env.VP_DRIVE_ID}`;
const enc = (s) => s.split("/").map(encodeURIComponent).join("/");
const itemPath = (rel) => `${drive()}/root:/${enc(ROOT + "/" + rel)}`;

// Les JSON-skrá; skilar {d, eTag} eða null ef hún er ekki til.
async function readJson(rel) {
  const r = await g(itemPath(rel) + "?$select=id,eTag,@microsoft.graph.downloadUrl");
  if (r.status === 404) return null;
  if (!r.ok) throw new Error("Graph " + r.status);
  const it = await r.json();
  const c = await fetch(it["@microsoft.graph.downloadUrl"]);
  let d = {};
  try { d = await c.json(); } catch (e) {}
  return { d, eTag: it.eTag };
}
async function writeJson(rel, d, eTag) {
  const h = { "Content-Type": "application/json" };
  if (eTag) h["If-Match"] = eTag;
  return g(itemPath(rel) + ":/content", { method: "PUT", headers: h, body: JSON.stringify(d) });
}
// Uppfærir JSON-skrá með eTag-vörn (endurreynir ef appið vistaði á sama tíma).
async function updateJson(rel, fn) {
  for (let a = 0; a < 6; a++) {
    const cur = await readJson(rel);
    if (!cur) throw new Error("Skrá fannst ekki");
    const next = fn(cur.d);
    const r = await writeJson(rel, next, cur.eTag);
    if (r.ok) return next;
    if (r.status !== 412 && r.status !== 409) throw new Error("Vistun mistókst (" + r.status + ")");
    await new Promise((s) => setTimeout(s, 300 + Math.random() * 500));
  }
  throw new Error("Vistun mistókst – reyndu aftur");
}

const json = (status, body, extra = {}) => ({
  status,
  headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store", ...extra },
  body: JSON.stringify(body),
});
const clean = (s, n) => String(s == null ? "" : s).replace(/[\u0000-\u001f\u007f]/g, " ").trim().slice(0, n);
const label = (p) => clean(p.title, 40) || String(p.n || "");

async function loadShare(token) {
  const sh = await readJson(`${DATA}/${SHARE_DIR}/${token}.json`);
  if (!sh || !sh.d || sh.d.token !== token) return null;
  return sh.d;
}
function pagePins(sc, share) {
  const pg = share.page || 1;
  return Object.entries((sc && sc.pins) || {})
    .filter(([, p]) => p && !p.deleted && (p.page || 1) === pg)
    .sort((a, b) => (a[1].n || 0) - (b[1].n || 0));
}

async function notifyWix(share, info, labels) {
  const cid = process.env.VP_WIX_CLIENT_ID, fid = process.env.VP_WIX_FORM_ID;
  if (!cid || !fid) return;
  try {
    const t = await fetch("https://www.wixapis.com/oauth2/token", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ clientId: cid, grantType: "anonymous" }) });
    if (!t.ok) return;
    const tok = (await t.json()).access_token;
    const s = {
      fornafn_h9kui0: info.nafn,
      netfang_1qm3ub: info.netfang,
      verk_tgvxq7: share.title || "Teikning",
      valdir_gluggar_gyazro: labels.length ? labels.join(", ") + `\n(${labels.length} gluggar)` : "Engir gluggar merktir",
    };
    if (info.fyrirtaeki) s.fyrirtaeki_g8fmon = info.fyrirtaeki;
    if (info.athugasemd) s.athugasemd_eb4mwk = info.athugasemd;
    await fetch("https://www.wixapis.com/form-submission-service/v4/submissions", { method: "POST", headers: { "Content-Type": "application/json", Authorization: tok }, body: JSON.stringify({ submission: { formId: fid, submissions: s } }) });
  } catch (e) { /* tilkynning er aukaatriði */ }
}

module.exports = async function (context, req) {
  try {
    const token = String((req.params && req.params.token) || "");
    const part = String((req.params && req.params.part) || "");
    if (!TOKEN_RE.test(token)) { context.res = json(404, { error: "Tengill fannst ekki" }); return; }
    const share = await loadShare(token);
    if (!share) { context.res = json(404, { error: "Tengill fannst ekki" }); return; }
    if (share.active === false) { context.res = json(410, { error: "Þessum tengli hefur verið lokað" }); return; }

    if (req.method === "GET" && part === "img") {
      const r = await g(itemPath(`${DATA}/${SHARE_DIR}/${token}.jpg`) + ":/content");
      if (!r.ok) { context.res = json(404, { error: "Mynd fannst ekki" }); return; }
      const buf = Buffer.from(await r.arrayBuffer());
      context.res = { status: 200, isRaw: true, headers: { "Content-Type": "image/jpeg", "Cache-Control": "private, max-age=600" }, body: buf };
      return;
    }
    if (part) { context.res = json(404, { error: "Ekki til" }); return; }

    const scRel = `${DATA}/${share.fileId}.json`;

    if (req.method === "GET") {
      const sc = await readJson(scRel);
      const pins = pagePins(sc && sc.d, share).map(([id, p]) => ({
        id, label: label(p), x: p.x, y: p.y,
        merkt: !!(p.cust && p.cust.skipta),
      }));
      context.res = json(200, { title: share.title || "", lysing: share.lysing || "", w: share.w, h: share.h, img: `/api/g/${token}/img`, pins });
      return;
    }

    if (req.method === "POST") {
      const b = typeof req.body === "object" && req.body ? req.body : (() => { try { return JSON.parse(req.rawBody || "{}"); } catch (e) { return {}; } })();
      const info = {
        nafn: clean(b.nafn, 80), netfang: clean(b.netfang, 120), fyrirtaeki: clean(b.fyrirtaeki, 120), athugasemd: clean(b.athugasemd, 2000),
      };
      if (!info.nafn || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(info.netfang)) { context.res = json(400, { error: "Vantar nafn og gilt netfang" }); return; }
      const sel = new Set(Array.isArray(b.sel) ? b.sel.map(String).slice(0, 500) : []);
      const now = Date.now();
      let labels = [];
      await updateJson(scRel, (d) => {
        d.pins = d.pins || {};
        labels = [];
        for (const [id, p] of pagePins(d, share)) {
          const was = !!(p.cust && p.cust.skipta);
          if (sel.has(id)) {
            labels.push(label(p));
            d.pins[id] = { ...p, cust: { skipta: true, nafn: info.nafn, netfang: info.netfang, at: now, token }, updatedAt: now, updatedBy: "Viðskiptavinur: " + info.nafn };
          } else if (was) {
            d.pins[id] = { ...p, cust: { ...p.cust, skipta: false, at: now }, updatedAt: now, updatedBy: "Viðskiptavinur: " + info.nafn };
          }
        }
        return d;
      });
      await updateJson(`${DATA}/${SHARE_DIR}/${token}.json`, (d) => {
        d.svor = (d.svor || []).concat([{ ...info, gluggar: labels, at: now }]).slice(-50);
        d.lastAt = now;
        return d;
      });
      await notifyWix(share, info, labels);
      context.res = json(200, { ok: true, fjoldi: labels.length, gluggar: labels });
      return;
    }
    context.res = json(405, { error: "Óleyfileg aðgerð" });
  } catch (e) {
    context.log.error(e);
    context.res = json(500, { error: "Villa kom upp. Reyndu aftur síðar." });
  }
};
