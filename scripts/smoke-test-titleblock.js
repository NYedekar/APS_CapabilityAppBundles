#!/usr/bin/env node
// ════════════════════════════════════════════════════════════════════════════
// Two-mode smoke gate for AutoCADTitleBlockUpdater (UpdateTitleBlock+prod).
//
// Because the committed sample DWGs are opaque (we can't confirm which one carries
// an attributed title block without running the bundle), this gate PROBES:
//   (1) EXTRACT — run each candidate DWG with inputFile only. Assert a candidate
//       returns ok:true, mode:"extract", non-empty titleBlocks. That candidate
//       becomes the fixture. If NONE do, FAIL loudly listing what each returned
//       (→ commit a real title-block DWG and set SMOKE_SAMPLE_DWGS).
//   (2) UPDATE — pick the first non-constant attribute from the winning DWG's
//       schema, send it a new value inline via params.json, and assert
//       ok:true, mode:"update", updatedCount>0, and result.dwg > 0 bytes.
//
// Env:
//   SMOKE_ACTIVITY_ID   default UpdateTitleBlock
//   SMOKE_SAMPLE_DWGS   comma-separated DWG paths (default: test/sample.dwg,test/condo-skylight.dwg)
// Adapted from scripts/smoke-test-activity.js + smoke-test-revit-param-updater.js.
// ════════════════════════════════════════════════════════════════════════════
"use strict";

const fs    = require("fs");
const path  = require("path");
const https = require("https");

const CLIENT_ID     = required("APS_CLIENT_ID");
const CLIENT_SECRET = required("APS_CLIENT_SECRET");
const NICKNAME      = required("APS_NICKNAME");
const ACTIVITY_ID   = process.env.SMOKE_ACTIVITY_ID || "UpdateTitleBlock";
const ALIAS         = process.env.ALIAS || "prod";
const BUCKET_KEY    = (NICKNAME.toLowerCase().replace(/[^a-z0-9]/g, "") + "-smoke").slice(0, 60);
const BASE          = "developer.api.autodesk.com";

const CANDIDATES = (process.env.SMOKE_SAMPLE_DWGS ||
  "test/sample.dwg,test/condo-skylight.dwg")
  .split(",").map(s => s.trim()).filter(Boolean)
  .map(p => path.isAbsolute(p) ? p : path.join(process.cwd(), p));

function required(name) {
  const v = process.env[name];
  if (!v) { console.error(`❌ Missing env var: ${name}`); process.exit(1); }
  return v;
}

function req(method, host, p, headers, body) {
  return new Promise((resolve, reject) => {
    const r = https.request({ hostname: host, path: p, method, headers }, res => {
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => {
        const raw = Buffer.concat(chunks);
        let parsed; try { parsed = JSON.parse(raw.toString()); } catch { parsed = raw.toString(); }
        resolve({ status: res.statusCode, body: parsed, raw });
      });
    });
    r.on("error", reject);
    if (body) r.write(body);
    r.end();
  });
}

function putBytes(url, buf) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const r = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: "PUT",
      headers: { "Content-Length": buf.length } }, res => {
      res.on("data", () => {}); res.on("end", () => resolve(res.statusCode));
    });
    r.on("error", reject); r.write(buf); r.end();
  });
}

function getUrl(url) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    https.get({ hostname: u.hostname, path: u.pathname + u.search }, res => {
      const chunks = []; res.on("data", c => chunks.push(c));
      res.on("end", () => resolve({ status: res.statusCode, raw: Buffer.concat(chunks), text: Buffer.concat(chunks).toString() }));
    }).on("error", reject);
  });
}

async function uploadOss(t, buf, key) {
  const H = { Authorization: `Bearer ${t}`, "Content-Type": "application/json" };
  const up = await req("GET", BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${key}/signeds3upload`, H);
  if (up.status !== 200) throw new Error(`signeds3upload GET HTTP ${up.status}: ${JSON.stringify(up.body)}`);
  const s3 = await putBytes(up.body.urls[0], buf);
  if (s3 < 200 || s3 >= 300) throw new Error(`S3 PUT failed HTTP ${s3}`);
  const fin = await req("POST", BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${key}/signeds3upload`, H,
    JSON.stringify({ uploadKey: up.body.uploadKey }));
  if (fin.status < 200 || fin.status >= 300) throw new Error(`upload finalize HTTP ${fin.status}: ${JSON.stringify(fin.body)}`);
  const dl = await req("GET", BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${key}/signeds3download`, H);
  if (dl.status !== 200) throw new Error(`signeds3download HTTP ${dl.status}: ${JSON.stringify(dl.body)}`);
  console.log(`   uploaded ${key} (${buf.length} bytes)`);
  return dl.body.url;
}

async function signedWriteUrl(t, key) {
  const H = { Authorization: `Bearer ${t}`, "Content-Type": "application/json" };
  const r = await req("POST", BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${key}/signed?access=readwrite`, H, "{}");
  if (r.status < 200 || r.status >= 300) throw new Error(`signed write url HTTP ${r.status}: ${JSON.stringify(r.body)}`);
  return r.body.signedUrl;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// Submit a workitem and poll to a terminal state. Returns { status, reportUrl }.
async function runWorkitem(t, args, label) {
  const H = { Authorization: `Bearer ${t}`, "Content-Type": "application/json" };
  const wi = await req("POST", BASE, "/da/us-east/v3/workitems", H,
    JSON.stringify({ activityId: `${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`, arguments: args }));
  if (wi.status < 200 || wi.status >= 300) throw new Error(`workitem POST HTTP ${wi.status}: ${JSON.stringify(wi.body)}`);
  const id = wi.body.id;
  console.log(`   [${label}] workitem: ${id}`);
  let status = wi.body.status, reportUrl = wi.body.reportUrl;
  for (let i = 0; i < 80 && ["pending", "inprogress"].includes(status); i++) {
    await sleep(3000);
    const p = await req("GET", BASE, `/da/us-east/v3/workitems/${id}`, { Authorization: `Bearer ${t}` });
    status = p.body.status; reportUrl = p.body.reportUrl || reportUrl;
    process.stdout.write(`   [${label}] status: ${status}   \r`);
  }
  console.log(`\n   [${label}] final status: ${status}`);
  if (reportUrl) {
    const rep = await getUrl(reportUrl);
    console.log(`───── [${label}] work item report ─────`);
    console.log(rep.text.slice(0, 5000));
    console.log("──────────────────────────────────────");
  }
  return { status };
}

// Find the first non-constant attribute across the extracted title blocks.
function pickEditableAttr(titleBlocks) {
  for (const tb of titleBlocks || []) {
    for (const a of tb.attributes || []) {
      const constant = a.constant ?? a.Constant;
      const tag = a.tag ?? a.Tag;
      if (!constant && tag) {
        return { layout: tb.layout ?? tb.Layout, block: tb.block ?? tb.Block, tag,
                 currentValue: a.currentValue ?? a.CurrentValue ?? "" };
      }
    }
  }
  return null;
}

(async () => {
  console.log(`═══ Smoke test: AutoCADTitleBlockUpdater — ${ACTIVITY_ID}+${ALIAS} ═══`);
  console.log(`   Candidate DWGs: ${CANDIDATES.join(", ")}`);

  const auth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString("base64");
  const tokenRes = await req("POST", BASE, "/authentication/v2/token",
    { "Content-Type": "application/x-www-form-urlencoded", Authorization: `Basic ${auth}` },
    "grant_type=client_credentials&scope=code:all data:read data:write data:create bucket:create bucket:read");
  if (!tokenRes.body.access_token) throw new Error(`token failed: ${JSON.stringify(tokenRes.body)}`);
  const t = tokenRes.body.access_token;
  const H = { Authorization: `Bearer ${t}`, "Content-Type": "application/json" };

  const b = await req("POST", BASE, "/oss/v2/buckets", H, JSON.stringify({ bucketKey: BUCKET_KEY, policyKey: "transient" }));
  if (![200, 409].includes(b.status)) throw new Error(`bucket failed HTTP ${b.status}: ${JSON.stringify(b.body)}`);
  console.log(`   bucket: ${BUCKET_KEY} (HTTP ${b.status})`);

  const ts = Date.now();

  // ── PHASE 1 — EXTRACT probe across candidates ──────────────────────────────
  let winner = null;   // { dwgUrl, attr, dwgName }
  const probeLog = [];
  for (let i = 0; i < CANDIDATES.length && !winner; i++) {
    const dwgPath = CANDIDATES[i];
    if (!fs.existsSync(dwgPath)) { probeLog.push(`${dwgPath}: NOT FOUND`); continue; }
    const dwgName = path.basename(dwgPath);
    console.log(`\n── EXTRACT probe: ${dwgName} ──`);
    const dwgUrl = await uploadOss(t, fs.readFileSync(dwgPath), `tb-${ts}-${i}.dwg`);
    const resJsonUrl = await signedWriteUrl(t, `tb-extract-${ts}-${i}.json`);

    const { status } = await runWorkitem(t, {
      inputFile:  { url: dwgUrl,     verb: "get" },
      resultJson: { url: resJsonUrl, verb: "put" },
    }, `extract ${dwgName}`);
    if (status !== "success") { probeLog.push(`${dwgName}: workitem status=${status}`); continue; }

    const out = await getUrl(resJsonUrl);
    if (out.raw.length >= 3 && out.raw[0] === 0xEF && out.raw[1] === 0xBB && out.raw[2] === 0xBF)
      throw new Error("Smoke test FAILED: result.json has a UTF-8 BOM (write with new UTF8Encoding(false))");
    console.log(`───── ${dwgName} result.json ─────`);
    console.log(out.text.slice(0, 2000));
    console.log("──────────────────────────────────");
    let parsed; try { parsed = JSON.parse(out.text); } catch { throw new Error("Smoke test FAILED: extract result.json is not valid JSON"); }
    const ok   = parsed.ok   ?? parsed.Ok;
    const mode = parsed.mode ?? parsed.Mode;
    const tbs  = parsed.titleBlocks ?? parsed.TitleBlocks ?? [];
    if (ok !== true || mode !== "extract")
      throw new Error(`Smoke test FAILED: extract expected ok:true mode:"extract" — got ${JSON.stringify(parsed).slice(0, 200)}`);
    const tbCount = Array.isArray(tbs) ? tbs.length : 0;
    probeLog.push(`${dwgName}: ${tbCount} title block(s)`);
    if (tbCount > 0) {
      const attr = pickEditableAttr(tbs);
      if (attr) { winner = { dwgUrl, attr, dwgName }; }
      else probeLog.push(`${dwgName}: title blocks present but all attributes are constant`);
    }
  }

  if (!winner) {
    throw new Error(
      "Smoke test FAILED: no candidate DWG produced an editable title-block schema.\n" +
      "Probe results:\n  - " + probeLog.join("\n  - ") + "\n" +
      "→ Commit a DWG that has an attributed title block on a paper-space layout and\n" +
      "  set SMOKE_SAMPLE_DWGS to point at it (see scripts/smoke-test-titleblock.js).");
  }
  console.log(`\n✅ EXTRACT gate PASSED — fixture=${winner.dwgName}, ` +
    `editable attr: layout='${winner.attr.layout}' block='${winner.attr.block}' tag='${winner.attr.tag}'`);

  // ── PHASE 2 — UPDATE the picked attribute inline via params.json ────────────
  console.log(`\n── UPDATE: set ${winner.attr.tag} on ${winner.dwgName} ──`);
  const newValue = "SMOKE-" + String(ts).slice(-6);
  const params = {
    titleBlockName: "*",
    layoutScope: "all",
    changes: { [winner.attr.tag]: newValue },
  };
  const paramsUrl    = await uploadOss(t, Buffer.from(JSON.stringify(params), "utf8"), `tb-params-${ts}.json`);
  const resultDwgUrl = await signedWriteUrl(t, `tb-result-${ts}.dwg`);
  const resultJsonUrl = await signedWriteUrl(t, `tb-update-${ts}.json`);

  const { status } = await runWorkitem(t, {
    inputFile:  { url: winner.dwgUrl,   verb: "get" },
    params:     { url: paramsUrl,       verb: "get" },
    resultDwg:  { url: resultDwgUrl,    verb: "put" },
    resultJson: { url: resultJsonUrl,   verb: "put" },
  }, "update");
  if (status !== "success") throw new Error(`Smoke test FAILED: update workitem status='${status}'`);

  const out = await getUrl(resultJsonUrl);
  console.log("───── update result.json ─────");
  console.log(out.text.slice(0, 2000));
  console.log("──────────────────────────────");
  if (out.raw.length >= 3 && out.raw[0] === 0xEF && out.raw[1] === 0xBB && out.raw[2] === 0xBF)
    throw new Error("Smoke test FAILED: update result.json has a UTF-8 BOM");
  let parsed; try { parsed = JSON.parse(out.text); } catch { throw new Error("Smoke test FAILED: update result.json is not valid JSON"); }
  const ok    = parsed.ok    ?? parsed.Ok;
  const mode  = parsed.mode  ?? parsed.Mode;
  const count = parsed.updatedCount ?? parsed.UpdatedCount;
  if (ok !== true) throw new Error(`Smoke test FAILED: update ok=false: ${parsed.error ?? parsed.Error}`);
  if (mode !== "update") throw new Error(`Smoke test FAILED: expected mode:"update", got '${mode}'`);
  if (typeof count !== "number" || count < 1)
    throw new Error(`Smoke test FAILED: expected updatedCount>0, got ${count}`);

  // result.dwg must have been written.
  const dwgOut = await getUrl(resultDwgUrl);
  if (dwgOut.raw.length < 100)
    throw new Error(`Smoke test FAILED: result.dwg is suspiciously small (${dwgOut.raw.length} bytes)`);

  console.log(`\n✅ Smoke test PASSED — extract found title blocks (${winner.dwgName}); ` +
    `update applied ${count} change(s); result.dwg=${dwgOut.raw.length} bytes.`);
})().catch(err => { console.error("\n❌", err.message); process.exit(1); });
