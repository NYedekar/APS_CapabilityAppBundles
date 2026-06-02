#!/usr/bin/env node
// Smoke gate for RevitParameterUpdater: upload sample.rvt + an empty delta CSV + params.json,
// run the just-published activity, and FAIL unless:
//   • workitem status == "success"
//   • result.json is valid JSON with no BOM and has ok:true + numeric summary fields
//   • result.rvt was uploaded (size > 0 bytes)
//
// An empty delta CSV (header-only) is intentional — it proves the bundle loads and runs
// without requiring us to know the specific element names in sample.rvt.
// Adapted from scripts/smoke-test-revit.js.
"use strict";

const fs    = require("fs");
const path  = require("path");
const https = require("https");

const CLIENT_ID     = required("APS_CLIENT_ID");
const CLIENT_SECRET = required("APS_CLIENT_SECRET");
const NICKNAME      = required("APS_NICKNAME");
const ACTIVITY_ID   = process.env.SMOKE_ACTIVITY_ID || "UpdateRevitParameters";
const ALIAS         = process.env.ALIAS             || "prod";
const SAMPLE_RVT    = process.env.SAMPLE_RVT        || path.join(process.cwd(), "test", "sample.rvt");
const BUCKET_KEY    = (NICKNAME.toLowerCase().replace(/[^a-z0-9]/g, "") + "-smoke").slice(0, 60);
const BASE          = "developer.api.autodesk.com";

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
  console.log(`   uploaded ${key} (${buf.length} bytes)`);
  // Get signed download URL
  const dl = await req("GET", BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${key}/signeds3download`, H);
  if (dl.status !== 200) throw new Error(`signeds3download HTTP ${dl.status}: ${JSON.stringify(dl.body)}`);
  return dl.body.url;
}

async function signedWriteUrl(t, key) {
  const H = { Authorization: `Bearer ${t}`, "Content-Type": "application/json" };
  const r = await req("POST", BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${key}/signed?access=readwrite`, H, "{}");
  if (r.status < 200 || r.status >= 300) throw new Error(`signed write url HTTP ${r.status}: ${JSON.stringify(r.body)}`);
  return r.body.signedUrl;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  console.log(`═══ Smoke test: RevitParameterUpdater — ${ACTIVITY_ID}+${ALIAS} ═══`);
  if (!fs.existsSync(SAMPLE_RVT)) throw new Error(`Sample RVT not found: ${SAMPLE_RVT}`);

  // Inline test fixtures — no extra test files needed
  // params_input.dat: empty CSV (header row only) → 0 change requests → proves bundle loads
  const paramsInputBuf = Buffer.from("ElementId,ElementName,Category,Parameter,NewValue\n", "utf8");
  // params.json: delta_file mode (no delta needed for empty input)
  const paramsJsonBuf  = Buffer.from(JSON.stringify({ inputMode: "delta_file" }), "utf8");
  const rvtBuf         = fs.readFileSync(SAMPLE_RVT);

  const auth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString("base64");
  const tokenRes = await req("POST", BASE, "/authentication/v2/token",
    { "Content-Type": "application/x-www-form-urlencoded", Authorization: `Basic ${auth}` },
    "grant_type=client_credentials&scope=code:all data:read data:write data:create bucket:create bucket:read");
  if (!tokenRes.body.access_token) throw new Error(`token failed: ${JSON.stringify(tokenRes.body)}`);
  const t = tokenRes.body.access_token;
  const H = { Authorization: `Bearer ${t}`, "Content-Type": "application/json" };

  // 1. Ensure bucket
  const b = await req("POST", BASE, "/oss/v2/buckets", H, JSON.stringify({ bucketKey: BUCKET_KEY, policyKey: "transient" }));
  if (![200, 409].includes(b.status)) throw new Error(`bucket failed HTTP ${b.status}: ${JSON.stringify(b.body)}`);
  console.log(`   bucket: ${BUCKET_KEY} (HTTP ${b.status})`);

  const ts = Date.now();
  // 2. Upload all inputs
  const rvtUrl         = await uploadOss(t, rvtBuf,         `smoke-${ts}.rvt`);
  const paramsInputUrl = await uploadOss(t, paramsInputBuf, `smoke-params-input-${ts}.dat`);
  const paramsJsonUrl  = await uploadOss(t, paramsJsonBuf,  `smoke-params-${ts}.json`);

  // 3. Signed write URLs for outputs
  const resultRvtUrl  = await signedWriteUrl(t, `smoke-result-${ts}.rvt`);
  const resultJsonUrl = await signedWriteUrl(t, `smoke-result-${ts}.json`);

  // 4. Submit workitem
  const wiBody = JSON.stringify({
    activityId: `${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`,
    arguments: {
      rvtFile:     { url: rvtUrl,         verb: "get" },
      paramsInput: { url: paramsInputUrl, verb: "get" },
      params:      { url: paramsJsonUrl,  verb: "get" },
      resultRvt:   { url: resultRvtUrl,   verb: "put" },
      resultJson:  { url: resultJsonUrl,  verb: "put" },
    },
  });
  const wi = await req("POST", BASE, "/da/us-east/v3/workitems", H, wiBody);
  if (wi.status < 200 || wi.status >= 300) throw new Error(`workitem POST HTTP ${wi.status}: ${JSON.stringify(wi.body)}`);
  const id = wi.body.id;
  console.log(`   workitem: ${id} → ${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`);

  // 5. Poll (Revit cold-start can take 3+ minutes — allow 6 min)
  let status = wi.body.status, reportUrl = wi.body.reportUrl;
  for (let i = 0; i < 120 && ["pending", "inprogress"].includes(status); i++) {
    await sleep(3000);
    const p = await req("GET", BASE, `/da/us-east/v3/workitems/${id}`, { Authorization: `Bearer ${t}` });
    status = p.body.status; reportUrl = p.body.reportUrl || reportUrl;
    process.stdout.write(`   status: ${status}\r`);
  }
  console.log(`\n   final status: ${status}`);

  // 6. Always dump report
  if (reportUrl) {
    const rep = await getUrl(reportUrl);
    console.log("───── work item report ─────");
    console.log(rep.text.slice(0, 5000));
    console.log("────────────────────────────");
  }
  if (status !== "success")
    throw new Error(`Smoke FAILED: workitem status='${status}' (expected 'success')`);

  // 7. Verify result.json
  const out = await getUrl(resultJsonUrl);
  console.log("───── result.json ─────");
  console.log(out.text.slice(0, 2000));
  console.log("───────────────────────");
  if (out.raw.length >= 3 && out.raw[0] === 0xEF && out.raw[1] === 0xBB && out.raw[2] === 0xBF)
    throw new Error("Smoke FAILED: result.json has UTF-8 BOM (write with new UTF8Encoding(false))");
  let parsed;
  try { parsed = JSON.parse(out.text); } catch { throw new Error("Smoke FAILED: result.json is not valid JSON"); }
  if (typeof parsed.ok !== "boolean")
    throw new Error(`Smoke FAILED: result.json missing boolean 'ok' (got: ${JSON.stringify(parsed).slice(0, 200)})`);
  if (!parsed.ok)
    throw new Error(`Smoke FAILED: result.json has ok=false: ${parsed.error}`);
  if (typeof parsed.summary?.total !== "number")
    throw new Error(`Smoke FAILED: result.json missing numeric summary.total`);

  // 8. Verify result.rvt was written (> 0 bytes)
  const rvtOut = await getUrl(resultRvtUrl);
  if (rvtOut.raw.length < 100)
    throw new Error(`Smoke FAILED: result.rvt is suspiciously small (${rvtOut.raw.length} bytes)`);

  console.log(`✅ Smoke PASSED — status:success, ok=true, summary.total=${parsed.summary.total}, result.rvt=${rvtOut.raw.length} bytes`);
})().catch(err => { console.error("\n❌", err.message); process.exit(1); });
