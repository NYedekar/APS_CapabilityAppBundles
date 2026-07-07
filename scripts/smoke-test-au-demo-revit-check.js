#!/usr/bin/env node
// Post-publish Revit smoke gate: upload the committed sample RVT, run the just-published
// Revit activity end-to-end, and FAIL the build unless the work item reports status:"success"
// AND produces a valid, NO-BOM result.json. Catches a silent regression (e.g. the appbundle
// no longer loading, or DesignAutomationBridge.dll missing) before it reaches a user.
//
// Adapted from the real, shipped smoke-test-revit.js (APS_CapabilityAppBundles), generalized so
// it works for ANY Revit activity's result.json shape, not just SheetList. Activity I/O contract
// (see templates/ci/revit publish scripts): input arg `rvtFile` (get), output arg `resultJson` (put).
const fs    = require('fs');
const path  = require('path');
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const NICKNAME      = required('APS_NICKNAME');
const ACTIVITY_ID   = required('SMOKE_ACTIVITY_ID');
const ALIAS         = process.env.ALIAS || 'prod';
const SAMPLE_RVT    = process.env.SAMPLE_RVT || path.join(process.cwd(), 'test', 'sample.rvt');
const EXPECT_KEY    = process.env.EXPECT_KEY || null;
const BUCKET_KEY    = (NICKNAME.toLowerCase().replace(/[^a-z0-9]/g, '') + '-smoke').slice(0, 60);
const BASE          = 'developer.api.autodesk.com';

function required(name) {
  const v = process.env[name];
  if (!v) { console.error(`❌ Missing env var: ${name}`); process.exit(1); }
  return v;
}

// Minimal HTTPS helper. Returns { status, body, raw }.
function req(method, host, p, headers, body) {
  return new Promise((resolve, reject) => {
    const r = https.request({ hostname: host, path: p, method, headers }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const raw = Buffer.concat(chunks);
        let parsed; try { parsed = JSON.parse(raw.toString()); } catch { parsed = raw.toString(); }
        resolve({ status: res.statusCode, body: parsed, raw });
      });
    });
    r.on('error', reject);
    if (body) r.write(body);
    r.end();
  });
}

// PUT raw bytes to an arbitrary URL (S3 signed upload). Returns status.
function putBytes(url, buf) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const r = https.request({ hostname: u.hostname, path: u.pathname + u.search, method: 'PUT',
      headers: { 'Content-Length': buf.length } }, res => {
      res.on('data', () => {}); res.on('end', () => resolve(res.statusCode));
    });
    r.on('error', reject); r.write(buf); r.end();
  });
}

function getUrl(url) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    https.get({ hostname: u.hostname, path: u.pathname + u.search }, res => {
      const chunks = []; res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, raw: Buffer.concat(chunks), text: Buffer.concat(chunks).toString() }));
    }).on('error', reject);
  });
}

async function token() {
  const auth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  const res = await req('POST', BASE, '/authentication/v2/token',
    { 'Content-Type': 'application/x-www-form-urlencoded', 'Authorization': `Basic ${auth}` },
    'grant_type=client_credentials&scope=code:all data:read data:write data:create bucket:create bucket:read');
  if (!res.body.access_token) throw new Error(`token failed: ${JSON.stringify(res.body)}`);
  return res.body.access_token;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  console.log(`═══ Revit smoke test: run ${ACTIVITY_ID}+${ALIAS} end-to-end ═══`);
  if (!fs.existsSync(SAMPLE_RVT)) throw new Error(`Sample RVT not found: ${SAMPLE_RVT}`);
  const rvt = fs.readFileSync(SAMPLE_RVT);
  const t = await token();
  const H = { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' };

  // 1. Ensure transient bucket (OSS create bucket is POST, not PUT)
  const b = await req('POST', BASE, '/oss/v2/buckets', H, JSON.stringify({ bucketKey: BUCKET_KEY, policyKey: 'transient' }));
  if (![200, 409].includes(b.status)) throw new Error(`bucket failed HTTP ${b.status}: ${JSON.stringify(b.body)}`);
  console.log(`   bucket: ${BUCKET_KEY} (HTTP ${b.status})`);

  // 2. Signed S3 upload of input RVT
  const objKey = `smoke-${Date.now()}.rvt`;
  const up = await req('GET', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${objKey}/signeds3upload`, H);
  if (up.status !== 200) throw new Error(`signeds3upload GET HTTP ${up.status}: ${JSON.stringify(up.body)}`);
  const putStatus = await putBytes(up.body.urls[0], rvt);
  if (putStatus < 200 || putStatus >= 300) throw new Error(`S3 PUT failed HTTP ${putStatus}`);
  const fin = await req('POST', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${objKey}/signeds3upload`, H,
    JSON.stringify({ uploadKey: up.body.uploadKey }));
  if (fin.status < 200 || fin.status >= 300) throw new Error(`upload finalize HTTP ${fin.status}: ${JSON.stringify(fin.body)}`);
  console.log(`   uploaded input: ${objKey} (${rvt.length} bytes)`);

  // 3. Signed download URL for input
  const dl = await req('GET', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${objKey}/signeds3download`, H);
  if (dl.status !== 200) throw new Error(`signeds3download HTTP ${dl.status}: ${JSON.stringify(dl.body)}`);
  const inputUrl = dl.body.url;

  // 4. Readwrite signed URL for result.json (workitem PUTs here, we GET it after)
  const resObj = `smoke-result-${Date.now()}.json`;
  const sig = await req('POST', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${resObj}/signed?access=readwrite`, H, '{}');
  if (sig.status < 200 || sig.status >= 300) throw new Error(`signed result url HTTP ${sig.status}: ${JSON.stringify(sig.body)}`);
  const resultUrl = sig.body.signedUrl;

  // 5. Submit work item (arg names must match the activity: rvtFile in, resultJson out)
  const wiBody = JSON.stringify({
    activityId: `${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`,
    arguments: {
      rvtFile:    { url: inputUrl, verb: 'get' },
      resultJson: { url: resultUrl, verb: 'put' },
    },
  });
  const wi = await req('POST', BASE, '/da/us-east/v3/workitems', H, wiBody);
  if (wi.status < 200 || wi.status >= 300) throw new Error(`workitem POST HTTP ${wi.status}: ${JSON.stringify(wi.body)}`);
  const id = wi.body.id;
  console.log(`   workitem: ${id} → activity ${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`);

  // 6. Poll (Revit cold-start can be slow → allow ~6 min)
  let status = wi.body.status, reportUrl = wi.body.reportUrl;
  for (let i = 0; i < 120 && ['pending', 'inprogress'].includes(status); i++) {
    await sleep(3000);
    const p = await req('GET', BASE, `/da/us-east/v3/workitems/${id}`, { 'Authorization': `Bearer ${t}` });
    status = p.body.status; reportUrl = p.body.reportUrl || reportUrl;
    process.stdout.write(`   status: ${status}\r`);
  }
  console.log(`\n   final status: ${status}`);

  // 7. Always dump the report log between clear markers
  if (reportUrl) {
    const rep = await getUrl(reportUrl);
    console.log('───── work item report ─────');
    console.log(rep.text);
    console.log('────────────────────────────');
  }

  if (status !== 'success') {
    console.error(`\n❌ Revit smoke test FAILED: work item status='${status}' (expected 'success'). The appbundle/activity is not running correctly.`);
    process.exit(1);
  }

  // 8. Verify result.json: present, valid JSON, NO BOM, and generically well-formed
  const out = await getUrl(resultUrl);
  console.log('───── result.json ─────');
  console.log(out.text.slice(0, 2000));
  console.log('───────────────────────');

  // Strict no-BOM gate (a BOM is a real bug — strict JSON parsers choke on it; write result.json
  // with new UTF8Encoding(false) on the .NET side). No auto-strip: reject outright.
  if (out.raw.length >= 3 && out.raw[0] === 0xEF && out.raw[1] === 0xBB && out.raw[2] === 0xBF) {
    console.error('\n❌ Revit smoke test FAILED: result.json starts with a UTF-8 BOM. Write it with new UTF8Encoding(false).');
    console.error(out.text.slice(0, 2000));
    process.exit(1);
  }

  let parsed;
  try { parsed = JSON.parse(out.text); } catch {
    console.error('\n❌ Revit smoke test FAILED: result.json is not valid JSON.');
    console.error(out.text.slice(0, 2000));
    process.exit(1);
  }

  // Generic success check: this activity's result.json shape is not known ahead of time (that's
  // specific to whatever capability the bundle implements), so we only assert the parsed JSON
  // exists and does not carry an explicit failure marker (ok: false). Set EXPECT_KEY to also
  // require a specific field to be present (e.g. EXPECT_KEY=SheetCount, EXPECT_KEY=Count).
  if (!parsed || typeof parsed !== 'object') {
    console.error(`\n❌ Revit smoke test FAILED: result.json did not parse to an object (got: ${JSON.stringify(parsed).slice(0, 200)}).`);
    process.exit(1);
  }
  if (parsed.ok === false) {
    console.error(`\n❌ Revit smoke test FAILED: result.json reports ok:false (got: ${JSON.stringify(parsed).slice(0, 200)}).`);
    process.exit(1);
  }
  if (EXPECT_KEY && !(EXPECT_KEY in parsed)) {
    console.error(`\n❌ Revit smoke test FAILED: result.json is missing expected key '${EXPECT_KEY}' (got: ${JSON.stringify(parsed).slice(0, 200)}).`);
    process.exit(1);
  }

  console.log(`✅ Revit smoke test PASSED — status:success, no BOM, valid JSON${EXPECT_KEY ? `, has '${EXPECT_KEY}'` : ''}.`);
})().catch(err => { console.error('\n❌ Revit smoke test FAILED:', err.message); process.exit(1); });
