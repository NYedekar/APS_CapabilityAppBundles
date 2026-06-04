#!/usr/bin/env node
// Post-publish smoke gate: uploads the committed sample Inventor zip, runs the
// ExtractInventorBOM+prod activity end-to-end, and FAILS the build unless
// the work item reports status:"success" with valid result.json.
const fs    = require('fs');
const path  = require('path');
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const NICKNAME      = required('APS_NICKNAME');
const ACTIVITY_ID   = process.env.SMOKE_ACTIVITY_ID || 'ExtractInventorBOM';
const ALIAS         = process.env.ALIAS || 'prod';
// sample-inventor.zip = official Autodesk 2022 "Blower" sample (Blower.iam + 31 sub-assemblies
// + 38 parts), a real self-contained multi-level assembly for genuine BOM extraction.
const SAMPLE_ZIP    = process.env.SAMPLE_INVENTOR_ZIP || path.join(process.cwd(), 'test', 'inventor', 'sample-inventor.zip');
const BUCKET_KEY    = (NICKNAME.toLowerCase().replace(/[^a-z0-9]/g, '') + '-smoke').slice(0, 60);
const BASE          = 'developer.api.autodesk.com';

function required(name) {
  const v = process.env[name];
  if (!v) { console.error(`❌ Missing env var: ${name}`); process.exit(1); }
  return v;
}

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
      res.on('end', () => resolve({ status: res.statusCode, text: Buffer.concat(chunks).toString() }));
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
  console.log('═══ Smoke test: run ExtractInventorBOM end-to-end ═══');
  if (!fs.existsSync(SAMPLE_ZIP)) throw new Error(`Sample Inventor zip not found: ${SAMPLE_ZIP}`);
  const zipBuf = fs.readFileSync(SAMPLE_ZIP);
  const t = await token();
  const H = { 'Authorization': `Bearer ${t}`, 'Content-Type': 'application/json' };

  // 1. Ensure transient bucket
  const b = await req('POST', BASE, '/oss/v2/buckets', H, JSON.stringify({ bucketKey: BUCKET_KEY, policyKey: 'transient' }));
  if (![200, 409].includes(b.status)) throw new Error(`bucket failed HTTP ${b.status}: ${JSON.stringify(b.body)}`);
  console.log(`   bucket: ${BUCKET_KEY} (HTTP ${b.status})`);

  // 2. Signed S3 upload — key ends in .zip so DA treats it as a zipped assembly
  const objKey = `smoke-inventor-${Date.now()}.zip`;
  const up = await req('GET', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${objKey}/signeds3upload`, H);
  if (up.status !== 200) throw new Error(`signeds3upload GET HTTP ${up.status}: ${JSON.stringify(up.body)}`);
  const putStatus = await putBytes(up.body.urls[0], zipBuf);
  if (putStatus < 200 || putStatus >= 300) throw new Error(`S3 PUT failed HTTP ${putStatus}`);
  const fin = await req('POST', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${objKey}/signeds3upload`, H,
    JSON.stringify({ uploadKey: up.body.uploadKey }));
  if (fin.status < 200 || fin.status >= 300) throw new Error(`upload finalize HTTP ${fin.status}: ${JSON.stringify(fin.body)}`);
  console.log(`   uploaded: ${objKey} (${zipBuf.length} bytes)`);

  // 3. Signed download URL
  const dl = await req('GET', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${objKey}/signeds3download`, H);
  if (dl.status !== 200) throw new Error(`signeds3download HTTP ${dl.status}: ${JSON.stringify(dl.body)}`);
  const inputUrl = dl.body.url;

  // 4. Readwrite signed URL for result.json
  const resObj = `smoke-inventor-result-${Date.now()}.json`;
  const sig = await req('POST', BASE, `/oss/v2/buckets/${BUCKET_KEY}/objects/${resObj}/signed?access=readwrite`, H, '{}');
  if (sig.status < 200 || sig.status >= 300) throw new Error(`signed result url HTTP ${sig.status}: ${JSON.stringify(sig.body)}`);
  const resultUrl = sig.body.signedUrl;

  // 5. Submit work item
  const wiBody = JSON.stringify({
    activityId: `${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`,
    arguments: {
      inputFile:  { url: inputUrl,  verb: 'get', pathInZip: 'Blower.iam' },
      resultJson: { url: resultUrl, verb: 'put' },
    },
  });
  const wi = await req('POST', BASE, '/da/us-east/v3/workitems', H, wiBody);
  if (wi.status < 200 || wi.status >= 300) throw new Error(`workitem POST HTTP ${wi.status}: ${JSON.stringify(wi.body)}`);
  const id = wi.body.id;
  console.log(`   workitem: ${id}  →  ${NICKNAME}.${ACTIVITY_ID}+${ALIAS}`);

  // 6. Poll (Inventor workitems can take longer than AutoCAD — allow 90 polls × 5s = 7.5 min)
  let status = wi.body.status, reportUrl = wi.body.reportUrl;
  for (let i = 0; i < 90 && ['pending', 'inprogress'].includes(status); i++) {
    await sleep(5000);
    const p = await req('GET', BASE, `/da/us-east/v3/workitems/${id}`, { 'Authorization': `Bearer ${t}` });
    status = p.body.status; reportUrl = p.body.reportUrl || reportUrl;
    process.stdout.write(`   status: ${status}\r`);
  }
  console.log(`\n   final status: ${status}`);

  // 7. Always dump the report log
  if (reportUrl) {
    const rep = await getUrl(reportUrl);
    console.log('───── work item report ─────');
    console.log(rep.text);
    console.log('────────────────────────────');
  }

  if (status !== 'success') {
    throw new Error(`Smoke test FAILED: work item status='${status}'. The appbundle is not running correctly.`);
  }

  // 8. Verify result.json
  const out = await getUrl(resultUrl);
  console.log('───── result.json (first 2000 chars) ─────');
  console.log(out.text.slice(0, 2000));
  console.log('──────────────────────────────────────────');
  let parsed;
  try { parsed = JSON.parse(out.text.replace(/^﻿/, '')); } catch { throw new Error('result.json is not valid JSON.'); }
  if (!parsed.topLevelRows) throw new Error('result.json missing expected field "topLevelRows".');

  console.log(`✅ Smoke test PASSED — ${parsed.totalComponents} components, ${parsed.topLevelRows.length} top-level rows.`);
})().catch(err => { console.error('\n❌', err.message); process.exit(1); });
