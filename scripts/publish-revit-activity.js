#!/usr/bin/env node
// Generic single-activity publisher for Revit DA bundles, written by scaffold-bundle.js.
// Unlike templates/ci/revit/publish-pdf-activity.js (hardcoded to the RevitPDFExport
// dual-activity / params.json "operation" contract), this publishes ONE activity per
// invocation, driven entirely by env vars, mirroring publish-autocad-activity.js.
'use strict';
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const NICKNAME      = required('APS_NICKNAME');
const BUNDLE_NAME   = required('BUNDLE_NAME');
const ENGINE_VER    = required('ENGINE_VERSION');
const ALIAS         = process.env.ALIAS || 'prod';
const ACTIVITY_ID   = required('ACTIVITY_ID');
const ENGINE_ID     = `Autodesk.Revit+${ENGINE_VER}`;
const BUNDLE_REF    = `${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`;
const BASE          = 'developer.api.autodesk.com';

function required(name) {
  const v = process.env[name];
  if (!v) { console.error(`Missing env var: ${name}`); process.exit(1); }
  return v;
}

function request(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getToken() {
  const auth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  const res = await request({
    hostname: BASE, path: '/authentication/v2/token', method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Authorization': `Basic ${auth}` },
  }, 'grant_type=client_credentials&scope=code%3Aall');
  if (!res.body.access_token) throw new Error(`No access_token: ${JSON.stringify(res.body)}`);
  return res.body.access_token;
}

function daHeaders(token) { return { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }; }

function buildActivityDef() {
  return {
    engine: ENGINE_ID,
    appbundles: [BUNDLE_REF],
    commandLine: [`$(engine.path)\\revitcoreconsole.exe /i "$(args[rvtFile].path)" /al "$(appbundles[${BUNDLE_NAME}].path)"`],
    parameters: {
      rvtFile:    { verb: 'get', localName: 'input.rvt', required: true, description: 'Input Revit model' },
      resultJson: { verb: 'put', localName: 'result.json', required: false, description: 'result.json output' },
    },
    description: `${BUNDLE_NAME} activity`,
  };
}

async function deleteAndRecreate(token, def) {
  console.log('   100-version limit reached, deleting activity and recreating...');
  const del = await request({ hostname: BASE, path: `/da/us-east/v3/activities/${ACTIVITY_ID}`, method: 'DELETE', headers: daHeaders(token) });
  if (del.status !== 204 && del.status !== 200) throw new Error(`Delete failed HTTP ${del.status}: ${JSON.stringify(del.body)}`);
  return request({ hostname: BASE, path: '/da/us-east/v3/activities', method: 'POST', headers: daHeaders(token) }, JSON.stringify({ id: ACTIVITY_ID, ...def }));
}

async function publishActivity(token) {
  const def = buildActivityDef();
  let res = await request({ hostname: BASE, path: '/da/us-east/v3/activities', method: 'POST', headers: daHeaders(token) }, JSON.stringify({ id: ACTIVITY_ID, ...def }));
  if (res.status === 409) {
    res = await request({ hostname: BASE, path: `/da/us-east/v3/activities/${ACTIVITY_ID}/versions`, method: 'POST', headers: daHeaders(token) }, JSON.stringify(def));
    if (res.status === 403) res = await deleteAndRecreate(token, def);
  } else if (res.status === 403) {
    res = await deleteAndRecreate(token, def);
  }
  if (res.status < 200 || res.status >= 300) throw new Error(`HTTP ${res.status}: ${JSON.stringify(res.body)}`);
  console.log(`Activity version ${res.body.version} ready`);
  return res.body.version;
}

async function setAlias(token, version) {
  let res = await request({ hostname: BASE, path: `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases/${ALIAS}`, method: 'PATCH', headers: daHeaders(token) }, JSON.stringify({ version }));
  if (res.status === 404) {
    res = await request({ hostname: BASE, path: `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases`, method: 'POST', headers: daHeaders(token) }, JSON.stringify({ id: ALIAS, version }));
  }
  if (res.status < 200 || res.status >= 300) throw new Error(`Alias failed HTTP ${res.status}: ${JSON.stringify(res.body)}`);
  console.log(`Alias '${ALIAS}' -> v${version}`);
}

(async () => {
  console.log(`APS Activity Publisher - ${ACTIVITY_ID}`);
  const token = await getToken();
  const version = await publishActivity(token);
  await setAlias(token, version);
  console.log(`Done! ${ACTIVITY_ID}+${ALIAS} -> v${version}`);
})().catch(err => { console.error('Activity publish failed:', err.message); process.exit(1); });
