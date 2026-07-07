#!/usr/bin/env node
// One-off cleanup: delete a SPECIFIC, explicit list of throwaway Design Automation
// appbundles + activities created during the appbuilder live test-drive. Exact-id
// deletion only (no pattern sweep): it deletes exactly the ids in DELETE_IDS and
// nothing else. Run via cleanup-demo-bundles.yml workflow_dispatch (needs APS secrets).
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
// Exact ids to remove (comma-separated). These are the bundles this session created.
const DELETE_IDS = (process.env.DELETE_IDS || 'AUDemoAcadSmoke,AUDemoRevitCheck,AUDemoInvCheck')
  .split(',').map(s => s.trim()).filter(Boolean);
const BASE = 'developer.api.autodesk.com';

function required(name) {
  const v = process.env[name];
  if (!v) { console.error(`Missing env var: ${name}`); process.exit(1); }
  return v;
}

function req(method, path, headers, body) {
  return new Promise((resolve, reject) => {
    const r = https.request({ hostname: BASE, path, method, headers }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => { try { resolve({ status: res.statusCode, body: JSON.parse(data) }); } catch { resolve({ status: res.statusCode, body: data }); } });
    });
    r.on('error', reject);
    if (body) r.write(body);
    r.end();
  });
}

async function token() {
  const auth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  const res = await req('POST', '/authentication/v2/token',
    { 'Content-Type': 'application/x-www-form-urlencoded', 'Authorization': `Basic ${auth}` },
    'grant_type=client_credentials&scope=code%3Aall');
  if (!res.body.access_token) throw new Error(`token failed: ${JSON.stringify(res.body)}`);
  return res.body.access_token;
}

async function main() {
  const t = await token();
  const H = { 'Authorization': `Bearer ${t}` };
  console.log(`Deleting exact ids: ${DELETE_IDS.join(', ')}`);
  for (const id of DELETE_IDS) {
    for (const kind of ['appbundles', 'activities']) {
      const del = await req('DELETE', `/da/us-east/v3/${kind}/${id}`, H);
      // 204 = deleted, 404 = did not exist (fine, e.g. a bundle that never published)
      const note = del.status === 404 ? '(not present)' : '';
      console.log(`   ${kind}/${id} -> HTTP ${del.status} ${note}`);
    }
  }
  console.log('Cleanup done.');
}

main().catch(err => { console.error('Cleanup failed:', err.message); process.exit(1); });
