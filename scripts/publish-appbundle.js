#!/usr/bin/env node
/**
 * publish-appbundle.js
 *
 * Publishes the RevitExtractor AppBundle to APS Design Automation.
 * Steps:
 *   1. Get a 2-legged OAuth token
 *   2. Create or update the AppBundle definition (returns an upload URL)
 *   3. Upload the zip to the signed S3 upload URL
 *   4. Create / update an alias (e.g. "prod") pointing to the new version
 *
 * Usage:
 *   APS_CLIENT_ID=xxx APS_CLIENT_SECRET=yyy node scripts/publish-appbundle.js
 *
 * Optional env vars:
 *   BUNDLE_NAME      default: RevitExtractor
 *   ENGINE_VERSION   default: 2024   (used as Autodesk.Revit+<version>)
 *   ALIAS            default: prod
 *   BUNDLE_ZIP       default: RevitExtractor.zip  (relative to repo root)
 */

const fs   = require('fs');
const path = require('path');
const https = require('https');

// ─── Config ────────────────────────────────────────────────────────────────
const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const BUNDLE_NAME   = process.env.BUNDLE_NAME    || 'RevitExtractor';
const ENGINE_VER    = process.env.ENGINE_VERSION || '2024';
const ALIAS         = process.env.ALIAS          || 'prod';
const ZIP_PATH      = process.env.BUNDLE_ZIP     || path.join(process.cwd(), 'RevitExtractor.zip');

const ENGINE_ID     = `Autodesk.Revit+${ENGINE_VER}`;
const DA_BASE       = 'https://developer.api.autodesk.com/da/us-east/v3';

// ─── Helpers ───────────────────────────────────────────────────────────────
function required(name) {
  const val = process.env[name];
  if (!val) { console.error(`❌ Missing env var: ${name}`); process.exit(1); }
  return val;
}

function request(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        const status = res.statusCode;
        if (status >= 200 && status < 300) {
          try { resolve({ status, body: JSON.parse(data) }); }
          catch { resolve({ status, body: data }); }
        } else {
          reject(new Error(`HTTP ${status}: ${data}`));
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getToken() {
  console.log('🔑 Getting 2-legged token...');
  const body = `grant_type=client_credentials&scope=code%3Aall`;
  const auth  = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');

  const res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     '/authentication/v2/token',
    method:   'POST',
    headers: {
      'Content-Type':  'application/x-www-form-urlencoded',
      'Authorization': `Basic ${auth}`,
    },
  }, body);

  const token = res.body.access_token;
  if (!token) throw new Error('No access_token in response');
  console.log('   ✅ Token acquired');
  return token;
}

function daHeaders(token, extra = {}) {
  return { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', ...extra };
}

async function createOrUpdateAppBundle(token) {
  console.log(`\n📦 Creating/updating AppBundle: ${BUNDLE_NAME}`);

  const payload = JSON.stringify({
    id:       BUNDLE_NAME,
    engine:   ENGINE_ID,
    description: 'Extracts model data (rooms, walls, floors, elements) to result.json',
  });

  // POST always creates a new version; 409 means the bundle name exists — that's fine,
  // we re-POST to create a new version each time.
  let uploadUrl;
  try {
    const res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/appbundles`,
      method:   'POST',
      headers:  daHeaders(token),
    }, payload);

    console.log(`   ✅ AppBundle version ${res.body.version} created`);
    uploadUrl = res.body.uploadParameters?.endpointURL;

    // Store version for alias step
    process.env._BUNDLE_VERSION = String(res.body.version);
    return res.body;

  } catch (err) {
    // If we get 409 the bundle name exists — create a new version via POST again is correct.
    // The DA API always creates a NEW version on POST (it doesn't replace).
    throw err;
  }
}

async function uploadZip(uploadParams) {
  console.log('\n⬆️  Uploading zip to S3...');

  const zipBuffer = fs.readFileSync(ZIP_PATH);
  const { endpointURL, formData } = uploadParams;

  if (!endpointURL) throw new Error('No upload URL returned from DA API');

  // Build multipart/form-data manually (no external deps)
  const boundary = '----APSFormBoundary' + Date.now().toString(16);
  const parts = [];

  // formData fields (key, policy, x-amz-*, etc.)
  for (const [key, value] of Object.entries(formData || {})) {
    parts.push(
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="${key}"\r\n\r\n` +
      `${value}\r\n`
    );
  }

  // The actual file — MUST be last field in S3 pre-signed POST
  const fileHeader =
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="file"; filename="package.zip"\r\n` +
    `Content-Type: application/octet-stream\r\n\r\n`;

  const closing = `\r\n--${boundary}--\r\n`;

  const bodyParts  = parts.map(p => Buffer.from(p, 'utf8'));
  const filePrefix = Buffer.from(fileHeader, 'utf8');
  const fileSuffix = Buffer.from(closing, 'utf8');
  const totalBody  = Buffer.concat([...bodyParts, filePrefix, zipBuffer, fileSuffix]);

  const url = new URL(endpointURL);

  await new Promise((resolve, reject) => {
    const req = https.request({
      hostname: url.hostname,
      path:     url.pathname + url.search,
      method:   'POST',
      headers: {
        'Content-Type':   `multipart/form-data; boundary=${boundary}`,
        'Content-Length': totalBody.length,
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          console.log('   ✅ Zip uploaded successfully');
          resolve();
        } else {
          reject(new Error(`S3 upload failed: HTTP ${res.statusCode}: ${data}`));
        }
      });
    });
    req.on('error', reject);
    req.write(totalBody);
    req.end();
  });
}

async function setAlias(token, version) {
  console.log(`\n🏷️  Setting alias '${ALIAS}' → version ${version}...`);

  // Try PATCH first (update existing alias)
  try {
    await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/aliases/${ALIAS}`,
      method:   'PATCH',
      headers:  daHeaders(token),
    }, JSON.stringify({ version }));
    console.log(`   ✅ Alias '${ALIAS}' updated`);
    return;
  } catch (patchErr) {
    // 404 = alias doesn't exist yet, create it
    if (!patchErr.message.includes('404')) throw patchErr;
  }

  await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/aliases`,
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: ALIAS, version }));

  console.log(`   ✅ Alias '${ALIAS}' created`);
}

// ─── Main ──────────────────────────────────────────────────────────────────
(async () => {
  console.log('═══════════════════════════════════════════');
  console.log(' APS AppBundle Publisher — RevitExtractor  ');
  console.log('═══════════════════════════════════════════');
  console.log(`Bundle : ${BUNDLE_NAME}`);
  console.log(`Engine : ${ENGINE_ID}`);
  console.log(`Alias  : ${ALIAS}`);
  console.log(`Zip    : ${ZIP_PATH}`);
  console.log('');

  if (!fs.existsSync(ZIP_PATH)) {
    console.error(`❌ Zip not found: ${ZIP_PATH}`);
    process.exit(1);
  }

  const token  = await getToken();
  const bundle = await createOrUpdateAppBundle(token);

  await uploadZip(bundle.uploadParameters);

  await setAlias(token, bundle.version);

  console.log('\n═══════════════════════════════════════════');
  console.log(`✅ Done!  ${BUNDLE_NAME}+${ALIAS}  →  v${bundle.version}`);
  console.log('');
  console.log(`Full activity reference:`);
  console.log(`  <owner>.${BUNDLE_NAME}+${ALIAS}`);
  console.log('═══════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Publish failed:', err.message);
  process.exit(1);
});
