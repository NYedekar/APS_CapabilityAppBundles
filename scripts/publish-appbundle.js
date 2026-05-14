#!/usr/bin/env node
const fs   = require('fs');
const path = require('path');
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const BUNDLE_NAME   = process.env.BUNDLE_NAME    || 'RevitExtractor';
const ENGINE_VER    = process.env.ENGINE_VERSION || '2024';
const ALIAS         = process.env.ALIAS          || 'prod';
const ZIP_PATH      = process.env.BUNDLE_ZIP     || path.join(process.cwd(), `${BUNDLE_NAME}.zip`);
const ENGINE_ID     = `Autodesk.Revit+${ENGINE_VER}`;

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
        try { 
          const parsed = JSON.parse(data);
          resolve({ status, body: parsed });
        } catch { 
          resolve({ status, body: data }); 
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

function daHeaders(token) {
  return { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
}

async function createNewVersion(token) {
  console.log(`\n📦 Creating new AppBundle version: ${BUNDLE_NAME}`);

  const payload = JSON.stringify({
    engine:      ENGINE_ID,
    description: 'Extracts all Revit instance and type parameters to result.json and result.csv',
  });

  // First try: POST to /appbundles (creates bundle + version 1 if brand new)
  // If 409: bundle exists, POST to /appbundles/{id}/versions to add a new version
  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/appbundles`,
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: BUNDLE_NAME, engine: ENGINE_ID, description: 'Extracts all Revit parameters' }));

  if (res.status === 409) {
    console.log('   Bundle exists — creating new version...');
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/versions`,
      method:   'POST',
      headers:  daHeaders(token),
    }, payload);
  }

  if (res.status < 200 || res.status >= 300) {
    throw new Error(`HTTP ${res.status}: ${JSON.stringify(res.body)}`);
  }

  console.log(`   ✅ AppBundle version ${res.body.version} ready`);
  return res.body;
}

async function uploadZip(uploadParams) {
  console.log('\n⬆️  Uploading zip to S3...');
  const zipBuffer = fs.readFileSync(ZIP_PATH);
  const { endpointURL, formData } = uploadParams;
  if (!endpointURL) throw new Error('No upload URL returned from DA API');

  const boundary = '----APSFormBoundary' + Date.now().toString(16);
  const parts = [];

  for (const [key, value] of Object.entries(formData || {})) {
    parts.push(
      `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="${key}"\r\n\r\n` +
      `${value}\r\n`
    );
  }

  const fileHeader =
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="file"; filename="package.zip"\r\n` +
    `Content-Type: application/octet-stream\r\n\r\n`;
  const closing = `\r\n--${boundary}--\r\n`;

  const totalBody = Buffer.concat([
    ...parts.map(p => Buffer.from(p, 'utf8')),
    Buffer.from(fileHeader, 'utf8'),
    zipBuffer,
    Buffer.from(closing, 'utf8'),
  ]);

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
          console.log('   ✅ Zip uploaded');
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
  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/aliases/${ALIAS}`,
    method:   'PATCH',
    headers:  daHeaders(token),
  }, JSON.stringify({ version }));

  if (res.status === 404) {
    // Alias doesn't exist yet — create it
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/aliases`,
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify({ id: ALIAS, version }));
  }

  if (res.status < 200 || res.status >= 300) {
    throw new Error(`Alias failed: HTTP ${res.status}: ${JSON.stringify(res.body)}`);
  }

  console.log(`   ✅ Alias '${ALIAS}' → v${version}`);
}

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
  const bundle = await createNewVersion(token);
  await uploadZip(bundle.uploadParameters);
  await setAlias(token, bundle.version);

  console.log('\n═══════════════════════════════════════════');
  console.log(`✅ Done!  ${BUNDLE_NAME}+${ALIAS}  →  v${bundle.version}`);
  console.log('═══════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Publish failed:', err.message);
  process.exit(1);
});
