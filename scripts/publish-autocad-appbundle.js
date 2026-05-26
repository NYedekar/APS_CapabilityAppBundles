#!/usr/bin/env node
const fs    = require('fs');
const path  = require('path');
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const BUNDLE_NAME   = required('BUNDLE_NAME');
const ALIAS         = required('ALIAS');
const ZIP_PATH      = process.env.BUNDLE_ZIP || path.join(process.cwd(), `${BUNDLE_NAME}.zip`);
// ENGINE_VERSION env var is optional — if omitted the script auto-detects the latest AutoCAD engine.
const ENGINE_VERSION_HINT = process.env.ENGINE_VERSION || null;

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
  console.log('🔑 Getting 2-legged token...');
  const auth = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  const res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     '/authentication/v2/token',
    method:   'POST',
    headers: {
      'Content-Type':  'application/x-www-form-urlencoded',
      'Authorization': `Basic ${auth}`,
    },
  }, 'grant_type=client_credentials&scope=code%3Aall');
  const token = res.body.access_token;
  if (!token) throw new Error(`No access_token: ${JSON.stringify(res.body)}`);
  console.log('   ✅ Token acquired');
  return token;
}

function daHeaders(token) {
  return { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
}

// Query the DA engines list and return the latest non-deprecated AutoCAD engine ID.
async function resolveEngineId(token) {
  if (ENGINE_VERSION_HINT) {
    const id = `Autodesk.AutoCAD+${ENGINE_VERSION_HINT}`;
    console.log(`\n🔍 Using ENGINE_VERSION hint: ${id}`);
    return id;
  }

  console.log('\n🔍 Auto-detecting latest AutoCAD engine...');
  let allEngines = [];
  let nextPage = `/da/us-east/v3/engines?page[limit]=100`;

  while (nextPage) {
    const res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     nextPage,
      method:   'GET',
      headers:  { 'Authorization': `Bearer ${token}` },
    });
    if (res.status !== 200) throw new Error(`Engine list failed HTTP ${res.status}: ${JSON.stringify(res.body)}`);
    allEngines = allEngines.concat(res.body.data || []);
    nextPage = res.body.paginationToken
      ? `/da/us-east/v3/engines?page[limit]=100&page[cursor]=${res.body.paginationToken}`
      : null;
  }

  const acEngines = allEngines
    .filter(e => /^Autodesk\.AutoCAD\+\d+$/.test(e))
    .sort((a, b) => {
      const vA = parseInt(a.split('+')[1], 10);
      const vB = parseInt(b.split('+')[1], 10);
      return vB - vA; // descending — highest version first
    });

  console.log(`   Available AutoCAD engines: ${acEngines.join(', ')}`);
  if (acEngines.length === 0) throw new Error('No AutoCAD engines found in DA API response.');

  const latest = acEngines[0];
  console.log(`   ✅ Selected engine: ${latest}`);
  return latest;
}

async function createNewVersion(token, engineId) {
  console.log(`\n📦 Creating AppBundle version: ${BUNDLE_NAME} (${engineId})`);

  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     '/da/us-east/v3/appbundles',
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: BUNDLE_NAME, engine: engineId, description: `${BUNDLE_NAME} plugin` }));

  if (res.status === 409) {
    console.log('   Bundle exists — adding new version...');
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/versions`,
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify({ engine: engineId, description: `${BUNDLE_NAME} plugin` }));
  }

  if (res.status < 200 || res.status >= 300)
    throw new Error(`HTTP ${res.status}: ${JSON.stringify(res.body)}`);

  console.log(`   ✅ AppBundle version ${res.body.version} ready`);
  return res.body;
}

async function uploadZip(uploadParams) {
  console.log('\n⬆️  Uploading zip to S3...');
  const zipBuffer = fs.readFileSync(ZIP_PATH);
  const { endpointURL, formData } = uploadParams;
  if (!endpointURL) throw new Error('No upload URL from DA API');

  const boundary = '----APSFormBoundary' + Date.now().toString(16);
  const parts = Object.entries(formData || {}).map(([k, v]) =>
    `--${boundary}\r\nContent-Disposition: form-data; name="${k}"\r\n\r\n${v}\r\n`
  );
  const fileHeader = `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="package.zip"\r\nContent-Type: application/octet-stream\r\n\r\n`;
  const closing    = `\r\n--${boundary}--\r\n`;

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
        if (res.statusCode >= 200 && res.statusCode < 300) { console.log('   ✅ Zip uploaded'); resolve(); }
        else reject(new Error(`S3 upload failed HTTP ${res.statusCode}: ${data}`));
      });
    });
    req.on('error', reject);
    req.write(totalBody);
    req.end();
  });
}

async function setAlias(token, version) {
  console.log(`\n🏷️  Setting alias '${ALIAS}' → version ${version}...`);

  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/aliases/${ALIAS}`,
    method:   'PATCH',
    headers:  daHeaders(token),
  }, JSON.stringify({ version }));

  if (res.status === 404) {
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/appbundles/${BUNDLE_NAME}/aliases`,
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify({ id: ALIAS, version }));
  }

  if (res.status < 200 || res.status >= 300)
    throw new Error(`Alias failed HTTP ${res.status}: ${JSON.stringify(res.body)}`);

  console.log(`   ✅ Alias '${ALIAS}' → v${version}`);
}

(async () => {
  console.log('═══════════════════════════════════════════════════');
  console.log(` APS AppBundle Publisher — ${BUNDLE_NAME}`);
  console.log('═══════════════════════════════════════════════════');
  console.log(`Bundle : ${BUNDLE_NAME}`);
  console.log(`Alias  : ${ALIAS}`);
  console.log(`Zip    : ${ZIP_PATH}`);
  console.log('');

  if (!fs.existsSync(ZIP_PATH)) {
    console.error(`❌ Zip not found: ${ZIP_PATH}`);
    process.exit(1);
  }

  const token    = await getToken();
  const engineId = await resolveEngineId(token);
  const bundle   = await createNewVersion(token, engineId);
  await uploadZip(bundle.uploadParameters);
  await setAlias(token, bundle.version);

  console.log('\n═══════════════════════════════════════════════════');
  console.log(`✅ Done!  ${BUNDLE_NAME}+${ALIAS}  →  v${bundle.version}`);
  console.log(`   Engine : ${engineId}`);
  console.log('═══════════════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Publish failed:', err.message);
  process.exit(1);
});
