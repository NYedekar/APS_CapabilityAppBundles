#!/usr/bin/env node
const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
const NICKNAME      = required('APS_NICKNAME');
const BUNDLE_NAME   = process.env.BUNDLE_NAME    || 'RevitExtractor';
const ENGINE_VER    = process.env.ENGINE_VERSION || '2024';
const ALIAS         = process.env.ALIAS          || 'prod';
const ACTIVITY_ID   = process.env.ACTIVITY_ID    || 'ExtractRevitData';
const ENGINE_ID     = `Autodesk.Revit+${ENGINE_VER}`;

function required(name) {
  const v = process.env[name];
  if (!v) { console.error(`❌ Missing env var: ${name}`); process.exit(1); }
  return v;
}

function request(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, body: JSON.parse(data) });
        } catch {
          resolve({ status: res.statusCode, body: data });
        }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getToken() {
  console.log('🔑 Getting token...');
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

function buildActivityDef() {
  return {
    engine:     ENGINE_ID,
    appbundles: [`${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`],
    commandLine: [
      `$(engine.path)\\revitcoreconsole.exe /i "$(args[rvtFile].path)" /al "$(appbundles[${BUNDLE_NAME}].path)"`,
    ],
    parameters: {
      rvtFile: {
        verb:        'get',
        description: 'Input RVT file',
        required:    true,
        localName:   'input.rvt',
      },
      resultJson: {
        verb:        'put',
        description: 'Extracted parameters as JSON',
        required:    false,
        localName:   'result.json',
      },
      resultCsv: {
        verb:        'put',
        description: 'Extracted parameters as CSV',
        required:    false,
        localName:   'result.csv',
      },
    },
    description: 'Extracts all Revit instance and type parameters to JSON and CSV',
  };
}

async function publishActivity(token) {
  console.log(`\n🔧 Creating/updating Activity: ${ACTIVITY_ID}`);

  const def = buildActivityDef();

  // Try creating new (first time)
  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/activities`,
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: ACTIVITY_ID, ...def }));

  if (res.status === 409) {
    // Activity exists — create a new version instead
    console.log('   Activity exists — creating new version...');
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/versions`,
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify(def));
  }

  if (res.status < 200 || res.status >= 300) {
    throw new Error(`HTTP ${res.status}: ${JSON.stringify(res.body)}`);
  }

  console.log(`   ✅ Activity version ${res.body.version} ready`);
  return res.body.version;
}

async function setAlias(token, version) {
  console.log(`\n🏷️  Setting alias '${ALIAS}' → version ${version}...`);

  // Try PATCH (update existing alias)
  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases/${ALIAS}`,
    method:   'PATCH',
    headers:  daHeaders(token),
  }, JSON.stringify({ version }));

  if (res.status === 404) {
    // Alias doesn't exist — create it
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases`,
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
  console.log('══════════════════════════════════════════');
  console.log(' APS Activity Publisher — ExtractRevitData');
  console.log('══════════════════════════════════════════');
  console.log(`   Nickname : ${NICKNAME}`);
  console.log(`   Bundle   : ${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`);
  console.log(`   Activity : ${ACTIVITY_ID}`);
  console.log(`   Engine   : ${ENGINE_ID}`);

  const token   = await getToken();
  const version = await publishActivity(token);
  await setAlias(token, version);

  console.log('\n══════════════════════════════════════════');
  console.log(`✅ Done!  ${ACTIVITY_ID}+${ALIAS} → v${version}`);
  console.log(`   AppBundle: ${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`);
  console.log('══════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Activity publish failed:', err.message);
  process.exit(1);
});
