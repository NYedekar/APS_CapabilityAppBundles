#!/usr/bin/env node
/**
 * publish-activity.js
 *
 * Creates or updates the "ExtractRevitData" Activity in APS Design Automation.
 *
 * The Activity wires the RevitExtractor AppBundle to:
 *   INPUT  rvtFile    — the Revit file to process (.rvt)
 *   OUTPUT resultJson — result.json  (full structured data with all parameters)
 *   OUTPUT resultCsv  — result.csv   (flat table, one row per element)
 *
 * Usage:
 *   APS_CLIENT_ID=xxx APS_CLIENT_SECRET=yyy node scripts/publish-activity.js
 *
 * Reads the same env vars as publish-appbundle.js.
 */

const https = require('https');

const CLIENT_ID     = required('APS_CLIENT_ID');
const CLIENT_SECRET = required('APS_CLIENT_SECRET');
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
        const status = res.statusCode;
        try {
          const parsed = JSON.parse(data);
          if (status >= 200 && status < 300) resolve({ status, body: parsed });
          else reject(new Error(`HTTP ${status}: ${JSON.stringify(parsed)}`));
        } catch {
          if (status >= 200 && status < 300) resolve({ status, body: data });
          else reject(new Error(`HTTP ${status}: ${data}`));
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
  if (!token) throw new Error('No access_token');
  console.log('   ✅ Token acquired');
  return token;
}

function daHeaders(token) {
  return { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
}

async function getNickname(token) {
  const res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     '/da/us-east/v3/forgeapps/me',
    method:   'GET',
    headers:  daHeaders(token),
  });
  return res.body.id;
}

async function publishActivity(token, nickname) {
  console.log(`\n🔧 Creating/updating Activity: ${ACTIVITY_ID}`);

  // Fully-qualified AppBundle reference
  const bundleRef = `${nickname}.${BUNDLE_NAME}+${ALIAS}`;

  const activityDef = {
    id:     ACTIVITY_ID,
    engine: ENGINE_ID,
    appbundles: [ bundleRef ],

    // Command line: launch Revit core console with the input RVT and the AppBundle
    commandLine: [
      `$(engine.path)\\revitcoreconsole.exe /i "$(args[rvtFile].path)" /al "$(appbundles[${BUNDLE_NAME}].path)"`,
    ],

    parameters: {
      // ── INPUT ──────────────────────────────────────────────────────────
      rvtFile: {
        verb:        'get',
        description: 'Revit model to extract (.rvt)',
        required:    true,
        localName:   'input.rvt',
      },

      // ── OUTPUTS ────────────────────────────────────────────────────────
      // result.json — structured report (project info, element counts, full params)
      resultJson: {
        verb:        'put',
        description: 'Extracted data as JSON (all instance + type parameters)',
        required:    false,
        localName:   'result.json',   // must match File.WriteAllText("result.json") in plugin
      },

      // result.csv — flat table, one row per element, all params as columns
      resultCsv: {
        verb:        'put',
        description: 'Extracted data as CSV (one row per element, Type_ prefix for type params)',
        required:    false,
        localName:   'result.csv',
      },
    },

    description: 'Extracts all instance and type parameters from a Revit model. Outputs result.json and result.csv.',
  };

  // POST creates a new version
  let res;
  try {
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     '/da/us-east/v3/activities',
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify(activityDef));
    console.log(`   ✅ Activity version ${res.body.version} created`);
  } catch (err) {
    // If it already exists (409), POST to create a new version of the existing one
    // The DA API treats POST as "create new version", not "upsert" — 409 on id conflict
    throw err;
  }

  return res.body.version;
}

async function setAlias(token, version) {
  console.log(`\n🏷️  Setting alias '${ALIAS}' → version ${version}...`);

  // Try PATCH (update existing alias)
  try {
    await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases/${ALIAS}`,
      method:   'PATCH',
      headers:  daHeaders(token),
    }, JSON.stringify({ version }));
    console.log(`   ✅ Alias updated`);
    return;
  } catch (e) {
    if (!e.message.includes('404')) throw e;
  }

  // 404 → create alias
  await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases`,
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: ALIAS, version }));
  console.log(`   ✅ Alias created`);
}

(async () => {
  console.log('══════════════════════════════════════════');
  console.log(' APS Activity Publisher — ExtractRevitData');
  console.log('══════════════════════════════════════════');

  const token    = await getToken();
  const nickname = await getNickname(token);

  console.log(`\n   Nickname : ${nickname}`);
  console.log(`   Bundle   : ${nickname}.${BUNDLE_NAME}+${ALIAS}`);
  console.log(`   Activity : ${ACTIVITY_ID}`);
  console.log(`   Engine   : ${ENGINE_ID}`);

  const version = await publishActivity(token, nickname);
  await setAlias(token, version);

  console.log('\n══════════════════════════════════════════');
  console.log(`✅ Done!  ${ACTIVITY_ID}+${ALIAS} → v${version}`);
  console.log('');
  console.log('WorkItem example:');
  console.log(`  activityId : "${nickname}.${ACTIVITY_ID}+${ALIAS}"`);
  console.log('  arguments  :');
  console.log('    rvtFile    → { verb: "get",  url: "<signed-rvt-url>" }');
  console.log('    resultJson → { verb: "put",  url: "<signed-json-url>" }');
  console.log('    resultCsv  → { verb: "put",  url: "<signed-csv-url>" }');
  console.log('══════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Activity publish failed:', err.message);
  process.exit(1);
});
