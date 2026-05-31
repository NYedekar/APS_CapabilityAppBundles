#!/usr/bin/env node
const https = require('https');

const CLIENT_ID           = required('APS_CLIENT_ID');
const CLIENT_SECRET       = required('APS_CLIENT_SECRET');
const NICKNAME            = required('APS_NICKNAME');
const BUNDLE_NAME         = process.env.BUNDLE_NAME     || 'AutoCADDrawingMetadataExtractor';
const ALIAS               = process.env.ALIAS           || 'prod';
const ACTIVITY_ID         = process.env.ACTIVITY_ID     || 'ExtractAutoCADDrawingMetadata';
// COMMAND is the AutoCAD command the activity's script invokes. Lets one script
// publish both activities: ExtractAutoCADDrawingMetadata (EXTRACTDWGMETADATA) and
// ExtractAutoCADDrawingMetadataAll (EXTRACTALLDRAWINGMETADATA).
const COMMAND             = process.env.COMMAND         || 'EXTRACTDWGMETADATA';
const ENGINE_VERSION_HINT = process.env.ENGINE_VERSION  || null;
// ENGINE_ID is resolved at runtime via resolveEngineId() — do not hardcode.

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
    if (res.status !== 200) throw new Error(`Engine list HTTP ${res.status}: ${JSON.stringify(res.body)}`);
    allEngines = allEngines.concat(res.body.data || []);
    nextPage = res.body.paginationToken
      ? `/da/us-east/v3/engines?page[limit]=100&page[cursor]=${res.body.paginationToken}`
      : null;
  }
  // Engine ID format: Autodesk.AutoCAD+<major>_<minor>  e.g. Autodesk.AutoCAD+25_1
  const acEngines = allEngines
    .filter(e => /^Autodesk\.AutoCAD\+\d+_\d+$/.test(e))
    .sort((a, b) => {
      const [majA, minA] = a.split('+')[1].split('_').map(Number);
      const [majB, minB] = b.split('+')[1].split('_').map(Number);
      return majB !== majA ? majB - majA : minB - minA;
    });
  console.log(`   Available AutoCAD engines: ${acEngines.join(', ')}`);
  if (acEngines.length === 0) throw new Error('No AutoCAD engines found.');
  console.log(`   ✅ Selected engine: ${acEngines[0]}`);
  return acEngines[0];
}

function buildActivityDef(engineId) {
  return {
    engine:     engineId,
    appbundles: [`${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`],
    // accoreconsole.exe loads the .dwg via /i, loads the bundle via /al, then
    // executes the script via /s. The script just contains the command name.
    // $(settings[script].path) creates a temp file from the inline value.
    commandLine: [
      `$(engine.path)\\accoreconsole.exe /i "$(args[inputFile].path)" /al "$(appbundles[${BUNDLE_NAME}].path)" /s "$(settings[script].path)"`,
    ],
    settings: {
      script: {
        value: COMMAND + '\n',
      },
    },
    parameters: {
      inputFile: {
        verb:        'get',
        description: 'Input DWG file',
        required:    true,
        localName:   '$(inputFile)',
      },
      resultJson: {
        verb:        'put',
        description: 'Extracted DWG metadata as JSON',
        required:    false,
        localName:   'result.json',
      },
    },
    description: 'Extracts all DWG metadata (summary info, layers, blocks, layouts, entity counts, symbol tables) to result.json',
  };
}

async function deleteAndRecreate(token, def) {
  console.log('   ⚠️  100-version limit reached — deleting activity and recreating...');
  const del = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/activities/${ACTIVITY_ID}`,
    method:   'DELETE',
    headers:  daHeaders(token),
  });
  if (del.status !== 204 && del.status !== 200)
    throw new Error(`Delete failed HTTP ${del.status}: ${JSON.stringify(del.body)}`);
  return request({
    hostname: 'developer.api.autodesk.com',
    path:     '/da/us-east/v3/activities',
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: ACTIVITY_ID, ...def }));
}

async function publishActivity(token, engineId) {
  console.log(`\n🔧 Creating/updating Activity: ${ACTIVITY_ID}`);
  const def = buildActivityDef(engineId);

  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     '/da/us-east/v3/activities',
    method:   'POST',
    headers:  daHeaders(token),
  }, JSON.stringify({ id: ACTIVITY_ID, ...def }));

  if (res.status === 409) {
    console.log('   Activity exists — adding new version...');
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/versions`,
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify(def));

    if (res.status === 403) {
      res = await deleteAndRecreate(token, def);
    }
  } else if (res.status === 403) {
    // DA sometimes returns 403 "Maximum number of versions is 100" directly on
    // POST /activities when the activity exists and is already at the limit
    // (instead of 409 → POST /versions → 403). Handle it the same way.
    res = await deleteAndRecreate(token, def);
  }

  if (res.status < 200 || res.status >= 300)
    throw new Error(`HTTP ${res.status}: ${JSON.stringify(res.body)}`);

  console.log(`   ✅ Activity version ${res.body.version} ready`);
  return res.body.version;
}

async function setAlias(token, version) {
  console.log(`\n🏷️  Setting alias '${ALIAS}' → version ${version}...`);

  let res = await request({
    hostname: 'developer.api.autodesk.com',
    path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases/${ALIAS}`,
    method:   'PATCH',
    headers:  daHeaders(token),
  }, JSON.stringify({ version }));

  if (res.status === 404) {
    res = await request({
      hostname: 'developer.api.autodesk.com',
      path:     `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases`,
      method:   'POST',
      headers:  daHeaders(token),
    }, JSON.stringify({ id: ALIAS, version }));
  }

  if (res.status < 200 || res.status >= 300)
    throw new Error(`Alias failed HTTP ${res.status}: ${JSON.stringify(res.body)}`);

  console.log(`   ✅ Alias '${ALIAS}' → v${version}`);
}

(async () => {
  console.log('══════════════════════════════════════════════════════');
  console.log(' APS Activity Publisher — ExtractAutoCADDrawingMetadata');
  console.log('══════════════════════════════════════════════════════');
  console.log(`   Nickname : ${NICKNAME}`);
  console.log(`   Bundle   : ${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`);
  console.log(`   Activity : ${ACTIVITY_ID}`);
  console.log(`   Engine   : ${ENGINE_VERSION_HINT ? 'Autodesk.AutoCAD+' + ENGINE_VERSION_HINT : '(auto-detect)'}`);
  console.log(`   Command  : ${COMMAND}`);

  const token    = await getToken();
  const engineId = await resolveEngineId(token);
  const version  = await publishActivity(token, engineId);
  await setAlias(token, version);

  console.log('\n══════════════════════════════════════════════════════');
  console.log(`✅ Done!  ${ACTIVITY_ID}+${ALIAS} → v${version}`);
  console.log(`   AppBundle: ${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`);
  console.log('══════════════════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Activity publish failed:', err.message);
  process.exit(1);
});
