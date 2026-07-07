#!/usr/bin/env node
// Custom activity publisher for UpdateTitleBlock (AutoCADTitleBlockUpdater).
//
// Write-back activity, single verb APSTITLEBLOCK, self-selecting two modes:
//   - EXTRACT: supply only inputFile → bundle emits the title-block schema (result.json).
//   - UPDATE : supply a changes payload (inline via params.json, or an uploaded
//              changes.dat file) → bundle writes result.dwg + result.json.
//
// Parameters:
//   inputFile   (get, required)  — the DWG                 (localName: input.dwg)
//   params      (get, optional)  — control + inline changes (localName: params.json)
//   changesFile (get, optional)  — uploaded JSON/CSV changes (localName: changes.dat)
//   resultDwg   (put, optional)  — modified drawing          (localName: result.dwg)
//   resultJson  (put, optional)  — schema or update summary  (localName: result.json)
//
// Modeled on scripts/publish-cad-standards-activity.js (multi optional get-args).
const https = require('https');

const CLIENT_ID           = required('APS_CLIENT_ID');
const CLIENT_SECRET       = required('APS_CLIENT_SECRET');
const NICKNAME            = required('APS_NICKNAME');
const BUNDLE_NAME         = process.env.BUNDLE_NAME     || 'AutoCADTitleBlockUpdater';
const ALIAS               = process.env.ALIAS           || 'prod';
const ACTIVITY_ID         = process.env.ACTIVITY_ID     || 'UpdateTitleBlock';
// APSTITLEBLOCK avoids any blocked/built-in AutoCAD command that shadows plugins in accoreconsole.
const COMMAND             = process.env.COMMAND         || 'APSTITLEBLOCK';
const ENGINE_VERSION_HINT = process.env.ENGINE_VERSION  || null;

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
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Authorization': `Basic ${auth}` },
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
    const res = await request({ hostname: 'developer.api.autodesk.com', path: nextPage, method: 'GET',
      headers: { 'Authorization': `Bearer ${token}` } });
    if (res.status !== 200) throw new Error(`Engine list HTTP ${res.status}`);
    allEngines = allEngines.concat(res.body.data || []);
    nextPage = res.body.paginationToken
      ? `/da/us-east/v3/engines?page[limit]=100&page[cursor]=${res.body.paginationToken}` : null;
  }
  const acEngines = allEngines
    .filter(e => /^Autodesk\.AutoCAD\+\d+_\d+$/.test(e))
    .sort((a, b) => {
      const [majA, minA] = a.split('+')[1].split('_').map(Number);
      const [majB, minB] = b.split('+')[1].split('_').map(Number);
      return majB !== majA ? majB - majA : minB - minA;
    });
  console.log(`   Available AutoCAD engines: ${acEngines.join(', ')}`);
  if (!acEngines.length) throw new Error('No AutoCAD engines found.');
  console.log(`   ✅ Selected: ${acEngines[0]}`);
  return acEngines[0];
}

function buildActivityDef(engineId) {
  return {
    engine:     engineId,
    appbundles: [`${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`],
    commandLine: [
      `$(engine.path)\\accoreconsole.exe /i "$(args[inputFile].path)" /al "$(appbundles[${BUNDLE_NAME}].path)" /s "$(settings[script].path)"`,
    ],
    settings: {
      script: { value: COMMAND + '\n' },
    },
    parameters: {
      // Required: the DWG to read/modify (opened via /i → active document).
      inputFile: {
        verb:        'get',
        description: 'Input DWG file with title blocks',
        required:    true,
        localName:   'input.dwg',
      },
      // Optional: control + inline changes payload (chat/webapp path).
      //   { mode?, titleBlockName?, layoutScope?, changes:{ TAG: value, ... } }
      // Absent or no "changes" → EXTRACT. Present with "changes" → UPDATE.
      params: {
        verb:        'get',
        description: 'Control JSON: { mode?, titleBlockName?, layoutScope?, changes:{TAG:value} }',
        required:    false,
        localName:   'params.json',
      },
      // Optional: uploaded changes file — webapp JSON (same shape) or CSV.
      changesFile: {
        verb:        'get',
        description: 'Uploaded changes: JSON (same shape as params) or CSV (tag,value)',
        required:    false,
        localName:   'changes.dat',
      },
      // Modified drawing (UPDATE only, but declared for both).
      resultDwg: {
        verb:        'put',
        description: 'Modified drawing with stamped title-block values (UPDATE mode)',
        required:    false,
        localName:   'result.dwg',
      },
      // Smoke gate / primary result: EXTRACT schema or UPDATE summary.
      resultJson: {
        verb:        'put',
        description: 'EXTRACT: {ok,mode:"extract",needsInput,rerunVerb,titleBlocks[]}. UPDATE: {ok,mode:"update",updatedCount,source,changesApplied[],unmapped[]}.',
        required:    false,
        localName:   'result.json',
      },
    },
    description: 'Extract or update AutoCAD title-block attributes across paper-space layouts. Self-selects extract vs update from inputs; UPDATE saves result.dwg.',
  };
}

async function deleteAndRecreate(token, def) {
  console.log('   ⚠️  100-version limit — deleting and recreating activity...');
  const del = await request({ hostname: 'developer.api.autodesk.com',
    path: `/da/us-east/v3/activities/${ACTIVITY_ID}`, method: 'DELETE', headers: daHeaders(token) });
  if (del.status !== 204 && del.status !== 200)
    throw new Error(`Delete failed HTTP ${del.status}: ${JSON.stringify(del.body)}`);
  return request({ hostname: 'developer.api.autodesk.com', path: '/da/us-east/v3/activities', method: 'POST',
    headers: daHeaders(token) }, JSON.stringify({ id: ACTIVITY_ID, ...def }));
}

async function publishActivity(token, engineId) {
  console.log(`\n🔧 Publishing Activity: ${ACTIVITY_ID}`);
  const def = buildActivityDef(engineId);

  let res = await request({ hostname: 'developer.api.autodesk.com',
    path: '/da/us-east/v3/activities', method: 'POST', headers: daHeaders(token) },
    JSON.stringify({ id: ACTIVITY_ID, ...def }));

  if (res.status === 409) {
    console.log('   Activity exists — adding new version...');
    res = await request({ hostname: 'developer.api.autodesk.com',
      path: `/da/us-east/v3/activities/${ACTIVITY_ID}/versions`, method: 'POST', headers: daHeaders(token) },
      JSON.stringify(def));
    if (res.status === 403) res = await deleteAndRecreate(token, def);
  } else if (res.status === 403) {
    res = await deleteAndRecreate(token, def);
  }

  if (res.status < 200 || res.status >= 300)
    throw new Error(`HTTP ${res.status}: ${JSON.stringify(res.body)}`);

  console.log(`   ✅ Activity v${res.body.version} ready`);
  return res.body.version;
}

async function setAlias(token, version) {
  console.log(`\n🏷️  Setting alias '${ALIAS}' → v${version}...`);
  let res = await request({ hostname: 'developer.api.autodesk.com',
    path: `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases/${ALIAS}`, method: 'PATCH', headers: daHeaders(token) },
    JSON.stringify({ version }));
  if (res.status === 404) {
    res = await request({ hostname: 'developer.api.autodesk.com',
      path: `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases`, method: 'POST', headers: daHeaders(token) },
      JSON.stringify({ id: ALIAS, version }));
  }
  if (res.status < 200 || res.status >= 300)
    throw new Error(`Alias failed HTTP ${res.status}: ${JSON.stringify(res.body)}`);
  console.log(`   ✅ Alias '${ALIAS}' → v${version}`);
}

(async () => {
  console.log('══════════════════════════════════════════════════════');
  console.log(` APS Activity Publisher — UpdateTitleBlock`);
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
  console.log(`   Args (get):  inputFile (req), params (opt), changesFile (opt)`);
  console.log(`   Args (put):  resultDwg, resultJson`);
  console.log('══════════════════════════════════════════════════════');
})().catch(err => {
  console.error('\n❌ Activity publish failed:', err.message);
  process.exit(1);
});
