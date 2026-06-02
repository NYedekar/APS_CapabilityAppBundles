/**
 * publish-param-updater-activity.js
 *
 * Creates or updates the APS Design Automation Activity for RevitParameterUpdater.
 *
 * Activity: UpdateRevitParameters
 *   Engine:   Autodesk.Revit+2026
 *   Bundle:   NICKNAME.RevitParameterUpdater+prod
 *
 * Parameters:
 *   rvtFile      (get, required)  — the Revit model to modify     (localName: input.rvt)
 *   paramsInput  (get, required)  — CSV / XLSX / JSON change data  (localName: params_input.dat)
 *   params       (get, required)  — control JSON { "inputMode": "..." } (localName: params.json)
 *   resultRvt    (put, optional)  — modified model output          (localName: result.rvt)
 *   resultJson   (put, optional)  — change summary JSON            (localName: result.json)
 */

"use strict";

const https = require("https");

const CLIENT_ID     = process.env.APS_CLIENT_ID;
const CLIENT_SECRET = process.env.APS_CLIENT_SECRET;
const NICKNAME      = process.env.APS_NICKNAME;
const BUNDLE_NAME   = process.env.BUNDLE_NAME   || "RevitParameterUpdater";
const ENGINE_VER    = process.env.ENGINE_VERSION || "2026";
const ALIAS         = process.env.ALIAS         || "prod";
const ACTIVITY_ID   = process.env.ACTIVITY_ID   || "UpdateRevitParameters";

const ENGINE_ID  = `Autodesk.Revit+${ENGINE_VER}`;
const BUNDLE_REF = `${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`;
const BASE       = "developer.api.autodesk.com";

function validate() {
  const missing = [];
  if (!CLIENT_ID)     missing.push("APS_CLIENT_ID");
  if (!CLIENT_SECRET) missing.push("APS_CLIENT_SECRET");
  if (!NICKNAME)      missing.push("APS_NICKNAME");
  if (missing.length) throw new Error(`Missing required env vars: ${missing.join(", ")}`);
  if (BUNDLE_REF.startsWith("undefined."))
    throw new Error(`BUNDLE_REF is "${BUNDLE_REF}" — APS_NICKNAME is undefined`);
}

function httpsRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      let data = "";
      res.on("data", c => (data += c));
      res.on("end", () => resolve({ status: res.statusCode, body: tryParse(data) }));
    });
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

function tryParse(text) {
  try { return JSON.parse(text); } catch { return text; }
}

async function getToken() {
  const body = `client_id=${encodeURIComponent(CLIENT_ID)}&client_secret=${encodeURIComponent(CLIENT_SECRET)}&grant_type=client_credentials&scope=code%3Aall`;
  const res = await httpsRequest({
    hostname: BASE,
    path: "/authentication/v2/token",
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", "Content-Length": Buffer.byteLength(body) },
  }, body);
  if (res.status !== 200) throw new Error(`Auth failed (${res.status}): ${JSON.stringify(res.body)}`);
  console.log("✔ Token obtained");
  return res.body.access_token;
}

function post(token, path, payload) {
  const body = JSON.stringify(payload);
  return httpsRequest({
    hostname: BASE, path, method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) },
  }, body);
}

function patch(token, path, payload) {
  const body = JSON.stringify(payload);
  return httpsRequest({
    hostname: BASE, path, method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) },
  }, body);
}

function del(token, path) {
  return httpsRequest({
    hostname: BASE, path, method: "DELETE",
    headers: { Authorization: `Bearer ${token}` },
  });
}

function buildActivityDef(activityId) {
  return {
    id:          activityId,
    engine:      ENGINE_ID,
    appbundles:  [BUNDLE_REF],
    commandLine: [
      `$(engine.path)\\revitcoreconsole.exe /i "$(args[rvtFile].path)" /al "$(appbundles[${BUNDLE_NAME}].path)"`
    ],
    parameters: {
      rvtFile: {
        verb:      "get",
        localName: "input.rvt",
        required:  true,
        description: "Revit model (.rvt) to modify"
      },
      paramsInput: {
        verb:      "get",
        localName: "params_input.dat",
        required:  true,
        description: "Change requests: CSV (columns: ElementId?, ElementName, Category?, Parameter, NewValue) or XLSX (same columns) or JSON array [{ elementId?, elementName, category?, parameter, value }]"
      },
      params: {
        verb:      "get",
        localName: "params.json",
        required:  true,
        description: JSON.stringify({
          inputMode: "delta_file | full_file | text_input",
          "_notes": "delta_file: apply all rows; full_file: only rows where NewValue != current value; text_input: same as delta_file (input is JSON from chat)"
        })
      },
      resultRvt: {
        verb:      "put",
        localName: "result.rvt",
        required:  false,
        description: "Modified Revit model. Always written even if 0 changes were applied."
      },
      resultJson: {
        verb:      "put",
        localName: "result.json",
        required:  false,
        description: "Change summary: { ok, summary: { total, applied, skipped, errors }, changes: [ { elementId, elementName, parameter, oldValue, newValue, status, reason? } ] }"
      }
    }
  };
}

async function deleteAndRecreate(token, activityId, def) {
  console.log(`  ⚠️  100-version limit — deleting and recreating '${activityId}'...`);
  const r = await del(token, `/da/us-east/v3/activities/${activityId}`);
  if (r.status !== 204 && r.status !== 200)
    throw new Error(`Delete failed (${r.status}): ${JSON.stringify(r.body)}`);
  return post(token, "/da/us-east/v3/activities", def);
}

async function publishActivity(token) {
  console.log(`\n── Activity: ${NICKNAME}.${ACTIVITY_ID} ──`);
  const def = buildActivityDef(ACTIVITY_ID);

  let res = await post(token, "/da/us-east/v3/activities", def);
  let version;

  if (res.status === 200 || res.status === 201) {
    console.log(`✔ Created Activity '${ACTIVITY_ID}' (v${res.body.version})`);
    version = res.body.version;
  } else if (res.status === 409) {
    console.log(`  Activity exists — adding new version...`);
    const { id: _, ...versionDef } = def;
    res = await post(token, `/da/us-east/v3/activities/${ACTIVITY_ID}/versions`, versionDef);
    if (res.status === 403) res = await deleteAndRecreate(token, ACTIVITY_ID, def);
    if (res.status !== 200 && res.status !== 201)
      throw new Error(`Version failed (${res.status}): ${JSON.stringify(res.body)}`);
    console.log(`✔ New version of '${ACTIVITY_ID}': v${res.body.version}`);
    version = res.body.version;
  } else if (res.status === 403) {
    res = await deleteAndRecreate(token, ACTIVITY_ID, def);
    if (res.status !== 200 && res.status !== 201)
      throw new Error(`Recreate failed (${res.status}): ${JSON.stringify(res.body)}`);
    console.log(`✔ Recreated '${ACTIVITY_ID}' v${res.body.version}`);
    version = res.body.version;
  } else {
    throw new Error(`Unexpected status (${res.status}): ${JSON.stringify(res.body)}`);
  }

  // Set / update alias
  let ar = await patch(token, `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases/${ALIAS}`, { version });
  if (ar.status === 200) {
    console.log(`✔ Alias '${ALIAS}' → v${version}`);
  } else if (ar.status === 404) {
    ar = await post(token, `/da/us-east/v3/activities/${ACTIVITY_ID}/aliases`, { id: ALIAS, version });
    if (ar.status !== 200 && ar.status !== 201)
      throw new Error(`Alias create failed (${ar.status}): ${JSON.stringify(ar.body)}`);
    console.log(`✔ Alias '${ALIAS}' created → v${version}`);
  } else {
    throw new Error(`Alias failed (${ar.status}): ${JSON.stringify(ar.body)}`);
  }
}

async function main() {
  console.log("=== publish-param-updater-activity.js ===");
  console.log(`  Engine  : ${ENGINE_ID}`);
  console.log(`  Bundle  : ${BUNDLE_REF}`);
  console.log(`  Activity: ${ACTIVITY_ID}`);
  validate();
  const token = await getToken();
  await publishActivity(token);
  console.log(`\n✅ Activity published: ${NICKNAME}.${ACTIVITY_ID}+${ALIAS}  (engine: ${ENGINE_ID})`);
  console.log(`   Bundle ref: ${BUNDLE_REF}`);
}

main().catch(err => { console.error(`\n❌ ${err.message}`); process.exit(1); });
