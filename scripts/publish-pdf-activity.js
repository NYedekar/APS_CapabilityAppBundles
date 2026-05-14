/**
 * publish-activity.js
 *
 * Creates or updates two APS Design Automation Activities:
 *
 *   1. ExportSheetsToPDF — caller provides params.json with { "operation": "sheets" }
 *   2. ExportViewsToPDF  — caller provides params.json with { "operation": "views"  }
 *
 * Both Activities reference the same AppBundle: NICKNAME.RevitPDFExport+prod
 *
 * Activity I/O parameters:
 *   rvtFile   (get, required) — the Revit model to process
 *   params    (get, optional) — params.json controlling operation, filters, paper size
 *   resultZip (put, optional) — result.zip containing all exported PDFs
 *   resultJson(put, optional) — result.json manifest (success/count/errors)
 *
 * CRITICAL: NICKNAME must come from process.env.APS_NICKNAME — never from an API call.
 *   A missing env var returns undefined → appbundles becomes "undefined.RevitPDFExport+prod".
 */

"use strict";

const https = require("https");

// ── Config ────────────────────────────────────────────────────────────────────
const CLIENT_ID      = process.env.APS_CLIENT_ID;
const CLIENT_SECRET  = process.env.APS_CLIENT_SECRET;
const NICKNAME       = process.env.APS_NICKNAME;       // ← read directly, NEVER from API
const BUNDLE_NAME    = process.env.BUNDLE_NAME    || "RevitPDFExport";
const ENGINE_VER     = process.env.ENGINE_VERSION || "2026";
const ALIAS          = process.env.ALIAS          || "prod";
const ACTIVITY_SHEETS = process.env.ACTIVITY_SHEETS || "ExportSheetsToPDF";
const ACTIVITY_VIEWS  = process.env.ACTIVITY_VIEWS  || "ExportViewsToPDF";

const ENGINE_ID  = `Autodesk.Revit+${ENGINE_VER}`;
const BUNDLE_REF = `${NICKNAME}.${BUNDLE_NAME}+${ALIAS}`;  // e.g. "MyNick.RevitPDFExport+prod"
const BASE_URL   = "developer.api.autodesk.com";

// ── Validation ────────────────────────────────────────────────────────────────
function validate() {
  const missing = [];
  if (!CLIENT_ID)    missing.push("APS_CLIENT_ID");
  if (!CLIENT_SECRET) missing.push("APS_CLIENT_SECRET");
  if (!NICKNAME)     missing.push("APS_NICKNAME");
  if (missing.length)
    throw new Error(`Missing required env vars: ${missing.join(", ")}`);

  // Catch the silent failure from a missing env var before it causes
  // a confusing "Activity appbundles field shows undefined.RevitPDFExport+prod" error
  if (BUNDLE_REF.startsWith("undefined."))
    throw new Error(
      `BUNDLE_REF resolved to "${BUNDLE_REF}" — APS_NICKNAME env var is missing or undefined`
    );
}

// ── HTTP helpers ──────────────────────────────────────────────────────────────
function httpsRequest(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", (c) => (data += c));
      res.on("end", () =>
        resolve({ status: res.statusCode, body: data ? tryParse(data) : null })
      );
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
    hostname: BASE_URL,
    path:     "/authentication/v2/token",
    method:   "POST",
    headers:  {
      "Content-Type":   "application/x-www-form-urlencoded",
      "Content-Length": Buffer.byteLength(body)
    }
  }, body);

  if (res.status !== 200)
    throw new Error(`Auth failed (${res.status}): ${JSON.stringify(res.body)}`);

  console.log("✔ Token obtained");
  return res.body.access_token;
}

async function apiPost(token, apiPath, payload) {
  const body = JSON.stringify(payload);
  return httpsRequest({
    hostname: BASE_URL,
    path:     apiPath,
    method:   "POST",
    headers:  {
      "Authorization":  `Bearer ${token}`,
      "Content-Type":   "application/json",
      "Content-Length": Buffer.byteLength(body)
    }
  }, body);
}

async function apiPatch(token, apiPath, payload) {
  const body = JSON.stringify(payload);
  return httpsRequest({
    hostname: BASE_URL,
    path:     apiPath,
    method:   "PATCH",
    headers:  {
      "Authorization":  `Bearer ${token}`,
      "Content-Type":   "application/json",
      "Content-Length": Buffer.byteLength(body)
    }
  }, body);
}

// ── Activity definition builder ───────────────────────────────────────────────
/**
 * Builds the Activity payload shared by both operations.
 * The operation type is controlled by the WorkItem caller via params.json.
 *
 * Command line breakdown:
 *   revitcoreconsole.exe
 *     /i "$(args[rvtFile].path)"         ← the input .rvt file
 *     /al "$(appbundles[RevitPDFExport].path)"   ← the AppBundle to load
 *
 * Parameters:
 *   rvtFile   - required input  (localName: input.rvt)
 *   params    - optional input  (localName: params.json)
 *   resultZip - optional output (localName: result.zip)
 *   resultJson- optional output (localName: result.json)
 */
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
        description: "The Revit model (.rvt) to export from"
      },
      params: {
        verb:      "get",
        localName: "params.json",
        required:  false,
        description: JSON.stringify({
          operation:    "sheets | views",
          sheetNumbers: ["A1", "A2"],
          viewNames:    ["Level 1 Floor Plan"],
          combine:      false,
          paperSize:    "Default | A4 | A3 | A2 | A1 | A0 | Letter | Tabloid"
        })
      },
      resultZip: {
        verb:      "put",
        localName: "result.zip",
        required:  false,
        description: "Zip archive of all exported PDF files"
      },
      resultJson: {
        verb:      "put",
        localName: "result.json",
        required:  false,
        description: "Manifest: { operation, success, exportedCount, exportedItems, errors }"
      }
    }
  };
}

// ── Create or update one Activity ────────────────────────────────────────────
async function publishActivity(token, activityId) {
  console.log(`\n── Activity: ${NICKNAME}.${activityId} ──`);

  const def = buildActivityDef(activityId);

  // Try to create
  let res = await apiPost(token, "/da/us-east/v3/activities", def);

  let version;

if (res.status === 200 || res.status === 201) {
    console.log(`✔ Created Activity '${activityId}' (version ${res.body.version})`);
    version = res.body.version;
  } else if (res.status === 409) {
    // Exists — create a new version
    console.log(`  Activity '${activityId}' exists — creating new version...`);
    // Remove id from payload when creating a version
    const { id: _, ...versionDef } = def;
    res = await apiPost(token,
      `/da/us-east/v3/activities/${activityId}/versions`,
      versionDef
    );
    if (res.status !== 200 && res.status !== 201)
      throw new Error(
        `Failed to version Activity '${activityId}' (${res.status}): ${JSON.stringify(res.body)}`
      );
    console.log(`✔ New version of '${activityId}': ${res.body.version}`);
    version = res.body.version;
  } else {
    throw new Error(
      `Unexpected status creating Activity '${activityId}' (${res.status}): ${JSON.stringify(res.body)}`
    );
  }

  // Set / update the alias
  await setActivityAlias(token, activityId, version);
  return version;
}

async function setActivityAlias(token, activityId, version) {
  let res = await apiPatch(token,
    `/da/us-east/v3/activities/${activityId}/aliases/${ALIAS}`,
    { version }
  );

  if (res.status === 200) {
    console.log(`✔ Alias '${ALIAS}' → version ${version}`);
  } else if (res.status === 404) {
    console.log(`  Creating alias '${ALIAS}'...`);
    res = await apiPost(token,
      `/da/us-east/v3/activities/${activityId}/aliases`,
      { id: ALIAS, version }
    );
    if (res.status !== 200 && res.status !== 201)
      throw new Error(
        `Failed to create alias on '${activityId}' (${res.status}): ${JSON.stringify(res.body)}`
      );
    console.log(`✔ Alias '${ALIAS}' created → version ${version}`);
  } else {
    throw new Error(
      `Unexpected status setting alias on '${activityId}' (${res.status}): ${JSON.stringify(res.body)}`
    );
  }
}

// ── Entry point ───────────────────────────────────────────────────────────────
async function main() {
  console.log("=== publish-activity.js ===");
  console.log(`  Engine   : ${ENGINE_ID}`);
  console.log(`  Bundle   : ${BUNDLE_REF}`);
  console.log(`  Alias    : ${ALIAS}`);
  console.log(`  Activity1: ${NICKNAME}.${ACTIVITY_SHEETS}`);
  console.log(`  Activity2: ${NICKNAME}.${ACTIVITY_VIEWS}`);

  validate();

  const token = await getToken();

  await publishActivity(token, ACTIVITY_SHEETS);
  await publishActivity(token, ACTIVITY_VIEWS);

  console.log("\n✅ Both Activities published:");
  console.log(`   ${NICKNAME}.${ACTIVITY_SHEETS}+${ALIAS}  (engine: ${ENGINE_ID})`);
  console.log(`   ${NICKNAME}.${ACTIVITY_VIEWS}+${ALIAS}   (engine: ${ENGINE_ID})`);
  console.log(`\n  Both reference AppBundle: ${BUNDLE_REF}`);
  console.log("  Operation is controlled by the params.json the WorkItem caller provides.");
}

main().catch((err) => {
  console.error(`\n❌ publish-activity.js failed: ${err.message}`);
  process.exit(1);
});
