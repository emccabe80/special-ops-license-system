/**
 * Special OPS — Access Control Logic
 * =============================================
 * Placeholder you MUST replace before this works:
 *   BACKEND_URL -> your deployed Apps Script Web App URL (see SETUP_GUIDE.md)
 *
 * Flow on every open:
 *   1. Get the signed-in user's email via Office SSO (silent, no prompt
 *      if they're already signed into Microsoft 365 in Excel).
 *   2. Read this document's stored FileID (a document Setting — persists
 *      inside the file itself, travels with copies).
 *   3. If no FileID exists yet: this is a first-ever open. Mint one from
 *      base name + today's date + this user's email, save it into the
 *      document, and register it with the backend (as "Pending").
 *   4. If a FileID already exists: send it + this user's email + the
 *      document's current URL/name to the backend for a decision.
 *   5. Act on the decision: unlock and show status, or protect every
 *      worksheet so the content can't be used.
 *
 * BASE_NAME identifies which *product* this is (not which instance) —
 * set this once per template you build from this add-in.
 */

const BACKEND_URL = "https://script.google.com/macros/s/AKfycbwuZo3rEHZ0ymXIx3-NSHeqhGkwBmUfBYYC_KgAMVGas0376c0yFtJGk4XFeOA3Q2mb/exec";
const BASE_NAME = "SpecialOps.ReportEngine.EMCREV0";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    runAccessCheck();
  }
});

async function runAccessCheck() {
  setStatus("checking", "Checking authorization\u2026");

  let email;
  try {
    email = await getSignedInEmail();
  } catch (err) {
    return lockWorkbook("Could not verify your Microsoft 365 identity. Make sure you're signed in to Office, then reopen this file.");
  }

  let fileId = await getStoredFileId();
  const path = Office.context.document.url || "";
  const filename = (path.split("/").pop() || "").split("?")[0];

  let result;
  if (!fileId) {
    fileId = mintInstanceId(email);
    await setStoredFileId(fileId);
    result = await callBackend({
      action: "registerInstance",
      fileId, baseName: BASE_NAME, email, path, filename
    });
  } else {
    result = await callBackend({
      action: "checkAccess",
      fileId, baseName: BASE_NAME, email, path, filename
    });
  }

  handleResult(result, fileId, email);
}

function mintInstanceId(email) {
  const today = new Date().toISOString().slice(0, 10);
  return `${BASE_NAME}-${today}-${email}`;
}

async function getSignedInEmail() {
  // Office SSO: returns an Azure AD token for the signed-in user without
  // a manual login prompt, as long as they're already signed into Office.
  const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: false });
  const claims = decodeJwtClaims(token);
  return claims.preferred_username || claims.upn || claims.email;
}

function decodeJwtClaims(token) {
  const payload = token.split(".")[1];
  const decoded = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));
  return JSON.parse(decoded);
}

async function getStoredFileId() {
  return new Promise((resolve) => {
    const settings = Office.context.document.settings;
    resolve(settings.get("specialops_fileId") || null);
  });
}

async function setStoredFileId(fileId) {
  return new Promise((resolve, reject) => {
    const settings = Office.context.document.settings;
    settings.set("specialops_fileId", fileId);
    settings.saveAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        reject(asyncResult.error);
      } else {
        resolve();
      }
    });
  });
}

async function callBackend(payload) {
  try {
    const resp = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // avoids CORS preflight against Apps Script
      body: JSON.stringify(payload)
    });
    return await resp.json();
  } catch (err) {
    // Network/backend unreachable — fail safe, but distinguish this from
    // an explicit "unauthorized" so a dead-zone site doesn't look like a
    // revoked license. See SETUP_GUIDE.md for the offline-grace variant.
    return { status: "unreachable", message: "Could not reach the access-control server. Check your internet connection." };
  }
}

function handleResult(result, fileId, email) {
  switch (result.status) {
    case "authorized":
      setStatus("authorized", "Access confirmed.");
      unlockWorkbook();
      break;
    case "pending_approval":
      setStatus("pending", result.message);
      lockWorkbook(result.message);
      break;
    case "path_mismatch":
      setStatus("pending", result.message);
      lockWorkbook(result.message);
      break;
    case "expired":
      setStatus("unauthorized", result.message);
      lockWorkbook(result.message + " Contact the file owner to renew.");
      break;
    case "unauthorized":
      setStatus("unauthorized", result.message);
      lockWorkbook(result.message);
      break;
    default:
      setStatus("unauthorized", result.message || "Access could not be verified.");
      lockWorkbook(result.message || "Access could not be verified.");
  }
  document.getElementById("meta").textContent = `${email} \u00B7 ${fileId}`;
}

function setStatus(cls, text) {
  const el = document.getElementById("status");
  el.className = "status " + cls;
  el.textContent = text;
}

async function lockWorkbook(reason) {
  document.getElementById("detail").textContent = reason || "";
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      sheets.items.forEach((sheet) => {
        sheet.protection.protect({
          allowFormatCells: false, allowInsertRows: false, allowDeleteRows: false,
          allowSort: false, allowAutoFilter: false, allowPivotTables: false
        });
      });
      await context.sync();
    });
  } catch (err) {
    // Sheets may already be protected — non-fatal.
  }
}

async function unlockWorkbook() {
  document.getElementById("detail").textContent = "";
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      sheets.items.forEach((sheet) => sheet.protection.unprotect());
      await context.sync();
    });
  } catch (err) {
    // Non-fatal — sheet may not have been protected.
  }
}
