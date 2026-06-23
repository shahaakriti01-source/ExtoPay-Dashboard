import { msalInstance, loginRequest } from "./authConfig";

// ── OneDrive/SharePoint sharing URL for the live pilot Excel file ───────────
const SHARE_URL = "https://mergencompass-my.sharepoint.com/:x:/p/sandeep/IQBxf7rS-YmgRJG5clEQAPwgAcK13A5u8XVoU5QTRZB796c";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function encodeShareUrl(url) {
  const base64 = btoa(unescape(encodeURIComponent(url)))
    .replace(/=/g, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-");
  return "u!" + base64;
}

// Guard against MSAL's "block_nested_popups" error — only one popup interaction
// (login or token-acquisition) is allowed to be in flight at any time.
let popupInFlight = null;

async function getAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) throw new Error("No signed-in account found");

  try {
    const response = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0],
    });
    return response.accessToken;
  } catch (err) {
    if (popupInFlight) {
      await popupInFlight;
      const retry = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: msalInstance.getAllAccounts()[0],
      });
      return retry.accessToken;
    }
    popupInFlight = msalInstance.acquireTokenPopup(loginRequest);
    try {
      const response = await popupInFlight;
      return response.accessToken;
    } finally {
      popupInFlight = null;
    }
  }
}

async function graphFetch(path, token) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API error ${res.status}: ${body}`);
  }
  return res.json();
}

async function resolveShareLink(token) {
  const encoded = encodeShareUrl(SHARE_URL);
  const data = await graphFetch(`/shares/${encoded}/driveItem`, token);
  return { driveId: data.parentReference.driveId, itemId: data.id };
}

async function getSheetValues(driveId, itemId, sheetName, token) {
  const path = `/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange(valuesOnly=true)`;
  const data = await graphFetch(path, token);
  return data.values || [];
}

// ── MAIN EXPORT: fetch and parse all sheets we need ──────────────────────
export async function fetchLivePilotData() {
  const token = await getAccessToken();
  const { driveId, itemId } = await resolveShareLink(token);

  const [onboardingRows, paymentsRows] = await Promise.all([
    getSheetValues(driveId, itemId, "Accounts Trans - Onboarding(DD)", token),
    getSheetValues(driveId, itemId, "Payments Transaction", token),
  ]);

  // Parse Onboarding(DD)
  // Columns: 0 TxnID, 1 Type, 2 Account, 3 Owner, 4 Destination, 5 Authorize,
  //          6 Submission, 7 DLT Close Time, 8 Created at, 9 Sequence, 10 User name, 11 Agent name
  const onboardingEvents = [];
  onboardingRows.slice(4).forEach(row => {
    if (!row[0] || row[0] === "Transaction ID") return;
    if (row[6] !== "tesSUCCESS") return; // only successful submissions count
    const agent = row[11] ? String(row[11]).trim() : null;
    onboardingEvents.push({
      txn_id: String(row[0]).slice(0, 16),
      owner: row[3] ? String(row[3]) : "",
      destination: row[4] ? String(row[4]) : "",
      authorize: row[5] ? String(row[5]) : "",
      dlt_close: row[7] ? String(row[7]) : "",
      user_name: row[10] ? String(row[10]) : "Unknown",
      agent_name: agent && agent !== "#N/A" ? agent : null,
    });
  });

  // Parse Payments Transaction
  // Columns (offset by leading blank col): 1 TxnID, 2 Type, 3 Account, 4 Destination,
  //          5 Amount(raw), 6 Submission, 7 DLT Close Time, 8 Created at, 9 Sequence,
  //          10 Include flag (ignored per spec), 11 Amount(BWP), 12 Sender name, 13 Receiver name
  const payments = [];
  paymentsRows.slice(3).forEach(row => {
    if (!row[1] || row[1] === "Transaction ID") return;
    if (row[6] !== "tesSUCCESS") return; // only successful submissions count
    const amtRaw = row[11] !== undefined && row[11] !== null ? String(row[11]).trim() : "0";
    const amt = parseFloat(amtRaw.replace("BWP", "").trim()) || 0;
    const sender = row[12] ? String(row[12]).trim() : "";
    const receiver = row[13] ? String(row[13]).trim() : "";

    payments.push({
      txn_id: String(row[1]).slice(0, 16),
      txn_type: row[2] ? String(row[2]).trim() : "",
      account: row[3] ? String(row[3]) : "",
      destination: row[4] ? String(row[4]) : "",
      amount: Math.round(amt * 100) / 100,
      dlt_close: row[7] ? String(row[7]) : "",
      created_at: row[8] ? String(row[8]) : "",
      sender: sender && sender !== "#N/A" ? sender : "",
      receiver: receiver && receiver !== "#N/A" ? receiver : "",
    });
  });

  return { onboardingEvents, payments, fetchedAt: new Date().toISOString() };
}
