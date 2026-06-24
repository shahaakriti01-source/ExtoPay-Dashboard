// ─── ONBOARDING IDENTITY & DEDUPLICATION LOGIC ──────────────────────────────
import { EXTO_BACKEND_ID, EXTO_BACKEND_ONBOARD_ID } from "./transactionLogic";

// Field coordinators / Exto Pay representatives on the ground. Whenever the
// Authorize column resolves to one of these people (by NAME, not by a fixed ID
// list — since each person can hold multiple account IDs across card + phone
// wallet activations, and new coordinators may be added later), the person in
// the Destination column is being onboarded as a MERCHANT, not a consumer.
//
// To add a new coordinator in the future: just add their exact resolved name
// (as it appears in the "Merchant/User name" column) to this list.
export const FIELD_COORDINATOR_NAMES = [
  "Sandeep Taterway",
  "Segomotso Sadie M Mckenzie",
  "Cindy Sean Sibanda",
  "Brian Ketumile K Moeng",
];

function normalizeName(name) {
  return (name || "").trim().toUpperCase().replace(/\s+/g, " ");
}

const NORMALIZED_COORDINATOR_NAMES = new Set(FIELD_COORDINATOR_NAMES.map(normalizeName));

function isFieldCoordinatorOrExto(authorizeId, authorizeResolvedName) {
  if (authorizeId === EXTO_BACKEND_ID || authorizeId === EXTO_BACKEND_ONBOARD_ID) return true;
  return NORMALIZED_COORDINATOR_NAMES.has(normalizeName(authorizeResolvedName));
}

// Display name used for Exto's own account in payment sender/receiver names.
const EXTO_DISPLAY_NAMES = new Set(["Exto Backend ID", "Exto Backend ID - Pay", "Exto Backend ID - Onboard"]);

/** All "special" account names: Exto (any of its display variants) + the 4 field coordinators. */
export const SPECIAL_ACCOUNT_NAMES = [...FIELD_COORDINATOR_NAMES, ...EXTO_DISPLAY_NAMES];
const NORMALIZED_SPECIAL_NAMES = new Set(SPECIAL_ACCOUNT_NAMES.map(normalizeName));

/** Whether a given (already-resolved) person/account name is Exto or a field coordinator. */
export function isSpecialAccountName(name) {
  return NORMALIZED_SPECIAL_NAMES.has(normalizeName(name));
}

/**
 * Given raw onboarding rows (from Accounts Trans - Onboarding(DD), already filtered
 * to Submission === "tesSUCCESS"), dedupes by unique Owner ID and classifies each
 * unique onboarding event as either a merchant-onboarding (by Exto or a field
 * coordinator) or a consumer-onboarding (by an actual merchant).
 *
 * Row shape expected: { owner, destination, authorize, dlt_close, user_name, agent_name }
 *   - owner: stable ID; duplicate owners (online+offline pair) = same event
 *   - authorize: the Account ID of who performed the onboarding
 *   - destination: who got onboarded
 *   - agent_name: the RESOLVED NAME of whoever is in the `authorize` column
 *   - user_name: the resolved name of whoever is in the `destination` column
 *
 * IMPORTANT: `agent_name` must be the resolved name of the Authorize column
 * (via the Mapping Person to ID sheet), not left null just because the ID
 * isn't the single old Exto onboarding ID — field coordinators have real,
 * resolvable names and we classify by matching those names.
 */
export function dedupeOnboardingEvents(rawRows) {
  const seenOwners = new Set();
  const events = [];

  rawRows.forEach(row => {
    const ownerKey = row.owner || row.destination; // fallback if owner missing
    if (seenOwners.has(ownerKey)) return;
    seenOwners.add(ownerKey);

    const isMerchantOnboarding = isFieldCoordinatorOrExto(row.authorize, row.agent_name);

    events.push({
      owner: ownerKey,
      destination: row.destination,
      authorize: row.authorize,
      dlt_close: row.dlt_close,
      user_name: row.user_name,
      // Keep the coordinator's name even when they're acting as an onboarder
      // (useful for the Merchants tab / debugging), but isMerchantOnboarding
      // is the flag that actually drives merchant-vs-consumer classification.
      agent_name: isMerchantOnboarding ? null : row.agent_name,
      isMerchantOnboarding,
    });
  });

  return events;
}

/** All unique merchants — either onboarded directly by Exto/a field coordinator, or who have onboarded a consumer. */
export function getAllMerchantNames(onboardingEvents, payments) {
  const names = new Set();
  onboardingEvents.forEach(e => {
    if (e.isMerchantOnboarding) names.add(e.user_name); // they ARE the merchant
    if (e.agent_name) names.add(e.agent_name); // they onboarded a consumer, so they're a merchant
  });
  payments.forEach(p => {
    if (p.sender && p.sender !== "Exto Backend ID") names.add(p.sender);
    if (p.receiver && p.receiver !== "Exto Backend ID") names.add(p.receiver);
  });
  // Field coordinators themselves are not merchants — remove them if they
  // ended up in the set via a stray payment record.
  FIELD_COORDINATOR_NAMES.forEach(n => names.delete(n));
  return [...names].sort();
}

/** Consumers onboarded by a specific merchant (Owner-deduped). */
export function getConsumersForMerchant(merchantName, onboardingEvents) {
  return onboardingEvents.filter(e => !e.isMerchantOnboarding && e.agent_name === merchantName);
}

/** Whether a merchant was successfully onboarded at all (Owner-deduped, by Exto or a field coordinator). */
export function isMerchantOnboarded(merchantName, onboardingEvents) {
  return onboardingEvents.some(e => e.isMerchantOnboarding && e.user_name === merchantName);
}
