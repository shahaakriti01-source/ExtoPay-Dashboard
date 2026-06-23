// ─── ONBOARDING IDENTITY & DEDUPLICATION LOGIC ──────────────────────────────
import { EXTO_BACKEND_ID } from "./transactionLogic";

/**
 * Given raw onboarding rows (from Accounts Trans - Onboarding(DD), already filtered
 * to Submission === "tesSUCCESS"), dedupes by unique Owner ID and classifies each
 * unique onboarding event as either a merchant-onboarding-by-Exto or a
 * consumer-onboarding-by-merchant.
 *
 * Row shape expected: { owner, destination, authorize, dlt_close, user_name, agent_name }
 *   - owner: stable ID; duplicate owners (online+offline pair) = same event
 *   - authorize: who performed the onboarding (Exto backend ID, or a merchant's account ID)
 *   - destination: who got onboarded
 *   - agent_name: the resolved name of whoever is in the `authorize` column (if not Exto)
 *   - user_name: the resolved name of whoever is in the `destination` column
 */
export function dedupeOnboardingEvents(rawRows) {
  const seenOwners = new Set();
  const events = [];

  rawRows.forEach(row => {
    const ownerKey = row.owner || row.destination; // fallback if owner missing
    if (seenOwners.has(ownerKey)) return;
    seenOwners.add(ownerKey);

    const isMerchantOnboarding = row.authorize === EXTO_BACKEND_ID;

    events.push({
      owner: ownerKey,
      destination: row.destination,
      authorize: row.authorize,
      dlt_close: row.dlt_close,
      user_name: row.user_name,
      agent_name: row.agent_name, // null/undefined if Exto onboarded them directly
      isMerchantOnboarding,
    });
  });

  return events;
}

/** All unique merchants — either onboarded directly by Exto, or who have onboarded a consumer. */
export function getAllMerchantNames(onboardingEvents, payments) {
  const names = new Set();
  onboardingEvents.forEach(e => {
    if (e.isMerchantOnboarding) names.add(e.user_name); // they ARE the merchant
    if (e.agent_name) names.add(e.agent_name); // they onboarded someone, so they're a merchant
  });
  payments.forEach(p => {
    if (p.sender && p.sender !== "Exto Backend ID") names.add(p.sender);
    if (p.receiver && p.receiver !== "Exto Backend ID") names.add(p.receiver);
  });
  return [...names].sort();
}

/** Consumers onboarded by a specific merchant (Owner-deduped). */
export function getConsumersForMerchant(merchantName, onboardingEvents) {
  return onboardingEvents.filter(e => !e.isMerchantOnboarding && e.agent_name === merchantName);
}

/** Whether a merchant was successfully onboarded at all (Owner-deduped, by Exto). */
export function isMerchantOnboarded(merchantName, onboardingEvents) {
  return onboardingEvents.some(e => e.isMerchantOnboarding && e.user_name === merchantName);
}
