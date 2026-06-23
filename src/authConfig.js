import { PublicClientApplication } from "@azure/msal-browser";

// ── AZURE APP CONFIG ─────────────────────────────────────────────────────
// These values are NOT secret — they identify the app, not authenticate it.
export const msalConfig = {
  auth: {
    clientId: "682c009f-633d-4073-9f9e-d4bcd6cd54b0",
    authority: "https://login.microsoftonline.com/6e03983e-0e96-41f7-806b-bb21f023fb19",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Files.Read"],
};

export const msalInstance = new PublicClientApplication(msalConfig);
