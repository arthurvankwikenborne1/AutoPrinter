/**
 * auth.js — MSAL authenticatie voor Outlook PDF Print Agent
 * Gebruikt @azure/msal-browser voor OAuth 2.0 via Microsoft 365
 */

// ─── CONFIGURATIE ────────────────────────────────────────────────────────────
// TODO: vervang CLIENT_ID door de Application (client) ID uit je Azure App Registration
// Zie README.md stap 1 voor instructies.
const MSAL_CONFIG = {
  auth: {
    clientId: "JOUW_CLIENT_ID_HIER",           // ← Azure App Registration > Overview > Application ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://arthurvankwikenborne1.github.io/AutoPrinter/src/auth/auth-redirect.html",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

// Welke Graph API rechten we nodig hebben
const GRAPH_SCOPES = [
  "Mail.Read",
  "Mail.ReadWrite",
];

// ─── MSAL INSTANTIE ──────────────────────────────────────────────────────────
let msalInstance = null;

/**
 * Initialiseer MSAL — roep dit aan bij het laden van de add-in
 */
export async function initAuth() {
  if (!window.msal) {
    throw new Error("MSAL library niet geladen. Controleer of msal-browser is geïmporteerd.");
  }
  msalInstance = new window.msal.PublicClientApplication(MSAL_CONFIG);
  await msalInstance.initialize();

  // Verwerk eventuele redirect-response (na login)
  await msalInstance.handleRedirectPromise();

  console.log("[Auth] MSAL geïnitialiseerd");
  return msalInstance;
}

/**
 * Haal een geldig access token op voor de Graph API.
 * Probeert eerst silent (geen popup), anders via popup.
 * @returns {Promise<string>} Access token
 */
export async function getAccessToken() {
  if (!msalInstance) {
    throw new Error("Auth niet geïnitialiseerd. Roep initAuth() eerst aan.");
  }

  const accounts = msalInstance.getAllAccounts();

  // Silent flow: token ophalen zonder gebruikersinteractie (gebruikt cached token)
  if (accounts.length > 0) {
    try {
      const silentRequest = {
        scopes: GRAPH_SCOPES,
        account: accounts[0],
      };
      const response = await msalInstance.acquireTokenSilent(silentRequest);
      console.log("[Auth] Token verkregen via silent flow");
      return response.accessToken;
    } catch (silentError) {
      console.warn("[Auth] Silent flow mislukt, popup starten:", silentError);
    }
  }

  // Popup flow: gebruiker moet inloggen of toestemming geven
  try {
    const popupRequest = { scopes: GRAPH_SCOPES };
    const response = await msalInstance.acquireTokenPopup(popupRequest);
    console.log("[Auth] Token verkregen via popup");
    return response.accessToken;
  } catch (popupError) {
    console.error("[Auth] Authenticatie mislukt:", popupError);
    throw new Error("Authenticatie mislukt. Probeer opnieuw.");
  }
}

/**
 * Geeft de naam en het e-mailadres van de ingelogde gebruiker terug.
 * @returns {{ name: string, email: string } | null}
 */
export function getCurrentUser() {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return null;
  return {
    name: accounts[0].name || accounts[0].username,
    email: accounts[0].username,
  };
}

/**
 * Uitloggen en sessie wissen
 */
export async function signOut() {
  if (!msalInstance) return;
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await msalInstance.logoutPopup({ account: accounts[0] });
  }
}