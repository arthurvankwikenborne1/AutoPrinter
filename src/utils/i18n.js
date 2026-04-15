/**
 * i18n.js — Vertalingen NL-BE / EN
 */

export const translations = {
  nl: {
    appTitle: "PDF Print Agent",
    appSubtitle: "Druk bijlagen af vanuit ongelezen mails",
    scanButton: "Scan ongelezen mails",
    scanning: "Scannen...",
    statusFetching: "Ongelezen mails ophalen...",
    statusAnalysing: "Mails analyseren op bijlagen...",
    statusIdentifying: "PDF- en Word-bijlagen identificeren...",
    statusPreparing: "Lijst voorbereiden...",
    noResults: "Geen PDF- of Word-bijlagen gevonden in ongelezen mails.",
    noUnread: "Geen ongelezen mails gevonden.",
    rescan: "Opnieuw scannen",
    selectAll: "Alles selecteren",
    clearSelection: "Wissen",
    printButton: "Afdrukken",
    printButtonWithCount: (n) => `Afdrukken (${n})`,
    selected: (n) => `${n} geselecteerd`,
    largeFile: "Groot bestand",
    today: "Vandaag",
    yesterday: "Gisteren",
    errorNetwork: "Geen verbinding. Controleer uw internetverbinding.",
    errorAuth: "Authenticatie mislukt. Probeer opnieuw.",
    errorGeneral: "Er is een fout opgetreden. Probeer opnieuw.",
    history: "Historiek",
    historyEmpty: "Nog geen afdrukken geregistreerd.",
    settings: "Instellingen",
    language: "Taal",
    loggedInAs: "Ingelogd als",
    signOut: "Afmelden",
  },
  en: {
    appTitle: "PDF Print Agent",
    appSubtitle: "Print attachments from unread emails",
    scanButton: "Scan unread emails",
    scanning: "Scanning...",
    statusFetching: "Fetching unread emails...",
    statusAnalysing: "Analysing emails for attachments...",
    statusIdentifying: "Identifying PDF and Word attachments...",
    statusPreparing: "Preparing list...",
    noResults: "No PDF or Word attachments found in unread emails.",
    noUnread: "No unread emails found.",
    rescan: "Scan again",
    selectAll: "Select all",
    clearSelection: "Clear",
    printButton: "Print",
    printButtonWithCount: (n) => `Print (${n})`,
    selected: (n) => `${n} selected`,
    largeFile: "Large file",
    today: "Today",
    yesterday: "Yesterday",
    errorNetwork: "No connection. Please check your internet connection.",
    errorAuth: "Authentication failed. Please try again.",
    errorGeneral: "An error occurred. Please try again.",
    history: "History",
    historyEmpty: "No print jobs recorded yet.",
    settings: "Settings",
    language: "Language",
    loggedInAs: "Logged in as",
    signOut: "Sign out",
  },
};

let currentLang = "nl";

export function setLanguage(lang) {
  currentLang = lang === "en" ? "en" : "nl";
}

export function t(key, ...args) {
  const val = translations[currentLang][key];
  if (typeof val === "function") return val(...args);
  return val ?? key;
}

export function getCurrentLanguage() {
  return currentLang;
}