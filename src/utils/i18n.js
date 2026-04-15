/**
 * i18n.js — Vertalingen NL-BE / EN (Fase 3: volledig)
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
    printing: "Afdrukken...",
    printSuccess: (n) => `✅ ${n} bestand(en) afgedrukt.`,
    printPartial: (s, f) => `✅ ${s} afgedrukt · ⚠️ ${f} mislukt`,
    printFailed: (n) => `⚠️ ${n} bestand(en) mislukt.`,
    errorNetwork: "Geen verbinding. Controleer uw internetverbinding.",
    errorAuth: "Authenticatie mislukt. Meld u opnieuw aan.",
    errorGeneral: "Er is een fout opgetreden. Probeer opnieuw.",
    errorDownload: (name) => `Download mislukt: ${name}`,
    errorPrint: (name) => `Afdrukken mislukt: ${name}`,
    errorMarkRead: "Mail markeren als gelezen mislukt.",
    errorTimeout: "De bewerking duurde te lang. Probeer opnieuw.",
    history: "Historiek",
    historyEmpty: "Nog geen afdrukken geregistreerd.",
    historyBack: "← Terug",
    historyTitle: "Afdrukhistoriek",
    historyPrintedAt: "Afgedrukt op",
    settings: "Instellingen",
    language: "Taal",
    loggedInAs: "Ingelogd als",
    signOut: "Afmelden",
    cancel: "Annuleren",
    retry: "Opnieuw proberen",
    scanDesc: "Klik op de knop om uw ongelezen mails te scannen op PDF en Word bijlagen.",
    largeFileWarning: (name) => `⚠️ "${name}" is groter dan 10 MB. Dit kan lang duren.`,
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
    printing: "Printing...",
    printSuccess: (n) => `✅ ${n} file(s) printed.`,
    printPartial: (s, f) => `✅ ${s} printed · ⚠️ ${f} failed`,
    printFailed: (n) => `⚠️ ${n} file(s) failed.`,
    errorNetwork: "No connection. Please check your internet connection.",
    errorAuth: "Authentication failed. Please sign in again.",
    errorGeneral: "An error occurred. Please try again.",
    errorDownload: (name) => `Download failed: ${name}`,
    errorPrint: (name) => `Print failed: ${name}`,
    errorMarkRead: "Failed to mark email as read.",
    errorTimeout: "The operation took too long. Please try again.",
    history: "History",
    historyEmpty: "No print jobs recorded yet.",
    historyBack: "← Back",
    historyTitle: "Print history",
    historyPrintedAt: "Printed at",
    settings: "Settings",
    language: "Language",
    loggedInAs: "Logged in as",
    signOut: "Sign out",
    cancel: "Cancel",
    retry: "Try again",
    scanDesc: "Click the button to scan your unread emails for PDF and Word attachments.",
    largeFileWarning: (name) => `⚠️ "${name}" is larger than 10 MB. This may take a while.`,
  },
};

let currentLang = "nl";

export function setLanguage(lang) {
  currentLang = lang === "en" ? "en" : "nl";
  try { localStorage.setItem("autoprinter_lang", currentLang); } catch(e) {}
}

export function initLanguage() {
  try {
    const saved = localStorage.getItem("autoprinter_lang");
    if (saved === "en" || saved === "nl") currentLang = saved;
  } catch(e) {}
}

export function t(key, ...args) {
  const val = translations[currentLang][key];
  if (typeof val === "function") return val(...args);
  return val ?? key;
}

export function getCurrentLanguage() {
  return currentLang;
}