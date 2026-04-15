/**
 * taskpane.js — Hoofdcontroller voor de PDF Print Agent taakvenster
 * Fase 1: Auth + scan + lijstweergave
 */

import { initAuth, getAccessToken, getCurrentUser, signOut } from "../auth/auth.js";
import { scanAllFoldersForAttachments } from "../api/graphApi.js";
import { t, setLanguage, getCurrentLanguage } from "../utils/i18n.js";

// ─── STATE ───────────────────────────────────────────────────────────────────
let accessToken = null;
let attachments = [];        // Alle gevonden bijlagen
let selectedIds = new Set(); // Geselecteerde bijlage-ID's
let scanAborted = false;

// ─── OFFICE.JS INITIALISATIE ─────────────────────────────────────────────────
Office.onReady(async () => {
  console.log("[App] Office.js geladen");
  applyTranslations();
  bindEventListeners();

  try {
    await initAuth();
    updateUserInfo();
    showState("idle");
  } catch (err) {
    showError(t("errorAuth"));
  }
});

// ─── UI STATES ────────────────────────────────────────────────────────────────
function showState(state) {
  const states = ["Idle", "Loading", "Loaded", "Empty", "Error"];
  states.forEach((s) => {
    document.getElementById(`state${s}`)?.classList.toggle("state--hidden", s.toLowerCase() !== state);
  });
}

function showError(message) {
  document.getElementById("errorMessage").textContent = message;
  showState("error");
}

// ─── SCAN ─────────────────────────────────────────────────────────────────────
async function startScan() {
  scanAborted = false;
  showState("loading");
  selectedIds.clear();

  try {
    // Token ophalen (silent of popup)
    accessToken = await getAccessToken();
    updateUserInfo();

    // Scan alle mappen
    attachments = await scanAllFoldersForAttachments(
      accessToken,
      (statusText) => {
        if (!scanAborted) {
          document.getElementById("loadingStatus").textContent = statusText;
        }
      }
    );

    if (attachments.length === 0) {
      showState("empty");
    } else {
      renderAttachmentList();
      showState("loaded");
    }

  } catch (err) {
    console.error("[Scan] Fout:", err);
    if (err.message?.includes("netwerk") || err.message?.includes("network") || !navigator.onLine) {
      showError(t("errorNetwork"));
    } else if (err.message?.includes("auth") || err.message?.includes("Auth")) {
      showError(t("errorAuth"));
    } else {
      showError(t("errorGeneral"));
    }
  }
}

// ─── RENDER LIJST ─────────────────────────────────────────────────────────────
function renderAttachmentList() {
  const listEl = document.getElementById("attachmentList");
  listEl.innerHTML = "";

  // Groepeer per afzender
  const groups = {};
  for (const att of attachments) {
    const key = att.sender.email || att.sender.name;
    if (!groups[key]) groups[key] = { sender: att.sender, items: [] };
    groups[key].items.push(att);
  }

  for (const [, group] of Object.entries(groups)) {
    const groupEl = document.createElement("div");
    groupEl.className = "sender-group";

    // Groepsheader
    const header = document.createElement("div");
    header.className = "sender-group__header";
    header.innerHTML = `
      <span>👤 ${escHtml(group.sender.name)}</span>
      <span class="sender-group__email">${escHtml(group.sender.email)}</span>
    `;
    groupEl.appendChild(header);

    // Bijlage items
    for (const att of group.items) {
      groupEl.appendChild(createAttachmentItem(att));
    }

    listEl.appendChild(groupEl);
  }

  updateToolbar();
}

function createAttachmentItem(att) {
  const itemEl = document.createElement("div");
  itemEl.className = "attachment-item";
  itemEl.dataset.id = att.attachmentId;

  const icon = att.fileType === "pdf" ? "📄" : "📝";
  const dateStr = formatDate(att.receivedDate);
  const sizeStr = formatSize(att.fileSize);
  const isSelected = selectedIds.has(att.attachmentId);

  if (isSelected) itemEl.classList.add("attachment-item--selected");

  itemEl.innerHTML = `
    <input type="checkbox" class="attachment-item__checkbox" ${isSelected ? "checked" : ""} />
    <div class="attachment-item__icon">${icon}</div>
    <div class="attachment-item__info">
      <div class="attachment-item__name" title="${escHtml(att.fileName)}">${escHtml(att.fileName)}</div>
      <div class="attachment-item__meta">${escHtml(att.subject)} · ${dateStr}</div>
      <div class="attachment-item__badges">
        <span class="badge badge--${att.fileType}">${att.fileType.toUpperCase()}</span>
        ${att.isLargeFile ? `<span class="badge badge--large">⚠ ${t("largeFile")}</span>` : ""}
      </div>
    </div>
    <div class="attachment-item__size">${sizeStr}</div>
  `;

  // Klikken op rij of checkbox togglet selectie
  itemEl.addEventListener("click", (e) => {
    toggleSelection(att.attachmentId, itemEl, itemEl.querySelector("input"));
  });

  return itemEl;
}

// ─── SELECTIE ─────────────────────────────────────────────────────────────────
function toggleSelection(id, itemEl, checkbox) {
  if (selectedIds.has(id)) {
    selectedIds.delete(id);
    itemEl.classList.remove("attachment-item--selected");
    if (checkbox) checkbox.checked = false;
  } else {
    selectedIds.add(id);
    itemEl.classList.add("attachment-item--selected");
    if (checkbox) checkbox.checked = true;
  }
  updateToolbar();
}

function updateToolbar() {
  const count = selectedIds.size;
  const total = attachments.length;

  // Teller
  document.getElementById("selectionCount").textContent =
    count > 0 ? t("selected", count) : "";

  // Afdrukknop
  const printBtn = document.getElementById("printBtn");
  printBtn.disabled = count === 0;
  document.getElementById("printBtnLabel").textContent =
    count > 0 ? t("printButtonWithCount", count) : t("printButton");

  // Selecteer-alles checkbox
  const selectAllCb = document.getElementById("selectAllCheckbox");
  selectAllCb.checked = count === total && total > 0;
  selectAllCb.indeterminate = count > 0 && count < total;
}

function selectAll(checked) {
  attachments.forEach((att) => {
    if (checked) selectedIds.add(att.attachmentId);
    else selectedIds.delete(att.attachmentId);
  });

  // Update alle checkboxes visueel
  document.querySelectorAll(".attachment-item").forEach((itemEl) => {
    const id = itemEl.dataset.id;
    const checkbox = itemEl.querySelector("input");
    if (checked) {
      itemEl.classList.add("attachment-item--selected");
      if (checkbox) checkbox.checked = true;
    } else {
      itemEl.classList.remove("attachment-item--selected");
      if (checkbox) checkbox.checked = false;
    }
  });

  updateToolbar();
}

// ─── AFDRUKKEN (Fase 2 placeholder) ──────────────────────────────────────────
function handlePrint() {
  const selected = attachments.filter((a) => selectedIds.has(a.attachmentId));
  // TODO Fase 2: bijlagen downloaden en window.print() aanroepen
  alert(`Afdrukken van ${selected.length} bestand(en) — wordt geïmplementeerd in Fase 2.\n\nGeselecteerd:\n${selected.map(a => a.fileName).join("\n")}`);
}

// ─── TAAL ─────────────────────────────────────────────────────────────────────
function switchLanguage(lang) {
  setLanguage(lang);
  document.getElementById("langNL").classList.toggle("lang-btn--active", lang === "nl");
  document.getElementById("langEN").classList.toggle("lang-btn--active", lang === "en");
  applyTranslations();
}

function applyTranslations() {
  document.getElementById("appSubtitle").textContent   = t("appSubtitle");
  document.getElementById("scanBtnLabel").textContent  = t("scanButton");
  document.getElementById("selectAllLabel").textContent = t("selectAll");
  document.getElementById("printBtnLabel").textContent = t("printButton");
  document.getElementById("rescanBtnLabel").textContent = t("rescan");
  document.getElementById("rescanEmptyLabel").textContent = t("rescan");
  const emptyMsg = document.getElementById("emptyMessage");
  if (emptyMsg) emptyMsg.textContent = t("noResults");
}

function updateUserInfo() {
  const user = getCurrentUser();
  if (user) {
    document.getElementById("userName").textContent = user.name || user.email;
    document.getElementById("userInfo").style.display = "flex";
  }
}

// ─── EVENT LISTENERS ──────────────────────────────────────────────────────────
function bindEventListeners() {
  document.getElementById("scanBtn").addEventListener("click", startScan);
  document.getElementById("cancelBtn").addEventListener("click", () => {
    scanAborted = true;
    showState("idle");
  });
  document.getElementById("retryBtn").addEventListener("click", startScan);
  document.getElementById("rescanBtn").addEventListener("click", startScan);
  document.getElementById("rescanBtnEmpty").addEventListener("click", startScan);
  document.getElementById("printBtn").addEventListener("click", handlePrint);
  document.getElementById("selectAllCheckbox").addEventListener("change", (e) => selectAll(e.target.checked));
  document.getElementById("clearBtn").addEventListener("click", () => selectAll(false));
  document.getElementById("signOutBtn").addEventListener("click", async () => {
    await signOut();
    document.getElementById("userInfo").style.display = "none";
    showState("idle");
  });
  document.getElementById("langNL").addEventListener("click", () => switchLanguage("nl"));
  document.getElementById("langEN").addEventListener("click", () => switchLanguage("en"));
}

// ─── HULPFUNCTIES ─────────────────────────────────────────────────────────────
function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatDate(isoString) {
  const date = new Date(isoString);
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  if (date.toDateString() === today.toDateString()) return t("today");
  if (date.toDateString() === yesterday.toDateString()) return t("yesterday");
  return date.toLocaleDateString(getCurrentLanguage() === "nl" ? "nl-BE" : "en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function formatSize(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}