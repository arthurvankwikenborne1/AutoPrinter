/**
 * taskpane.js — Fase 3: volledige foutafhandeling + NL/EN persistent
 */

import { initAuth, getAccessToken, getCurrentUser, signOut } from "../auth/auth.js";
import { scanAllFoldersForAttachments, downloadAttachment, markMessageAsRead } from "../api/graphApi.js";
import { t, setLanguage, getCurrentLanguage, initLanguage } from "../utils/i18n.js";

// ─── STATE ───────────────────────────────────────────────────────────────────
let accessToken = null;
let attachments = [];
let selectedIds = new Set();
let scanAborted = false;
let isPrinting = false;

// ─── OFFICE.JS INITIALISATIE ─────────────────────────────────────────────────
Office.onReady(async () => {
  initLanguage();
  applyTranslations();
  updateLangButtons();
  bindEventListeners();

  try {
    await initAuth();
    updateUserInfo();
    showState("idle");
  } catch (err) {
    showError(t("errorAuth"), true);
  }
});

// ─── UI STATES ────────────────────────────────────────────────────────────────
function showState(state) {
  const states = ["Idle", "Loading", "Loaded", "Empty", "Error", "Printing", "History"];
  states.forEach((s) => {
    document.getElementById(`state${s}`)?.classList.toggle("state--hidden", s.toLowerCase() !== state);
  });
}

function showError(message, canRetry = true) {
  document.getElementById("errorMessage").textContent = message;
  const retryBtn = document.getElementById("retryBtn");
  if (retryBtn) retryBtn.style.display = canRetry ? "inline-flex" : "none";
  showState("error");
}

function showBanner(message, type = "success") {
  const el = document.getElementById("printResult");
  if (!el) return;
  el.textContent = message;
  el.className = `print-result print-result--${type}`;
  el.style.display = "block";
  clearTimeout(el._timer);
  el._timer = setTimeout(() => { el.style.display = "none"; }, 6000);
}

// ─── SCAN ─────────────────────────────────────────────────────────────────────
async function startScan() {
  scanAborted = false;
  showState("loading");
  selectedIds.clear();

  try {
    accessToken = await getAccessToken();
    updateUserInfo();

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
    console.error("[Scan]", err);
    if (!navigator.onLine) {
      showError(t("errorNetwork"));
    } else if (
      err.message?.toLowerCase().includes("auth") ||
      err.message?.toLowerCase().includes("token") ||
      err.message?.includes("401")
    ) {
      showError(t("errorAuth"));
    } else if (err.message?.toLowerCase().includes("timeout")) {
      showError(t("errorTimeout"));
    } else {
      showError(t("errorGeneral"));
    }
  }
}

// ─── RENDER LIJST ─────────────────────────────────────────────────────────────
function renderAttachmentList() {
  const listEl = document.getElementById("attachmentList");
  listEl.innerHTML = "";

  const groups = {};
  for (const att of attachments) {
    const key = att.sender.email || att.sender.name;
    if (!groups[key]) groups[key] = { sender: att.sender, items: [] };
    groups[key].items.push(att);
  }

  for (const [, group] of Object.entries(groups)) {
    const groupEl = document.createElement("div");
    groupEl.className = "sender-group";

    const header = document.createElement("div");
    header.className = "sender-group__header";
    header.innerHTML = `
      <span>👤 ${escHtml(group.sender.name)}</span>
      <span class="sender-group__email">${escHtml(group.sender.email)}</span>
    `;
    groupEl.appendChild(header);

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
  const isSelected = selectedIds.has(att.attachmentId);
  if (isSelected) itemEl.classList.add("attachment-item--selected");

  itemEl.innerHTML = `
    <input type="checkbox" class="attachment-item__checkbox" ${isSelected ? "checked" : ""} />
    <div class="attachment-item__icon">${icon}</div>
    <div class="attachment-item__info">
      <div class="attachment-item__name" title="${escHtml(att.fileName)}">${escHtml(att.fileName)}</div>
      <div class="attachment-item__meta">${escHtml(att.subject)} · ${formatDate(att.receivedDate)}</div>
      <div class="attachment-item__badges">
        <span class="badge badge--${att.fileType}">${att.fileType.toUpperCase()}</span>
        ${att.isLargeFile ? `<span class="badge badge--large">⚠ ${t("largeFile")}</span>` : ""}
      </div>
    </div>
    <div class="attachment-item__size">${formatSize(att.fileSize)}</div>
  `;

  itemEl.addEventListener("click", () => {
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

  document.getElementById("selectionCount").textContent =
    count > 0 ? t("selected", count) : "";

  const printBtn = document.getElementById("printBtn");
  printBtn.disabled = count === 0 || isPrinting;
  document.getElementById("printBtnLabel").textContent =
    count > 0 ? t("printButtonWithCount", count) : t("printButton");

  const selectAllCb = document.getElementById("selectAllCheckbox");
  selectAllCb.checked = count === total && total > 0;
  selectAllCb.indeterminate = count > 0 && count < total;
}

function selectAll(checked) {
  attachments.forEach((att) => {
    if (checked) selectedIds.add(att.attachmentId);
    else selectedIds.delete(att.attachmentId);
  });

  document.querySelectorAll(".attachment-item").forEach((itemEl) => {
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

// ─── AFDRUKKEN ────────────────────────────────────────────────────────────────
async function handlePrint() {
  const selected = attachments.filter((a) => selectedIds.has(a.attachmentId));
  if (selected.length === 0) return;

  // Waarschuwing grote bestanden
  const largeFiles = selected.filter((a) => a.isLargeFile);
  if (largeFiles.length > 0) {
    const names = largeFiles.map((a) => a.fileName).join(", ");
    if (!confirm(t("largeFileWarning", names))) return;
  }

  isPrinting = true;
  updateToolbar();

  const printFrame = document.getElementById("printFrame");
  showState("printing");

  const results = { success: [], failed: [] };

  for (const att of selected) {
    const statusEl = document.getElementById("printStatus");
    if (statusEl) {
      statusEl.textContent = `${t("printing")} ${att.fileName} (${results.success.length + results.failed.length + 1}/${selected.length})`;
    }

    let blobUrl = null;

    try {
      // Download
      const blob = await withTimeout(
        downloadAttachment(att.messageId, att.attachmentId, accessToken),
        30000,
        t("errorDownload", att.fileName)
      );
      blobUrl = URL.createObjectURL(blob);

      // Print
      await printBlob(blobUrl, printFrame);

      // Gelezen markeren — fout hier stopt de print niet
      try {
        await markMessageAsRead(att.messageId, accessToken);
      } catch (markErr) {
        console.warn("[Print] Markeren mislukt:", markErr.message);
      }

      results.success.push(att);
      saveToHistory(att);

    } catch (err) {
      console.error(`[Print] Fout bij ${att.fileName}:`, err);
      results.failed.push({ att, error: err.message });
    } finally {
      if (blobUrl) URL.revokeObjectURL(blobUrl);
    }
  }

  isPrinting = false;

  // Banner tonen
  const s = results.success.length;
  const f = results.failed.length;
  if (f === 0) {
    showBanner(t("printSuccess", s), "success");
  } else if (s === 0) {
    showBanner(t("printFailed", f), "error");
  } else {
    showBanner(t("printPartial", s, f), "warning");
  }

  // Afgedrukte items verwijderen
  attachments = attachments.filter(
    (a) => !results.success.some((s) => s.attachmentId === a.attachmentId)
  );
  results.success.forEach((a) => selectedIds.delete(a.attachmentId));

  if (attachments.length === 0) {
    showState("empty");
  } else {
    renderAttachmentList();
    showState("loaded");
  }
}

function printBlob(url, printFrame) {
  return new Promise((resolve) => {
    printFrame.onload = () => {
      try {
        printFrame.contentWindow.focus();
        printFrame.contentWindow.print();
      } catch (e) {
        console.warn("[Print] Print dialoog fout:", e);
      }
      setTimeout(resolve, 1500);
    };
    printFrame.onerror = () => resolve();
    printFrame.src = url;
  });
}

function withTimeout(promise, ms, errorMsg) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(errorMsg || t("errorTimeout"))), ms);
    promise.then(
      (val) => { clearTimeout(timer); resolve(val); },
      (err) => { clearTimeout(timer); reject(err); }
    );
  });
}

// ─── HISTORIEK ────────────────────────────────────────────────────────────────
function saveToHistory(att) {
  try {
    const existing = JSON.parse(localStorage.getItem("printHistory") || "[]");
    existing.unshift({
      fileName: att.fileName,
      fileType: att.fileType,
      sender: att.sender.name,
      subject: att.subject,
      printedAt: new Date().toISOString(),
    });
    localStorage.setItem("printHistory", JSON.stringify(existing.slice(0, 200)));
  } catch (e) {
    console.warn("[History] Opslaan mislukt:", e);
  }
}

function renderHistory() {
  const historyEl = document.getElementById("historyList");
  if (!historyEl) return;

  try {
    const history = JSON.parse(localStorage.getItem("printHistory") || "[]");
    if (history.length === 0) {
      historyEl.innerHTML = `<div style="padding:16px 12px;color:var(--color-text-secondary,#605e5c);font-size:13px">${t("historyEmpty")}</div>`;
      return;
    }

    historyEl.innerHTML = history.map((item) => `
      <div class="history-item">
        <div class="history-item__icon">${item.fileType === "pdf" ? "📄" : "📝"}</div>
        <div class="history-item__info">
          <div class="history-item__name">${escHtml(item.fileName)}</div>
          <div class="history-item__meta">${escHtml(item.sender)} · ${formatDate(item.printedAt)}</div>
          <div class="history-item__subject">${escHtml(item.subject)}</div>
        </div>
        <div class="history-item__time">${formatTime(item.printedAt)}</div>
      </div>
    `).join("");
  } catch (e) {
    historyEl.innerHTML = `<div style="padding:16px 12px">${t("historyEmpty")}</div>`;
  }
}

// ─── TAAL ─────────────────────────────────────────────────────────────────────
function switchLanguage(lang) {
  setLanguage(lang);
  updateLangButtons();
  applyTranslations();
  if (document.getElementById("stateHistory")?.classList.contains("state--hidden") === false) {
    renderHistory();
  }
}

function updateLangButtons() {
  const lang = getCurrentLanguage();
  document.getElementById("langNL")?.classList.toggle("lang-btn--active", lang === "nl");
  document.getElementById("langEN")?.classList.toggle("lang-btn--active", lang === "en");
}

function applyTranslations() {
  const set = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  set("appSubtitle", t("appSubtitle"));
  set("scanBtnLabel", t("scanButton"));
  set("idleDesc", t("scanDesc"));
  set("selectAllLabel", t("selectAll"));
  set("printBtnLabel", t("printButton"));
  set("rescanBtnLabel", t("rescan"));
  set("rescanEmptyLabel", t("rescan"));
  set("emptyMessage", t("noResults"));
  set("historyTitle", t("historyTitle"));
  set("backFromHistoryLabel", t("historyBack"));
  set("retryBtn", t("retry"));
  set("cancelBtn", t("cancel"));
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
  const on = (id, ev, fn) => document.getElementById(id)?.addEventListener(ev, fn);

  on("scanBtn", "click", startScan);
  on("cancelBtn", "click", () => { scanAborted = true; showState("idle"); });
  on("retryBtn", "click", startScan);
  on("rescanBtn", "click", startScan);
  on("rescanBtnEmpty", "click", startScan);
  on("printBtn", "click", handlePrint);
  on("selectAllCheckbox", "change", (e) => selectAll(e.target.checked));
  on("clearBtn", "click", () => selectAll(false));
  on("signOutBtn", "click", async () => {
    await signOut();
    document.getElementById("userInfo").style.display = "none";
    accessToken = null;
    attachments = [];
    selectedIds.clear();
    showState("idle");
  });
  on("langNL", "click", () => switchLanguage("nl"));
  on("langEN", "click", () => switchLanguage("en"));
  on("historyTab", "click", () => {
    renderHistory();
    showState("history");
  });
  on("backFromHistory", "click", () => {
    showState(attachments.length > 0 ? "loaded" : "idle");
  });
}

// ─── HULPFUNCTIES ─────────────────────────────────────────────────────────────
function escHtml(str) {
  return String(str ?? "")
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
    day: "2-digit", month: "short", year: "numeric",
  });
}

function formatTime(isoString) {
  return new Date(isoString).toLocaleTimeString(
    getCurrentLanguage() === "nl" ? "nl-BE" : "en-GB",
    { hour: "2-digit", minute: "2-digit" }
  );
}

function formatSize(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}