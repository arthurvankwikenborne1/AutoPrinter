/**
 * taskpane.js — Hoofdcontroller voor de PDF Print Agent taakvenster
 * Fase 2: Afdrukken, gelezen markeren, historiek
 */

import { initAuth, getAccessToken, getCurrentUser, signOut } from "../auth/auth.js";
import { scanAllFoldersForAttachments, downloadAttachment, markMessageAsRead } from "../api/graphApi.js";
import { t, setLanguage, getCurrentLanguage } from "../utils/i18n.js";

// ─── STATE ───────────────────────────────────────────────────────────────────
let accessToken = null;
let attachments = [];
let selectedIds = new Set();
let scanAborted = false;
let isPrinting = false;

// ─── OFFICE.JS INITIALISATIE ─────────────────────────────────────────────────
Office.onReady(async () => {
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
  const states = ["Idle", "Loading", "Loaded", "Empty", "Error", "Printing"];
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
    if (!navigator.onLine) {
      showError(t("errorNetwork"));
    } else if (err.message?.toLowerCase().includes("auth")) {
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

  isPrinting = true;
  updateToolbar();

  const printStatus = document.getElementById("printStatus");
  const printFrame = document.getElementById("printFrame");
  showState("printing");

  const results = { success: [], failed: [] };

  for (const att of selected) {
    try {
      // Status tonen
      if (printStatus) {
        printStatus.textContent = `${t("printing")} ${att.fileName}... (${results.success.length + results.failed.length + 1}/${selected.length})`;
      }

      // Bijlage downloaden
      const blob = await downloadAttachment(att.messageId, att.attachmentId, accessToken);
      const url = URL.createObjectURL(blob);

      // Afdrukken via verborgen iframe
      await printBlob(url, att.fileType, printFrame);

      // Mail markeren als gelezen
      await markMessageAsRead(att.messageId, accessToken);

      results.success.push(att);

      // Historiek opslaan
      saveToHistory(att);

      // URL vrijgeven
      URL.revokeObjectURL(url);

    } catch (err) {
      console.error(`[Print] Fout bij ${att.fileName}:`, err);
      results.failed.push({ att, error: err.message });
    }
  }

  isPrinting = false;

  // Toon resultaat
  showPrintResult(results);

  // Verwijder afgedrukte items uit de lijst
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

function printBlob(url, fileType, printFrame) {
  return new Promise((resolve) => {
    printFrame.onload = () => {
      try {
        printFrame.contentWindow.focus();
        printFrame.contentWindow.print();
      } catch (e) {
        console.warn("[Print] Print dialoog kon niet worden geopend:", e);
      }
      setTimeout(resolve, 1500);
    };
    printFrame.src = url;
  });
}

function showPrintResult(results) {
  const successCount = results.success.length;
  const failCount = results.failed.length;

  let msg = "";
  if (successCount > 0) msg += `✅ ${successCount} bestand(en) afgedrukt. `;
  if (failCount > 0) {
    msg += `⚠️ ${failCount} mislukt: `;
    msg += results.failed.map((f) => f.att.fileName).join(", ");
  }

  const el = document.getElementById("printResult");
  if (el) {
    el.textContent = msg;
    el.style.display = "block";
    setTimeout(() => { el.style.display = "none"; }, 6000);
  }
}

// ─── HISTORIEK ────────────────────────────────────────────────────────────────
function saveToHistory(att) {
  try {
    const key = "printHistory";
    const existing = JSON.parse(localStorage.getItem(key) || "[]");
    existing.unshift({
      fileName: att.fileName,
      sender: att.sender.name,
      subject: att.subject,
      printedAt: new Date().toISOString(),
    });
    // Max 200 records
    localStorage.setItem(key, JSON.stringify(existing.slice(0, 200)));
  } catch (e) {
    console.warn("[History] Kon niet opslaan:", e);
  }
}

function renderHistory() {
  const historyEl = document.getElementById("historyList");
  if (!historyEl) return;

  try {
    const history = JSON.parse(localStorage.getItem("printHistory") || "[]");
    if (history.length === 0) {
      historyEl.innerHTML = `<p class="empty-message">${t("historyEmpty")}</p>`;
      return;
    }

    historyEl.innerHTML = history.map((item) => `
      <div class="history-item">
        <div class="history-item__name">${escHtml(item.fileName)}</div>
        <div class="history-item__meta">${escHtml(item.sender)} · ${formatDate(item.printedAt)}</div>
      </div>
    `).join("");
  } catch (e) {
    historyEl.innerHTML = `<p class="empty-message">${t("historyEmpty")}</p>`;
  }
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
  document.getElementById("historyTab")?.addEventListener("click", () => {
    renderHistory();
    document.getElementById("stateHistory")?.classList.remove("state--hidden");
    document.getElementById("stateLoaded")?.classList.add("state--hidden");
  });
  document.getElementById("backFromHistory")?.addEventListener("click", () => {
    document.getElementById("stateHistory")?.classList.add("state--hidden");
    showState(attachments.length > 0 ? "loaded" : "idle");
  });
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
    day: "2-digit", month: "short", year: "numeric",
  });
}

function formatSize(bytes) {
  if (!bytes) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}