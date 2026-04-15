/**
 * graphApi.js — Microsoft Graph API aanroepen
 * Haalt ongelezen mails en bijlagen op uit alle mappen
 */

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Hulpfunctie: voer een Graph API GET-aanroep uit
 */
async function graphGet(url, accessToken) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(
      `Graph API fout ${response.status}: ${error?.error?.message || response.statusText}`
    );
  }

  return response.json();
}

/**
 * Haal alle mailmappen op uit de mailbox van de gebruiker
 * @param {string} accessToken
 * @returns {Promise<Array<{id: string, displayName: string}>>}
 */
export async function getAllMailFolders(accessToken) {
  let folders = [];
  let url = `${GRAPH_BASE}/me/mailFolders?$top=100`;

  // Paginering: haal alle mappen op (kan meerdere pagina's zijn)
  while (url) {
    const data = await graphGet(url, accessToken);
    folders = folders.concat(data.value || []);
    url = data["@odata.nextLink"] || null;
  }

  console.log(`[Graph] ${folders.length} mailmappen gevonden`);
  return folders;
}

/**
 * Haal ongelezen mails MET bijlagen op uit één specifieke map
 * @param {string} folderId
 * @param {string} accessToken
 * @returns {Promise<Array>} Berichten met bijlagen
 */
export async function getUnreadMessagesWithAttachments(folderId, accessToken) {
  const url =
    `${GRAPH_BASE}/me/mailFolders/${folderId}/messages` +
    `?$filter=isRead eq false and hasAttachments eq true` +
    `&$select=id,subject,from,receivedDateTime,hasAttachments` +
    `&$expand=attachments($select=id,name,size,contentType,isInline)` +
    `&$top=50`;

  try {
    const data = await graphGet(url, accessToken);
    return data.value || [];
  } catch (err) {
    // Log per map maar gooi niet alles weg bij één fout
    console.warn(`[Graph] Map ${folderId} overgeslagen:`, err.message);
    return [];
  }
}

/**
 * Scan ALLE mappen en verzamel ongelezen mails met PDF/.docx bijlagen
 * @param {string} accessToken
 * @param {function} onProgress - callback(statusText) voor voortgangsmeldingen
 * @returns {Promise<Array<AttachmentItem>>}
 */
export async function scanAllFoldersForAttachments(accessToken, onProgress) {
  onProgress?.("Ongelezen mails ophalen...");
  const folders = await getAllMailFolders(accessToken);

  onProgress?.("Mails analyseren op bijlagen...");
  const allMessages = [];

  for (const folder of folders) {
    const messages = await getUnreadMessagesWithAttachments(folder.id, accessToken);
    allMessages.push(...messages);

    // Stop bij 500 berichten (configureerbaar)
    if (allMessages.length >= 500) break;
  }

  onProgress?.("PDF- en Word-bijlagen identificeren...");

  // Filter: alleen PDF en DOCX, geen inline afbeeldingen
  const attachmentItems = [];

  for (const message of allMessages) {
    const attachments = (message.attachments || []).filter((att) => {
      if (att.isInline) return false;
      const name = (att.name || "").toLowerCase();
      return name.endsWith(".pdf") || name.endsWith(".docx");
    });

    for (const att of attachments) {
      attachmentItems.push({
        attachmentId: att.id,
        messageId: message.id,
        fileName: att.name,
        fileSize: att.size,
        contentType: att.contentType,
        fileType: att.name.toLowerCase().endsWith(".pdf") ? "pdf" : "docx",
        sender: {
          name: message.from?.emailAddress?.name || "Onbekend",
          email: message.from?.emailAddress?.address || "",
        },
        subject: message.subject || "(geen onderwerp)",
        receivedDate: message.receivedDateTime,
        isLargeFile: att.size > 10 * 1024 * 1024, // > 10 MB
        selected: false,
      });
    }
  }

  onProgress?.("Lijst voorbereiden...");
  console.log(`[Graph] ${attachmentItems.length} bijlagen gevonden`);

  // Sorteer op ontvangstdatum (nieuwste eerst)
  return attachmentItems.sort(
    (a, b) => new Date(b.receivedDate) - new Date(a.receivedDate)
  );
}

/**
 * Download de inhoud van een bijlage als Blob
 * @param {string} messageId
 * @param {string} attachmentId
 * @param {string} accessToken
 * @returns {Promise<Blob>}
 */
export async function downloadAttachment(messageId, attachmentId, accessToken) {
  const url = `${GRAPH_BASE}/me/messages/${messageId}/attachments/${attachmentId}/$value`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) {
    throw new Error(`Download mislukt voor bijlage ${attachmentId}: ${response.status}`);
  }

  return response.blob();
}

/**
 * Markeer een mail als gelezen via PATCH
 * @param {string} messageId
 * @param {string} accessToken
 */
export async function markMessageAsRead(messageId, accessToken) {
  const url = `${GRAPH_BASE}/me/messages/${messageId}`;

  const response = await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ isRead: true }),
  });

  if (!response.ok) {
    throw new Error(`Markeren als gelezen mislukt voor ${messageId}: ${response.status}`);
  }
}