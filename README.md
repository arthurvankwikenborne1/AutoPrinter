# Outlook PDF Print Agent — Setup Instructies

## Stap 1 — Azure App Registration aanmaken

1. Ga naar [portal.azure.com](https://portal.azure.com)
2. Zoek naar **"App registrations"** → klik **"New registration"**
3. Vul in:
   - **Name:** `Outlook PDF Print Agent`
   - **Supported account types:** *Accounts in this organizational directory only*
   - **Redirect URI:** Kies `Single-page application (SPA)` → `https://localhost:3000/src/auth/auth-redirect.html`
4. Klik **Register**
5. Kopieer de **Application (client) ID** — je hebt dit zo dadelijk nodig

### API Permissions instellen
1. Klik op **"API permissions"** → **"Add a permission"** → **"Microsoft Graph"** → **"Delegated permissions"**
2. Voeg toe: `Mail.Read` en `Mail.ReadWrite`
3. Klik **"Grant admin consent"** (vereist tenant-beheerder rechten)

---

## Stap 2 — Client ID invullen

Vervang `JOUW_CLIENT_ID_HIER` in de volgende twee bestanden:

```
src/auth/auth.js          → regel: clientId: "JOUW_CLIENT_ID_HIER"
src/auth/auth-redirect.html → regel: clientId: "JOUW_CLIENT_ID_HIER"
```

---

## Stap 3 — Node.js installeren

Download en installeer [Node.js LTS](https://nodejs.org) (versie 18 of hoger).

Controleer installatie:
```bash
node --version
npm --version
```

---

## Stap 4 — Afhankelijkheden installeren

Open een terminal in de projectmap en voer uit:

```bash
npm install
```

---

## Stap 5 — Dev-certificaten installeren (eénmalig)

Outlook vereist HTTPS voor add-ins. Installeer lokale dev-certificaten:

```bash
npx office-addin-dev-certs install --machine
```

> ⚠️ Vereist beheerdersrechten. Accepteer het certificaat in Windows wanneer gevraagd.

---

## Stap 6 — De add-in starten

```bash
npm start
```

De server draait nu op `https://localhost:3000`.

---

## Stap 7 — Manifest laden in Outlook

### Optie A: Sideloading (voor testen)
1. Open **Outlook desktop**
2. Klik op een mail → Ga naar **Home** lint
3. Klik op **"Get Add-ins"** (of "Store")
4. Kies **"My add-ins"** → **"Add a custom add-in"** → **"Add from file..."**
5. Selecteer `manifest/manifest.xml`

### Optie B: Via Microsoft 365 Admin Center (voor uitrol)
1. Ga naar [admin.microsoft.com](https://admin.microsoft.com)
2. **Settings** → **Integrated apps** → **Upload custom apps**
3. Upload `manifest/manifest.xml`

---

## Projectstructuur

```
outlook-print-agent/
├── manifest/
│   └── manifest.xml          ← Add-in manifest voor Outlook
├── src/
│   ├── auth/
│   │   ├── auth.js           ← MSAL authenticatie module
│   │   └── auth-redirect.html ← OAuth redirect pagina
│   ├── api/
│   │   └── graphApi.js       ← Microsoft Graph API aanroepen
│   ├── taskpane/
│   │   ├── taskpane.html     ← Taakvenster UI
│   │   ├── taskpane.css      ← Styling
│   │   └── taskpane.js       ← Hoofdcontroller
│   └── utils/
│       └── i18n.js           ← NL/EN vertalingen
├── assets/                   ← Iconen (toe te voegen)
├── package.json
└── README.md
```

---

## Fase overzicht

| Fase | Status | Inhoud |
|------|--------|--------|
| **Fase 1** | ✅ Klaar | Scaffold, Graph auth, scan + lijstweergave |
| Fase 2 | 🔜 | Selectie, printen, gelezen markeren |
| Fase 3 | 🔜 | Historiek, foutafhandeling, NL/EN volledig |
| Fase 4 | 🔜 | Piloottest |
| Fase 5 | 🔜 | Bedrijfsbrede uitrol |

---

## Vereisten

- Windows 10/11
- Outlook desktop (Microsoft 365, versie 16.0.x+)
- Node.js 18+
- Microsoft 365 licentie met Exchange Online
- Azure tenant-beheerder voor App Registration
