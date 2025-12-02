## Von SharePoint-Site-ID zu Datei-IDs

Dieser Leitfaden zeigt, wie du mit den vorhandenen MCP-Tools aus einer **SharePoint-Site-ID** die **Datei-IDs** (DriveItem-IDs) in den zugehörigen Dokumentbibliotheken ermitteln kannst.

### Voraussetzungen

- **Server im Organization-Mode**: MCP-Server mit `--org-mode` gestartet
- **Authentifizierung**: Du bist angemeldet (z. B. über das `login`-Tool)
- **Berechtigungen** (in der Azure-/Entra-App und für den Benutzer):
  - `Sites.Read.All`
  - `Files.Read.All` bzw. `Files.Read`

---

### 1. Dokumentbibliotheken (Drives) der Site auflisten

Jede SharePoint-Site kann mehrere Dokumentbibliotheken (Drives) haben. Diese enthalten die eigentlichen Dateien.

- **Tool**: `list-sharepoint-site-drives`
- **Parameter**:
  - `site-id`: die Site-ID der gewünschten SharePoint-Site

**Beispiel:**

```json
{
  "name": "list-sharepoint-site-drives",
  "arguments": {
    "site-id": "site-id-xyz",
    "top": 10
  }
}
```

Aus der Antwort entnimmst du für jede Bibliothek u. a.:

- `id` → **`driveId`** (Dokumentbibliothek)
- `name` (z. B. „Dokumente“, „Shared Documents“)
- ggf. `webUrl`, `driveType`, etc.

---

### 2. Dateien in einem Drive/Folder auflisten (DriveItem-IDs)

Um an konkrete Datei-IDs zu kommen, arbeitest du auf Ebene der Drives.

#### 4.1 Ordner-Inhalt auflisten

- **Tool**: `list-folder-files`
- **Parameter**:
  - `driveId`: die `id` des Drives aus Schritt 3
  - `driveItemId`: die ID des Ordners, dessen Inhalt du sehen willst  
    - Für den **Root-Ordner** kannst du zuvor den Drive über passende Tools abfragen, oder du verwendest eine dir bereits bekannte Root-`id`.

**Beispiel (Ordner-Inhalt):**

```json
{
  "name": "list-folder-files",
  "arguments": {
    "driveId": "drive-id-abc",
    "driveItemId": "root-or-folder-id",
    "top": 50
  }
}
```

#### 4.2 Wichtige Felder in der Antwort

In der Antwort zu `list-folder-files` erhältst du pro Eintrag (`DriveItem`) u. a.:

- `id` → **Datei-ID / DriveItem-ID**
- `name` → Dateiname
- `folder` oder `file` → Typ (Ordner oder Datei)
- `webUrl` → Browser-URL
- `sharepointIds` → zusätzliche SharePoint-IDs (Liste, ListItem, Site, etc.)

Diese `id` verwendest du anschließend z. B. für:

- `get-onedrive-file` (Details einer einzelnen Datei)
- `download-file-to-local` (Datei auf den MCP-Server herunterladen)
- `delete-onedrive-file` (Datei löschen – Vorsicht!)

---

### 3. Kurz-Zusammenfassung

- **Ausgangspunkt**: vorhandene `site-id` der SharePoint-Site  
- **Drives der Site**: `list-sharepoint-site-drives` → `driveId`  
- **Dateien/Folders**: `list-folder-files` (mit `driveId` + `driveItemId`) → `id` pro Datei

---

### 4. Komfort-Tool: Alles in einem Aufruf (`list-sharepoint-site-files`)

Um die Schritte 1–3 nicht manuell kombinieren zu müssen, gibt es das Convenience-Tool `list-sharepoint-site-files`.  
Es übernimmt:

- das Auflisten der Drives der Site (`list-sharepoint-site-drives`)
- die Auswahl einer Dokumentbibliothek (optional gesteuert über `driveId` oder `driveName`)
- und das rekursive Durchlaufen der Ordnerstruktur (`list-folder-files`), bis zu einer definierten Tiefe.

**Tool-Name**: `list-sharepoint-site-files`  

**Parameter (wichtigste):**

- `siteId` (**required**): SharePoint-Site-ID
- `structure` (**required**): `"flat"` oder `"tree"`
- `driveId` (optional): explizite Drive-ID (Dokumentbibliothek)
- `driveName` (optional): Anzeigename der Bibliothek (z. B. „Dokumente“, „Documents“)
- `includeFolders` (optional, nur für `structure = "flat"` relevant): Ordner in der Liste mit ausgeben
- `maxDepth` (optional): maximale Ordner-Tiefe ab Bibliotheks-Root (Default ca. 10, Hard-Limit 20)
- `filter` (optional): einfacher Namensfilter, z. B. `"*.pdf"` oder `"report"`

#### 4.1 Beispiel: Flache Liste aller Dateien

```json
{
  "name": "list-sharepoint-site-files",
  "arguments": {
    "siteId": "site-id-xyz",
    "structure": "flat",
    "driveName": "Dokumente",
    "includeFolders": false,
    "maxDepth": 5,
    "filter": "*.pdf"
  }
}
```

Die Antwort enthält u. a.:

- `structure: "flat"`
- `driveId`, `driveName`, `siteId`
- `payload.items[]` – eine flache Liste mit allen passenden Dateien (und optional Ordnern)

#### 4.2 Beispiel: Baumstruktur

```json
{
  "name": "list-sharepoint-site-files",
  "arguments": {
    "siteId": "site-id-xyz",
    "structure": "tree",
    "driveName": "Dokumente",
    "maxDepth": 5
  }
}
```

Die Antwort enthält u. a.:

- `structure: "tree"`
- `payload.root` – Root-Knoten der Bibliothek
- `payload.root.children[]` – rekursiver Baum aus Ordnern und Dateien



