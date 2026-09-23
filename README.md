# WordTimestamp — Authorship Evidence Recording Plugin for Microsoft Word

## What It Does

Records every edit event made to a Word document with precise timestamps, providing forensic evidence of authorship. Captures:

- **What changed** — text inserted/deleted/modified
- **Where it happened** — offset position in the document
- **When it happened** — millisecond-precision timestamps
- **Who made the change** — user identity captured on session start
- **Session integrity** — SHA-256 hash of the recording data
- **Change type** — text, formatting, image, and other document modifications
- **Document context** — paragraph/section location of changes

Exports to `.wtp` files (gzip-compressed JSON) for backup and sharing. Timeline player replays all changes at 0.25x–32x speed with zoom functionality for detailed evidence review.

## Architecture

```
┌─────────── Microsoft Word ───────────────┐
│                                          │
│   Document Content                        │
│       ↕ Office.js API                    │
│       ├─ document.onChanged (events)     │
│       └─ browser.indexedDB (storage)     │
├──────────────────────────────────────────┤
│  Task Pane (HTML/CSS/JS iframe)          │
│       🔴 Record   ⏹ Stop   ⏸ Pause      │
│       [██████░░░░] ◀ ▶ 🔍 Speed select   │
│       🔍 Zoom In/Out      👤 Author info   │
│       💾 Save    📄 Export .wtp          │
├──────────────────────────────────────────┤
│  Browser IndexedDB (Device-Only)         │
│       Sessions + recordings              │
│       → Export as .wtp for backup        │
└──────────────────────────────────────────┘
```

## Installation

**This plugin uses Office.js web add-in technology.** Word add-ins must load from an `http://` or `https://` URL — `file://` is not supported. Two install paths:

### Method 1: Static hosting (recommended — one-time setup)

1. Upload the files in `dist/` to any static host (GitHub Pages, Netlify, your web server, etc.)
2. Edit `dist/manifest.xml` → replace both `YOUR_DOMAIN_OR_HOST` with your actual domain (e.g., `https://yourdomain.com`)
3. In Word: **Home → Add-ins → Browse my add-ins → Upload a custom add-in** and select the manifest
4. Done — plugin works forever, data stays local in IndexedDB

### Method 2: Local dev server (quick test)
```
cd "dist/"
python3 -m http.server 3000
# OR with Node.js installed: cd "../" && npm run serve
```
1. Edit `dist/manifest.xml` → set SourceLocation to `http://localhost:3000/taskpane.html` (or your local HTTPS URL)
2. In Word: **Home → Add-ins → Browse my add-ins** → sideload the manifest file

### File structure
```
Word Timestamp Plugin/
├── README.md               ← You are here
├── manifest.xml            ← Add-in manifest (drag & drop install)
├── src/
│   ├── core/
│   │   ├── recorder.js      ← Captures document.onChanged events with author tracking
│   │   ├── player.js        ← Timeline playback engine with zoom support
│   │   ├── format_encoding.js ← Export/import + SHA-256 hash with metadata
│   │   └── index.js         ← Barrel re-export
│   └── ui/
│       ├── taskpane.html    ← Task pane UI (host shell)
│       ├── taskpane.js      ← UI controller logic with timeline visualization
│       └── styles.css       ← Plugin styling
└── dist/                   ← Built output (ready to deploy)
```

### Features in this version
- Live recording with event-by-event capture
- Pause/resume without losing timeline continuity
- Playback at 0.25x–32x speed with seek slider
- Timeline zoom functionality for detailed evidence review
- Author attribution tracking
- Enhanced change metadata (type, location, formatting info)
- Auto-save to device IndexedDB every 5 seconds
- Export recordings as `.wtp` files (gzip-compressed JSON)
- Import `.wtp` files with integrity verification
- On-screen error banner for troubleshooting
- Debug status panel with filtering
- Contextual change details in event list
- Session start/end time tracking

### Limitations (Office.js platform limits)
- **Batched events, not keystrokes**: Edits batched ~100ms during typing — cannot capture per-keypress timing
- **No cursor tracking**: Cannot continuously track cursor position between edits
- **Requires web hosting**: Office.js add-ins load from URLs, not local files
