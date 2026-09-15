# TETORD architecture — 3.6.0

## Runtime boundaries

```text
Electron main process
  electron/main.js
       │ secure IPC
       ▼
Preload bridge
  electron/preload.js
       │ contextBridge
       ▼
Renderer
  MiWord/js/app.js          legacy Writer engine
  MiWord/js/office.js       Office orchestration boundary
  MiWord/src/renderer/
    core/storage.js
    modules/app-module.js
    modules/writer.js
    modules/calc.js
    modules/slides.js
    modules/files.js
    modules/messenger.js
```

## Storage

Internal TETORD data is stored below the user's `Documents/TETORD` directory. Renderer code never receives Node.js `fs`; it calls a narrow preload API. JSON writes are atomic (temporary file + rename), and slide images are stored as independent binary assets rather than Base64 blobs inside JSON/localStorage.

`localStorage` remains available to the legacy Writer engine and as a browser-development fallback. Slides no longer use it for normal persistence; a one-time migration imports the previous `tetord.office.slides` value when present.

## Security

- `contextIsolation: true`
- `nodeIntegration: false`
- `sandbox: true`
- Storage IPC accepts only relative paths resolved beneath `Documents/TETORD`.
- JSON is written atomically.
- Slides rendering uses DOM APIs and `textContent` instead of HTML string interpolation.

## Versioning

`VERSION` is the source of truth. Use `npm run sync-version` to update the current-version metadata and regenerate the distribution template. Historical changelog entries are intentionally not rewritten.
