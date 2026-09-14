# TETORD 4.2.0 — Git upload

This folder is prepared to be used as the repository root.

## Upload

1. Extract this ZIP.
2. Open the extracted folder in GitHub Desktop or Git.
3. Commit the files with:
   `Release TETORD 4.2.0 — Core Refactor`
4. Push to the `main` branch.

## Versioning

`VERSION` is the single source of truth. Run `npm run sync-version` after changing it. The synchronizer updates the current-version metadata atomically and creates `dist-template/`.

## Important

- `node_modules/` and `dist/` are ignored by `.gitignore`.
- Legacy Python patch scripts are intentionally removed.
- Electron internal application data is stored through the secure preload IPC bridge under the user's Documents/TETORD directory.
