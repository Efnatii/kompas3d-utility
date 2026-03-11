# kompas3d-utility

This repository now contains only the remote-updatable KOMPAS Pages web UI and the minimal dev tooling around it. The production page lives at the repository root: `index.html` + `assets/`.

The browser side is plain HTML/CSS/JS only. XLSX parsing, preview, layout math, token derivation, session registration, WebSocket `hello`/`heartbeat`, and browser-managed `/config/load` all happen in the page. Product flow no longer depends on repo-local PowerShell, Python, VBS, CMD, desktop launchers, `workspaceRoot`, or query-string bridge bootstrap; the page auto-connects to `http://127.0.0.1:38741`, tries derived and legacy tokens, and loads its runtime overlay with `persist:false`.

## Layout

- `index.html` and `assets/` are the deployable web UI.
- `tools/web/` contains dev-only helpers for local serving, runtime config generation, fixtures, tests, and E2E.

## Tests

Browser E2E for the tabbed executor:

```powershell
node .\tools\web\run_e2e.mjs --browser msedge
```

Node checks for token derivation, runtime overlay decisions, linked layout math, and export chunking:

```powershell
node --test .\tools\web\tests\pages-executor.test.mjs
```
*** Delete File: c:\__MY_PROJECTS__\git\kompas3d-utility\index.web.tmp.html
