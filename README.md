# kompas3d-utility

`docs/` now contains a remote-updatable Pages UI for install-once `WebBridge.Utility`. The shell is tab-based and generic by design: the first working module is `xlsx-to-kompas-tbl`, while the remaining tabs are placeholders for future utilities from this repository.

The browser side is plain HTML/CSS/JS only. XLSX parsing, preview, layout math, token derivation, session registration, WebSocket `hello`/`heartbeat`, and browser-managed `/config/load` all happen in the page. Product flow no longer depends on repo-local PowerShell helpers or `workspaceRoot`; the page auto-connects to `http://127.0.0.1:38741`, tries derived and legacy tokens, and loads its runtime overlay with `persist:false`.

## Local start

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\start_pages_xlsx_to_kompas_tbl.ps1
```

The bootstrap config now only points `WebBridge.Utility` at the UI URL and a fallback token. The page opens without query params and auto-connects on load.

Stop local runtime:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\stop_pages_xlsx_to_kompas_tbl.ps1
```

## Tests

Existing utility tests:

```powershell
py -3 -m pytest -v .\xlsx-to-kompas-tbl\tests
```

Browser E2E for the tabbed executor:

```powershell
node .\tools\e2e\run_e2e.mjs --browser msedge
```

Node checks for token derivation, runtime overlay decisions, linked layout math, and export chunking:

```powershell
node --test .\tools\runtime\tests\pages-executor.test.mjs
```
