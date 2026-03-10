# kompas3d-utility

`docs/` now contains a static GitHub Pages style UI Executor for `WebBridge.Utility`. The shell is tab-based and generic by design: the first working module is `xlsx-to-kompas-tbl`, while the remaining tabs are placeholders for future utilities from this repository.

The browser side is plain HTML/CSS/JS only. XLSX parsing, preview, layout math, runtime profile assembly, session registration, WebSocket `hello`/`heartbeat`, and `/config/load` all happen in the page. Local KOMPAS/file actions stay behind `WebBridge.Utility` and the thin PowerShell bridge in [scripts/webbridge_xlsx_to_kompas_tbl.ps1](/c:/_GIT_/kompas3d-utility/scripts/webbridge_xlsx_to_kompas_tbl.ps1).

## Local start

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\start_pages_xlsx_to_kompas_tbl.ps1
```

The script prints a launch URL with `utilityUrl`, `pairingToken`, `autoConnect=1`, and `workspaceRoot`.

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
