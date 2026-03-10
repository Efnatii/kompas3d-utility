# kompas3d-utility Pages

Static Pages UI lives in `docs/`.

Current browser-first utility:

- `xlsx-to-kompas-tbl`

What runs in the browser:

- XLSX upload
- first-sheet parsing
- matrix preview
- size calculation
- compact multi-tab UI

What stays local through `WebBridge.Utility`:

- KOMPAS document status
- opening a sample 2D drawing for tests
- creating and saving `.tbl` in the active document through a PowerShell + C# bridge

## Local runtime

Start local Pages + `WebBridge.Utility`:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\start_pages_xlsx_to_kompas_tbl.ps1
```

Stop:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\stop_pages_xlsx_to_kompas_tbl.ps1
```

## Runtime config

Generate config only:

```powershell
node .\tools\runtime\build_runtime_config.mjs `
  --output .\out\web-pages-runtime\config.runtime.json `
  --listen-url http://127.0.0.1:38741 `
  --ui-url https://<user>.github.io/<repo>/ `
  --origin https://<user>.github.io `
  --pairing-token your-local-token
```

## E2E

Run Pages e2e:

```powershell
node .\tools\e2e\run_e2e.mjs --browser msedge
```

Artifacts are written to `out/e2e/<timestamp>/`.
