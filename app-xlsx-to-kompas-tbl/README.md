# app-xlsx-to-kompas-tbl

Quick-path app for KOMPAS-3D:
- button in KOMPAS runs a local launcher;
- launcher opens a small GUI window;
- GUI imports Excel `.xlsx` to KOMPAS `.tbl` using existing exporter from:
`..\xlsx-to-kompas-tbl\src\create_tbl.vbs`.

## Folder layout

- `bin/app-xlsx-to-kompas-tbl.exe` - GUI launcher with app icon (for taskbar and KOMPAS command icon).
- `assets/app.ico` - app icon file.
- `scripts/run_gui.cmd` - starts GUI.
- `scripts/run_gui.vbs` - starts GUI without console window (prefers launcher `.exe`).
- `scripts/gui_import.ps1` - WinForms GUI.
- `scripts/kompas_button_launcher.vbs` - script for KOMPAS button command.
- `scripts/build_launcher.ps1` - rebuild launcher `.exe` from C# source.
- `scripts/AppLauncher.cs` - launcher source.
- `scripts/bind_to_kompas.ps1` - prepares binding values for KOMPAS button.
- `scripts/selfcheck.ps1` - quick environment check.
- `out/` - output `.tbl`.

## Manual run

```bat
cd C:\_GIT_\kompas3d-utility\app-xlsx-to-kompas-tbl
.\bin\app-xlsx-to-kompas-tbl.exe
```

If launcher `.exe` is missing:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_launcher.ps1
```

In the window:
1. Select input `.xlsx`.
2. Output `.tbl` is auto-reset to the folder of current active KOMPAS drawing (if saved) on app start and when input `.xlsx` changes; otherwise fallback is `app-xlsx-to-kompas-tbl\out`.
3. Or drag-and-drop `.xlsx` directly into the app window.
4. Configure table layout in `Full settings`:
   - `mode=cell` for direct `Cell width/height`;
   - `mode=table` for `Table width/height` with auto cell-size calculation.
5. Click `Apply settings` once (or just run export, settings are auto-saved).
6. Keep a 2D KOMPAS document active (Drawing/Fragment).
7. Click export icon to create/update `.tbl` from `.xlsx`.
8. Click insert icon to load existing `.tbl` into active KOMPAS document (`DrawingTables.Load`).

## Bind to KOMPAS-3D

Run helper script:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bind_to_kompas.ps1
```

It writes `out\kompas_button_binding.txt`, copies command line to clipboard, and prints current COM status.

In KOMPAS UI customization (custom/external command), set values from `out\kompas_button_binding.txt`.
Recommended:

- Program: `C:\_GIT_\kompas3d-utility\app-xlsx-to-kompas-tbl\bin\app-xlsx-to-kompas-tbl.exe`
- Arguments: *(empty)*
- Name (example): `Excel -> TBL`
- Icon (if KOMPAS allows explicit icon path):
`C:\_GIT_\kompas3d-utility\app-xlsx-to-kompas-tbl\assets\app.ico`

After that, pressing the KOMPAS button opens the GUI importer.

## Self-check

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\selfcheck.ps1
```

## Notes

- This is a fast integration path, not a compiled plugin.
- Export logic remains in `xlsx-to-kompas-tbl`.
- GUI now uses a full settings profile (`app-xlsx-to-kompas-tbl\config\app_settings.json`).
- Layout ini for exporter is generated automatically from settings and passed as the 3rd argument to `create_tbl.vbs`.
- Drag-and-drop `.xlsx` supports elevated launch path too (extra UAC message-filter handling is enabled at runtime).
- Launcher starts GUI without console and keeps caller context for stable KOMPAS COM detection.
- If direct PowerShell COM path detection fails, GUI uses Python `pywin32` fallback (`scripts/resolve_kompas_doc_dir.py`) to resolve active KOMPAS document directory.
- If `ActiveView.DrawingTables` is unavailable in pure VBS COM binding, exporter uses python fallback bridge (`pywin32`).
- GUI supports a dedicated insert action via `xlsx-to-kompas-tbl/src/insert_tbl_bridge.py` (loads `.tbl` to active 2D view).
- KOMPAS v24 stores part of UI bindings in non-trivial internal configs; helper script prepares reliable command values and diagnostics for manual button mapping.
