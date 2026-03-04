from __future__ import annotations

import sys
from pathlib import Path


EXIT_OK = 0
EXIT_USAGE = 1
EXIT_INPUT_NOT_FOUND = 2
EXIT_PYWIN32_MISSING = 10
EXIT_KOMPAS_ERROR = 20
EXIT_NO_ACTIVE_DOC = 30
EXIT_NO_ACTIVE_VIEW = 31
EXIT_TABLE_CREATE = 40
EXIT_TABLE_NOT_VISIBLE = 42


def log(message: str) -> None:
    print(message)


def _import_win32com_client():
    try:
        import win32com.client as client  # type: ignore
    except Exception as exc:
        log(
            "ERROR: Python module win32com.client is unavailable. "
            f"Install pywin32 for this interpreter. Details: {exc}"
        )
        return None
    return client


def connect_kompas():
    win32_client = _import_win32com_client()
    if win32_client is None:
        return None

    try:
        kompas5 = win32_client.GetObject(Class="Kompas.Application.5")
    except Exception:
        kompas5 = win32_client.Dispatch("Kompas.Application.5")

    kompas5.Visible = True
    kompas7 = kompas5.ksGetApplication7
    return kompas5, kompas7, win32_client


def main() -> int:
    if len(sys.argv) != 2:
        log("USAGE: python insert_tbl_bridge.py <table.tbl>")
        return EXIT_USAGE

    tbl_path = Path(sys.argv[1]).resolve()
    if not tbl_path.exists():
        log(f"ERROR: .tbl file not found: {tbl_path}")
        return EXIT_INPUT_NOT_FOUND

    try:
        connected = connect_kompas()
        if connected is None:
            return EXIT_PYWIN32_MISSING
        _, kompas7, win32_client = connected
    except Exception as exc:
        log(f"ERROR: failed to connect to KOMPAS COM: {exc}")
        return EXIT_KOMPAS_ERROR

    if kompas7 is None:
        log("ERROR: failed to get IKompasAPI7 application object from KOMPAS.")
        return EXIT_KOMPAS_ERROR

    doc = getattr(kompas7, "ActiveDocument", None)
    if doc is None:
        log("ERROR: no active KOMPAS document. Open Drawing/Fragment first.")
        return EXIT_NO_ACTIVE_DOC

    try:
        doc2d = win32_client.CastTo(doc, "IKompasDocument2D")
    except Exception as exc:
        log(f"ERROR: active document is not 2D (IKompasDocument2D): {exc}")
        return EXIT_NO_ACTIVE_DOC

    try:
        doc_path = getattr(doc2d, "PathName", "") or getattr(doc2d, "Name", "")
    except Exception:
        doc_path = ""
    if doc_path:
        log(f"INFO: active 2D document: {doc_path}")

    try:
        active_view = doc2d.ViewsAndLayersManager.Views.ActiveView
    except Exception as exc:
        log(f"ERROR: failed to get ActiveView: {exc}")
        return EXIT_NO_ACTIVE_VIEW

    if active_view is None:
        log("ERROR: ActiveView is missing.")
        return EXIT_NO_ACTIVE_VIEW

    try:
        symbols_container = win32_client.CastTo(active_view, "ISymbols2DContainer")
        drawing_tables = symbols_container.DrawingTables
    except Exception as exc:
        log(f"ERROR: failed to access DrawingTables: {exc}")
        return EXIT_TABLE_CREATE

    try:
        before_count = int(drawing_tables.Count)
    except Exception:
        before_count = -1

    try:
        table = drawing_tables.Load(str(tbl_path))
    except Exception as exc:
        log(f"ERROR: failed to load .tbl into active view: {exc}")
        return EXIT_TABLE_CREATE

    if table is None:
        log("ERROR: DrawingTables.Load returned null.")
        return EXIT_TABLE_CREATE

    # Shift loaded table a bit so it does not overlap the source position exactly.
    old_x = None
    old_y = None
    new_x = None
    new_y = None
    try:
        old_x = float(table.X)
        old_y = float(table.Y)
        new_x = old_x + 10.0
        new_y = old_y + 10.0
        table.X = new_x
        table.Y = new_y
    except Exception:
        pass

    try:
        table.Update()
    except Exception:
        pass

    try:
        active_view.Update()
    except Exception:
        pass

    try:
        after_count = int(drawing_tables.Count)
    except Exception:
        after_count = -1

    if before_count >= 0 and after_count >= 0:
        log(f"INFO: tables in active view: {before_count} -> {after_count}")

    if old_x is not None and old_y is not None and new_x is not None and new_y is not None:
        log(f"INFO: table position shifted: ({old_x:.2f}, {old_y:.2f}) -> ({new_x:.2f}, {new_y:.2f})")

    if before_count >= 0 and after_count == before_count:
        log("ERROR: table count did not change after load; insertion not confirmed visually.")
        return EXIT_TABLE_NOT_VISIBLE

    log(f"OK: inserted table from file: {tbl_path}")
    return EXIT_OK


if __name__ == "__main__":
    raise SystemExit(main())
