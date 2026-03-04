from __future__ import annotations

import json
import sys
from pathlib import Path


def _emit_result(ok: bool, directory: str | None, reason: str, code: int) -> int:
    payload = {
        "ok": bool(ok),
        "directory": directory or "",
        "reason": reason or "",
    }
    # ASCII-safe transport for PowerShell: no locale/codepage dependency.
    print(json.dumps(payload, ensure_ascii=True))
    return code


def _import_win32com_client():
    try:
        import win32com.client as client  # type: ignore
    except Exception as exc:
        return None, exc
    return client, None


def _split_document_path(path_value: str) -> str | None:
    raw = (path_value or "").strip().strip('"').replace("/", "\\")
    if not raw:
        return None

    p = Path(raw)
    if p.suffix:
        return str(p.parent)
    return str(p)


def _candidate_doc_paths(doc) -> list[str]:
    values: list[str] = []
    for attr in ("PathName", "Path", "Name"):
        try:
            value = getattr(doc, attr)
        except Exception:
            continue

        if not value:
            continue

        text = str(value).strip()
        if text:
            values.append(text)

    return values


def _iter_collection(collection):
    if collection is None:
        return

    count = None
    try:
        count = int(collection.Count)
    except Exception:
        count = None

    if count is not None and count > 0:
        for index in range(count):
            item = None
            try:
                item = collection.Item(index)
            except Exception:
                try:
                    item = collection.Item(index + 1)
                except Exception:
                    item = None
            if item is not None:
                yield item
        return

    try:
        for item in collection:
            if item is not None:
                yield item
    except Exception:
        return


def _resolve_from_document(doc) -> str | None:
    if doc is None:
        return None

    for value in _candidate_doc_paths(doc):
        directory = _split_document_path(value)
        if directory:
            return directory

    return None


def _resolve_active_kompas_directory() -> tuple[str | None, str]:
    win32_client, import_error = _import_win32com_client()
    if win32_client is None:
        return None, (
            "Python module win32com.client is unavailable. "
            f"Install pywin32 for this interpreter. Details: {import_error}"
        )

    app5 = None
    for progid in ("Kompas.Application.5", "KOMPAS.Application.5"):
        try:
            app5 = win32_client.GetObject(Class=progid)
            break
        except Exception:
            continue

    if app5 is None:
        return None, "KOMPAS instance is not running in current COM context."

    app7 = None
    try:
        app7 = getattr(app5, "ksGetApplication7")
    except Exception:
        app7 = None

    if callable(app7):
        try:
            app7 = app7()
        except Exception:
            app7 = None

    doc = None
    if app7 is not None:
        try:
            doc = getattr(app7, "ActiveDocument", None)
        except Exception:
            doc = None

    if doc is None:
        try:
            doc = getattr(app5, "ActiveDocument", None)
        except Exception:
            doc = None

    directory = _resolve_from_document(doc)
    if directory:
        return directory, "ActiveDocument"

    docs = None
    if app7 is not None:
        try:
            docs = getattr(app7, "Documents", None)
        except Exception:
            docs = None
    if docs is None:
        try:
            docs = getattr(app5, "Documents", None)
        except Exception:
            docs = None

    for item in _iter_collection(docs):
        directory = _resolve_from_document(item)
        if directory:
            return directory, "Documents collection"

    return None, "No saved document path found in ActiveDocument/Documents."


def main() -> int:
    try:
        directory, reason = _resolve_active_kompas_directory()
        if not directory:
            return _emit_result(False, None, reason, 3)

        return _emit_result(True, directory, reason, 0)
    except Exception as exc:
        return _emit_result(False, None, f"Unexpected error: {exc}", 9)


if __name__ == "__main__":
    raise SystemExit(main())
