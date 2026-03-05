from __future__ import annotations

import argparse
import json
import os
import sys
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from sync_logic import TextElement, build_bindings, build_signature, choose_sync_action, group_vertical, normalize_text

EXIT_OK = 0
EXIT_USAGE = 1
EXIT_PYWIN32_MISSING = 10
EXIT_KOMPAS_ERROR = 20
EXIT_EXCEL_ERROR = 21

XL_RED = 3
XL_NONE = -4142
STATE_VERSION = 1
SHEET_DEFAULT = "SyncData"
LOG_FILE: Path | None = None


def log(message: str) -> None:
    stamp = datetime.now().strftime("%H:%M:%S")
    line = f"[{stamp}] {message}"
    print(line, flush=True)
    if LOG_FILE is not None:
        try:
            LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with LOG_FILE.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except Exception:
            pass


@dataclass
class TextRuntime:
    text_id: str
    text: str
    x: float
    y: float
    item: Any


class SyncEngine:
    def __init__(self, corridor_mm: float, poll_ms: int, sheet_name: str):
        self.corridor_mm = abs(float(corridor_mm))
        self.poll_ms = max(int(poll_ms), 250)
        self.sheet_name = sheet_name

        self.win32 = None
        self.kompas5 = None
        self.kompas7 = None
        self.excel = None
        self.doc_states: dict[str, dict[str, Any]] = {}
        self.doc_state_paths: dict[str, Path] = {}
        self.last_wait = ""
        self.last_doc = ""
        self.last_save_ts = 0.0

    def run(self, once: bool = False) -> int:
        self._ensure_pywin32()
        error_count = 0
        try:
            while True:
                try:
                    self.tick()
                    error_count = 0
                except KeyboardInterrupt:
                    raise
                except Exception as exc:
                    error_count += 1
                    log(f"ERROR: sync tick failed (#{error_count}): {exc}")
                    tb = traceback.format_exc(limit=4).strip()
                    if tb:
                        log(tb)
                    if once:
                        return EXIT_USAGE
                    # Keep worker alive on transient COM failures.
                    time.sleep(min(5.0, 0.5 * error_count))
                if once:
                    return EXIT_OK
                time.sleep(self.poll_ms / 1000.0)
        finally:
            self._flush_dirty_states(force=True)

    def tick(self) -> None:
        context = self._get_active_doc_context()
        if context is None:
            self._flush_dirty_states(force=False)
            return

        doc_path, doc2d, active_view = context
        doc_key = os.path.normcase(str(doc_path))
        if doc_key != self.last_doc:
            self.last_doc = doc_key
            log(f"INFO: active drawing -> {doc_path}")

        workbook_path = doc_path.with_suffix(".xlsx")
        wb, ws = self._ensure_workbook(workbook_path)
        doc_state, state_file = self._get_doc_state(doc_path, workbook_path)

        elements = self._collect_texts(doc2d, active_view)
        if not elements:
            self._wait_once("no text elements found in active drawing")
            self._save_doc_state_if_needed(doc_state, state_file, force=False)
            return
        self.last_wait = ""

        snapshots = [TextElement(text_id=e.text_id, text=e.text, x=e.x, y=e.y) for e in elements]
        signature = build_signature(snapshots)
        by_id = {e.text_id: e for e in elements}

        state_changed = False
        must_rebuild = (
            doc_state.get("signature", "") != signature
            or abs(float(doc_state.get("corridor_mm", 0.0)) - self.corridor_mm) > 1e-9
            or not doc_state.get("bindings")
        )

        excel_changed = False
        kompas_changed = False
        if must_rebuild:
            self._rebuild_map(doc_state, elements, by_id, ws, signature)
            excel_changed = True
            state_changed = True

        sync_result = self._sync_cells(doc_state, by_id, ws)
        excel_changed = excel_changed or sync_result["excel_changed"]
        kompas_changed = kompas_changed or sync_result["kompas_changed"]
        state_changed = state_changed or sync_result["state_changed"]

        if self._mark_unbound(ws, doc_state.get("bindings", [])):
            excel_changed = True

        now = time.time()
        if excel_changed or (now - self.last_save_ts) > 4.0:
            self._save_workbook(wb)
            self.last_save_ts = now

        if kompas_changed:
            self._safe_call(active_view, ("Update", "Refresh", "Rebuild"))
            self._safe_call(doc2d, ("Update", "Refresh", "Rebuild"))

        if state_changed:
            doc_state["updated_at"] = datetime.now(timezone.utc).isoformat()
            doc_state["_dirty"] = True
        self._save_doc_state_if_needed(doc_state, state_file, force=False)

    def _ensure_pywin32(self) -> None:
        if self.win32 is not None:
            return
        try:
            import win32com.client as win32  # type: ignore
        except Exception as exc:
            raise RuntimeError(f"win32com.client unavailable: {exc}") from exc
        self.win32 = win32

    def _state_path_for_workbook(self, workbook_path: Path) -> Path:
        return workbook_path.with_suffix(".json")

    def _default_doc_state(self, doc_path: Path, workbook_path: Path) -> dict[str, Any]:
        return {
            "doc_path": str(doc_path),
            "workbook_path": str(workbook_path),
            "signature": "",
            "corridor_mm": self.corridor_mm,
            "bindings": [],
            "updated_at": "",
            "_dirty": True,
        }

    def _extract_doc_payload(self, loaded: Any, doc_key: str) -> dict[str, Any]:
        if not isinstance(loaded, dict):
            return {}
        direct = loaded.get("document")
        if isinstance(direct, dict):
            return direct
        docs = loaded.get("documents")
        if isinstance(docs, dict):
            selected = docs.get(doc_key)
            if isinstance(selected, dict):
                return selected
            if len(docs) == 1:
                only = next(iter(docs.values()))
                if isinstance(only, dict):
                    return only
        return loaded

    def _load_doc_state_from_file(self, state_file: Path, doc_path: Path, workbook_path: Path) -> dict[str, Any]:
        state = self._default_doc_state(doc_path, workbook_path)
        state["_dirty"] = False
        if not state_file.exists():
            return state

        loaded: Any = None
        try:
            loaded = json.loads(state_file.read_text(encoding="utf-8"))
        except Exception:
            return state

        payload = self._extract_doc_payload(loaded, os.path.normcase(str(doc_path)))
        if not isinstance(payload, dict):
            return state

        state["signature"] = normalize_text(payload.get("signature", ""))
        state["corridor_mm"] = self._safe_float(payload.get("corridor_mm"), self.corridor_mm)
        state["updated_at"] = normalize_text(payload.get("updated_at", ""))

        bindings: list[dict[str, Any]] = []
        raw_bindings = payload.get("bindings", [])
        if isinstance(raw_bindings, list):
            for raw in raw_bindings:
                if not isinstance(raw, dict):
                    continue
                row = max(1, self._safe_int(raw.get("row", 1), 1))
                col = max(1, self._safe_int(raw.get("col", 1), 1))
                text_id = normalize_text(raw.get("text_id", "")).strip()
                if not text_id:
                    continue
                bindings.append(
                    {
                        "row": row,
                        "col": col,
                        "text_id": text_id,
                        "last_excel": normalize_text(raw.get("last_excel", "")),
                        "last_kompas": normalize_text(raw.get("last_kompas", "")),
                    }
                )
        state["bindings"] = bindings
        return state

    def _save_doc_state_if_needed(self, doc_state: dict[str, Any], state_file: Path, force: bool) -> None:
        if not force and not bool(doc_state.get("_dirty")):
            return
        state_file.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": STATE_VERSION,
            "saved_at": datetime.now(timezone.utc).isoformat(),
            "document": {
                "doc_path": str(doc_state.get("doc_path", "")),
                "workbook_path": str(doc_state.get("workbook_path", "")),
                "signature": normalize_text(doc_state.get("signature", "")),
                "corridor_mm": self._safe_float(doc_state.get("corridor_mm"), self.corridor_mm),
                "bindings": doc_state.get("bindings", []),
                "updated_at": normalize_text(doc_state.get("updated_at", "")),
            },
        }
        tmp = state_file.with_suffix(state_file.suffix + ".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, state_file)
        doc_state["_dirty"] = False

    def _flush_dirty_states(self, force: bool) -> None:
        for key, doc_state in list(self.doc_states.items()):
            state_file = self.doc_state_paths.get(key)
            if state_file is None:
                continue
            self._save_doc_state_if_needed(doc_state, state_file, force=force)

    def _get_doc_state(self, doc_path: Path, workbook_path: Path) -> tuple[dict[str, Any], Path]:
        key = os.path.normcase(str(doc_path))
        state_file = self._state_path_for_workbook(workbook_path)
        state = self.doc_states.get(key)
        known_path = self.doc_state_paths.get(key)
        if state is None or known_path != state_file:
            state = self._load_doc_state_from_file(state_file, doc_path, workbook_path)
            self.doc_states[key] = state
            self.doc_state_paths[key] = state_file
            log(f"INFO: state json -> {state_file}")

        if state.get("doc_path") != str(doc_path):
            state["doc_path"] = str(doc_path)
            state["_dirty"] = True
        if state.get("workbook_path") != str(workbook_path):
            state["workbook_path"] = str(workbook_path)
            state["_dirty"] = True

        return state, state_file

    def _wait_once(self, reason: str) -> None:
        if reason != self.last_wait:
            self.last_wait = reason
            log(f"INFO: waiting - {reason}")

    def _connect_kompas(self) -> tuple[Any, Any]:
        self._ensure_pywin32()
        if self.kompas5 is None:
            for pid in ("Kompas.Application.5", "KOMPAS.Application.5"):
                try:
                    self.kompas5 = self.win32.GetObject(Class=pid)
                    break
                except Exception:
                    continue
            if self.kompas5 is None:
                for pid in ("Kompas.Application.5", "KOMPAS.Application.5"):
                    try:
                        self.kompas5 = self.win32.Dispatch(pid)
                        break
                    except Exception:
                        continue
            if self.kompas5 is None:
                raise RuntimeError("cannot connect to KOMPAS COM")
            try:
                self.kompas5.Visible = True
            except Exception:
                pass
        if self.kompas7 is None:
            self.kompas7 = self._safe_get(self.kompas5, "ksGetApplication7")
            if callable(self.kompas7):
                try:
                    self.kompas7 = self.kompas7()
                except Exception:
                    self.kompas7 = None
        return self.kompas5, self.kompas7

    def _get_active_doc_context(self) -> tuple[Path, Any, Any] | None:
        try:
            app5, app7 = self._connect_kompas()
        except RuntimeError:
            self._wait_once("cannot connect to KOMPAS COM")
            return None
        doc = self._safe_get(app7, "ActiveDocument") or self._safe_get(app5, "ActiveDocument")
        if doc is None:
            self._wait_once("no active KOMPAS document")
            return None
        doc2d = self._cast(doc, "IKompasDocument2D") or doc
        path = self._doc_path(doc2d)
        if path is None:
            self._wait_once("active drawing is not saved")
            return None
        view = self._active_view(doc2d)
        if view is None:
            self._wait_once("active view unavailable")
            return None
        return path, doc2d, view

    def _doc_path(self, doc: Any) -> Path | None:
        for name in ("PathName", "DocumentPath", "FullName", "FilePath", "PathAndName", "Path", "Name"):
            value = normalize_text(self._safe_get(doc, name)).strip().strip('"').replace("/", "\\")
            if not value:
                continue
            p = Path(value)
            if p.is_absolute() and p.suffix:
                return p
        return None

    def _active_view(self, doc2d: Any) -> Any:
        manager = self._safe_get(doc2d, "ViewsAndLayersManager")
        views = self._safe_get(manager, "Views") if manager is not None else None
        return self._safe_get(views, "ActiveView") or self._safe_get(doc2d, "ActiveView")

    def _ensure_excel(self) -> Any:
        self._ensure_pywin32()
        if self.excel is not None:
            return self.excel
        try:
            self.excel = self.win32.GetObject(Class="Excel.Application")
        except Exception:
            try:
                self.excel = self.win32.Dispatch("Excel.Application")
            except Exception as exc:
                raise RuntimeError(f"cannot connect to Excel COM: {exc}") from exc
        try:
            self.excel.Visible = True
            self.excel.DisplayAlerts = False
        except Exception:
            pass
        return self.excel

    def _ensure_workbook(self, workbook_path: Path) -> tuple[Any, Any]:
        app = self._ensure_excel()
        target = os.path.normcase(str(workbook_path.resolve()))
        existing = None
        books = self._safe_get(app, "Workbooks")
        for wb in self._iter_collection(books):
            full_name = normalize_text(self._safe_get(wb, "FullName")).strip()
            if full_name and os.path.normcase(full_name) == target:
                existing = wb
                break
        if existing is None:
            workbook_path.parent.mkdir(parents=True, exist_ok=True)
            if workbook_path.exists():
                existing = app.Workbooks.Open(str(workbook_path))
            else:
                existing = app.Workbooks.Add()
                try:
                    existing.SaveAs(str(workbook_path), 51)
                except Exception:
                    existing.SaveAs(str(workbook_path))
        ws = None
        try:
            ws = existing.Worksheets(self.sheet_name)
        except Exception:
            ws = None
        if ws is None:
            try:
                ws = existing.Worksheets.Add()
                ws.Name = self.sheet_name
            except Exception:
                ws = existing.Worksheets(1)
        try:
            existing.Activate()
            ws.Activate()
        except Exception:
            pass
        return existing, ws

    def _collect_texts(self, doc2d: Any, active_view: Any) -> list[TextRuntime]:
        views = [active_view]
        manager = self._safe_get(doc2d, "ViewsAndLayersManager")
        col = self._safe_get(manager, "Views") if manager is not None else None
        for v in self._iter_collection(col):
            views.append(v)

        unique_views: list[Any] = []
        seen_views: set[str] = set()
        for view in views:
            view_key = self._pointer(view)
            if not view_key:
                view_key = f"obj:{id(view)}"
            if view_key in seen_views:
                continue
            seen_views.add(view_key)
            unique_views.append(view)
        views = unique_views

        # Primary path for KOMPAS API7: IDrawingContainer.DrawingTexts -> IDrawingText -> IText.Str
        from_drawing_texts: list[TextRuntime] = []
        used_refs: set[str] = set()
        for view in views:
            container = self._cast(view, "IDrawingContainer")
            if container is None:
                continue
            drawing_texts = self._safe_get(container, "DrawingTexts")
            if drawing_texts is None:
                continue

            count = self._safe_int(self._safe_get(drawing_texts, "Count"), 0)
            for idx in range(count):
                drawing_text = None
                for raw_idx in (idx, idx + 1):
                    try:
                        drawing_text = drawing_texts.DrawingText(raw_idx)
                        if drawing_text is not None:
                            break
                    except Exception:
                        drawing_text = None
                if drawing_text is None:
                    continue

                text_obj = self._cast(drawing_text, "IText")
                if text_obj is None:
                    continue

                text_value = normalize_text(self._safe_get(text_obj, "Str"))
                ref = self._safe_int(self._safe_get(drawing_text, "Reference"), 0)
                x_val = self._safe_float(self._safe_get(drawing_text, "X"), 0.0)
                y_val = self._safe_float(self._safe_get(drawing_text, "Y"), 0.0)

                if ref > 0:
                    text_id = f"id:{ref}"
                else:
                    text_id = self._extract_id(drawing_text, x_val, y_val, idx)

                if text_id in used_refs:
                    if text_id.startswith("id:"):
                        continue
                    seq = 2
                    probe = f"{text_id}#{seq}"
                    while probe in used_refs:
                        seq += 1
                        probe = f"{text_id}#{seq}"
                    text_id = probe
                used_refs.add(text_id)

                from_drawing_texts.append(
                    TextRuntime(text_id=text_id, text=text_value, x=x_val, y=y_val, item=drawing_text)
                )

        if from_drawing_texts:
            from_drawing_texts.sort(key=lambda i: (-i.y, i.x, i.text_id))
            return from_drawing_texts

        sources: list[Any] = []
        for view in views:
            sources.append(view)
            casted = self._cast(view, "ISymbols2DContainer")
            if casted is not None:
                sources.append(casted)
        sources.append(doc2d)

        out: list[TextRuntime] = []
        used: set[str] = set()
        collections = ("Texts", "TextObjects", "TextLines", "DrawingTexts", "Notes", "Objects", "Symbols")
        for source in sources:
            for cname in collections:
                collection = self._safe_get(source, cname)
                if collection is None:
                    continue
                for idx, item in enumerate(self._iter_collection(collection)):
                    text = self._extract_text(item)
                    if text is None:
                        continue
                    x_val, y_val = self._extract_xy(item)
                    base_id = self._extract_id(item, x_val, y_val, idx)
                    text_id = base_id
                    seq = 2
                    while text_id in used:
                        text_id = f"{base_id}#{seq}"
                        seq += 1
                    used.add(text_id)
                    out.append(TextRuntime(text_id=text_id, text=text, x=x_val, y=y_val, item=item))
        out.sort(key=lambda i: (-i.y, i.x, i.text_id))
        return out

    def _extract_text(self, item: Any) -> str | None:
        for name in ("Str", "Text", "String", "Content", "Value", "Title", "Name"):
            val = self._safe_get(item, name)
            if isinstance(val, (str, int, float)):
                return normalize_text(val)
            for inner in ("Str", "Text", "String", "Value"):
                nested = self._safe_get(val, inner)
                if isinstance(nested, (str, int, float)):
                    return normalize_text(nested)
        return None

    def _extract_xy(self, item: Any) -> tuple[float, float]:
        x_val = self._safe_float(self._first(item, ("X", "PosX", "XPos", "PointX", "CoordX")), 0.0)
        y_val = self._safe_float(self._first(item, ("Y", "PosY", "YPos", "PointY", "CoordY")), 0.0)
        if x_val == 0.0 and y_val == 0.0:
            p = self._safe_get(item, "Point")
            x_val = self._safe_float(self._first(p, ("X", "x")), x_val)
            y_val = self._safe_float(self._first(p, ("Y", "y")), y_val)
        return x_val, y_val

    def _extract_id(self, item: Any, x_val: float, y_val: float, idx: int) -> str:
        raw = self._first(item, ("Id", "ID", "ObjID", "ObjectId", "ReferenceId", "EntityId", "Number"))
        if raw is not None:
            text = normalize_text(raw).strip()
            if text:
                return f"id:{text}"
        ptr = self._pointer(item)
        if ptr:
            return f"ptr:{ptr}"
        return f"xy:{x_val:.3f}:{y_val:.3f}:{idx + 1}"

    def _rebuild_map(self, doc_state: dict[str, Any], elements: list[TextRuntime], by_id: dict[str, TextRuntime], ws: Any, signature: str) -> None:
        snapshots = [TextElement(text_id=e.text_id, text=e.text, x=e.x, y=e.y) for e in elements]
        rows = group_vertical(snapshots, self.corridor_mm)
        layout = build_bindings(rows)

        for b in doc_state.get("bindings", []):
            self._write_cell(ws, int(b.get("row", 1)), int(b.get("col", 1)), "")
            self._set_red(ws, int(b.get("row", 1)), int(b.get("col", 1)), False)

        bindings: list[dict[str, Any]] = []
        for cell in layout:
            text_id = cell.text_id
            item = by_id.get(text_id)
            if item is None:
                continue
            value = normalize_text(item.text)
            self._write_cell(ws, cell.row, cell.col, value)
            self._set_red(ws, cell.row, cell.col, False)
            bindings.append(
                {
                    "row": int(cell.row),
                    "col": int(cell.col),
                    "text_id": text_id,
                    "last_excel": value,
                    "last_kompas": value,
                }
            )
        doc_state["bindings"] = bindings
        doc_state["signature"] = signature
        doc_state["corridor_mm"] = self.corridor_mm
        doc_state["_dirty"] = True
        rows_count = max((int(b["row"]) for b in bindings), default=0)
        log(f"INFO: mapping rebuilt ({len(bindings)} bindings, rows={rows_count}, corridor={self.corridor_mm:g} mm)")

    def _sync_cells(self, doc_state: dict[str, Any], by_id: dict[str, TextRuntime], ws: Any) -> dict[str, bool]:
        excel_changed = False
        kompas_changed = False
        state_changed = False
        e2k = 0
        k2e = 0
        missing = 0
        for b in doc_state.get("bindings", []):
            row = int(b.get("row", 1))
            col = int(b.get("col", 1))
            text_id = str(b.get("text_id", ""))
            item = by_id.get(text_id)
            if item is None:
                missing += 1
                continue
            cur_excel = self._read_cell(ws, row, col)
            cur_kompas = normalize_text(item.text)
            action = choose_sync_action(
                last_excel=normalize_text(b.get("last_excel", "")),
                last_kompas=normalize_text(b.get("last_kompas", "")),
                current_excel=cur_excel,
                current_kompas=cur_kompas,
            )
            if action == "excel_to_kompas":
                if self._set_text(item.item, cur_excel):
                    item.text = cur_excel
                    cur_kompas = cur_excel
                    kompas_changed = True
                    e2k += 1
                    state_changed = True
            elif action == "kompas_to_excel":
                self._write_cell(ws, row, col, cur_kompas)
                cur_excel = cur_kompas
                excel_changed = True
                k2e += 1
                state_changed = True
            prev_last_excel = normalize_text(b.get("last_excel", ""))
            prev_last_kompas = normalize_text(b.get("last_kompas", ""))
            b["last_excel"] = cur_excel
            b["last_kompas"] = cur_kompas
            if prev_last_excel != cur_excel or prev_last_kompas != cur_kompas:
                state_changed = True
        if missing > 0:
            doc_state["signature"] = ""
            state_changed = True
        if e2k > 0 or k2e > 0:
            log(f"INFO: sync updates excel->kompas={e2k}, kompas->excel={k2e}")
        return {"excel_changed": excel_changed, "kompas_changed": kompas_changed, "state_changed": state_changed}

    def _mark_unbound(self, ws: Any, bindings: list[dict[str, Any]]) -> bool:
        bound = {(int(b.get("row", 1)), int(b.get("col", 1))) for b in bindings}
        max_row = max((r for r, _ in bound), default=1)
        max_col = max((c for _, c in bound), default=1)
        used = self._safe_get(ws, "UsedRange")
        used_rows = self._safe_int(self._safe_get(self._safe_get(used, "Rows"), "Count"), 1)
        used_cols = self._safe_int(self._safe_get(self._safe_get(used, "Columns"), "Count"), 1)
        scan_rows = min(max(max_row + 5, used_rows), 1500)
        scan_cols = min(max(max_col + 2, used_cols), 200)
        changed = False
        for row in range(1, scan_rows + 1):
            for col in range(1, scan_cols + 1):
                if (row, col) in bound:
                    changed = self._set_red(ws, row, col, False) or changed
                    continue
                mark = bool(self._read_cell(ws, row, col).strip())
                changed = self._set_red(ws, row, col, mark) or changed
        return changed

    def _set_text(self, item: Any, value: str) -> bool:
        value = normalize_text(value)
        text_obj = self._cast(item, "IText")
        if text_obj is not None:
            try:
                text_obj.Str = value
                self._safe_call(item, ("Update", "Refresh", "Rebuild"))
                return True
            except Exception:
                pass

        for name in ("Str", "Text", "String", "Content", "Value"):
            if self._safe_set(item, name, value):
                self._safe_call(item, ("Update", "Refresh", "Rebuild"))
                return True
        text_obj = self._safe_get(item, "Text")
        for name in ("Str", "Text", "String", "Value"):
            if self._safe_set(text_obj, name, value):
                self._safe_call(text_obj, ("Update", "Refresh", "Rebuild"))
                return True
        for method in ("SetText", "SetStr", "PutText", "SetContent", "SetString"):
            if self._safe_call(item, (method,), args=(value,)) is not None:
                self._safe_call(item, ("Update", "Refresh", "Rebuild"))
                return True
        return False

    def _set_red(self, ws: Any, row: int, col: int, mark: bool) -> bool:
        try:
            cell = ws.Cells(row, col)
            interior = cell.Interior
            cur = self._safe_int(self._safe_get(interior, "ColorIndex"), XL_NONE)
            target = XL_RED if mark else XL_NONE
            if cur == target:
                return False
            interior.ColorIndex = target
            return True
        except Exception:
            return False

    def _read_cell(self, ws: Any, row: int, col: int) -> str:
        try:
            return normalize_text(ws.Cells(row, col).Value)
        except Exception:
            return ""

    def _write_cell(self, ws: Any, row: int, col: int, value: str) -> None:
        try:
            ws.Cells(row, col).Value = value
        except Exception:
            pass

    def _save_workbook(self, wb: Any) -> None:
        try:
            wb.Save()
        except Exception as exc:
            log(f"WARN: workbook save failed: {exc}")

    def _cast(self, obj: Any, target_type: str) -> Any:
        if self.win32 is None or obj is None:
            return None
        try:
            return self.win32.CastTo(obj, target_type)
        except Exception:
            return None

    def _iter_collection(self, collection: Any):
        if collection is None:
            return
        count = self._safe_int(self._first(collection, ("Count", "Length")), -1)
        if count >= 0:
            for index in range(count):
                item = None
                for raw in (index, index + 1):
                    try:
                        item = collection.Item(raw)
                        if item is not None:
                            break
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

    def _pointer(self, obj: Any) -> str:
        ole = getattr(obj, "_oleobj_", None)
        if ole is None:
            return ""
        try:
            return str(int(ole))
        except Exception:
            try:
                return str(hash(ole))
            except Exception:
                return ""

    def _first(self, obj: Any, names: tuple[str, ...]) -> Any:
        for name in names:
            value = self._safe_get(obj, name)
            if value is not None:
                return value
            value = self._safe_call(obj, (name,))
            if value is not None:
                return value
        return None

    @staticmethod
    def _safe_get(obj: Any, name: str) -> Any:
        if obj is None:
            return None
        try:
            return getattr(obj, name)
        except Exception:
            return None

    @staticmethod
    def _safe_set(obj: Any, name: str, value: Any) -> bool:
        if obj is None:
            return False
        try:
            setattr(obj, name, value)
            return True
        except Exception:
            return False

    @staticmethod
    def _safe_call(obj: Any, names: tuple[str, ...], args: tuple[Any, ...] = ()) -> Any:
        if obj is None:
            return None
        for name in names:
            try:
                method = getattr(obj, name)
            except Exception:
                method = None
            if callable(method):
                try:
                    return method(*args)
                except Exception:
                    continue
        return None

    @staticmethod
    def _safe_int(value: Any, default: int) -> int:
        try:
            return int(value)
        except Exception:
            return default

    @staticmethod
    def _safe_float(value: Any, default: float) -> float:
        try:
            return float(str(value).replace(",", "."))
        except Exception:
            return default


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Sync KOMPAS drawing texts with Excel via COM.")
    parser.add_argument("--corridor-mm", type=float, default=1.0, help="Vertical grouping corridor in mm.")
    parser.add_argument("--poll-ms", type=int, default=1200, help="Sync polling period in milliseconds.")
    # Backward compatibility: ignored, state file is now always <workbook>.json.
    parser.add_argument("--state-file", default="", help=argparse.SUPPRESS)
    parser.add_argument("--sheet-name", default=SHEET_DEFAULT, help="Excel worksheet name.")
    parser.add_argument("--log-file", default="", help="Optional log file path.")
    parser.add_argument("--once", action="store_true", help="Execute one tick and exit.")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    global LOG_FILE
    args = parse_args(argv)
    if not args.sheet_name.strip():
        return EXIT_USAGE

    if str(args.log_file).strip():
        try:
            LOG_FILE = Path(args.log_file).expanduser().resolve()
            LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            LOG_FILE = None

    try:
        engine = SyncEngine(
            corridor_mm=args.corridor_mm,
            poll_ms=args.poll_ms,
            sheet_name=args.sheet_name.strip(),
        )
        return engine.run(once=bool(args.once))
    except RuntimeError as exc:
        text = str(exc)
        log(f"ERROR: {text}")
        if "win32com.client" in text:
            return EXIT_PYWIN32_MISSING
        if "KOMPAS COM" in text or "Kompas.Application.5" in text:
            return EXIT_KOMPAS_ERROR
        if "Excel COM" in text:
            return EXIT_EXCEL_ERROR
        return EXIT_USAGE
    except KeyboardInterrupt:
        log("INFO: stopped by user")
        return EXIT_OK
    except Exception as exc:
        log(f"ERROR: unexpected failure: {exc}")
        return EXIT_USAGE


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
