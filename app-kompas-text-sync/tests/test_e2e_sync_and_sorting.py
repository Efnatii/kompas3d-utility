from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine, normalize_text  # noqa: E402


class FakeText:
    def __init__(self, value: str) -> None:
        self.Str = value


class FakeDrawingText:
    def __init__(self, reference: int, x: float, y: float, text: str) -> None:
        self.Reference = reference
        self.X = x
        self.Y = y
        self.Text = FakeText(text)


class FakeDrawingTextsOneBased:
    def __init__(self, items: list[FakeDrawingText]) -> None:
        self._items = list(items)
        self.Count = len(self._items)

    def DrawingText(self, index: int) -> FakeDrawingText:
        if 1 <= index <= len(self._items):
            return self._items[index - 1]
        raise IndexError(index)


class InMemoryE2EEngine(SyncEngine):
    def __init__(self, *, doc_path: Path, doc2d: Any, active_view: Any) -> None:
        super().__init__(corridor_mm=1.0, poll_ms=0, sheet_name="SyncData")
        self._doc_path = doc_path
        self._doc2d = doc2d
        self._active_view = active_view
        self._doc_state: dict[str, Any] | None = None
        self.ws: dict[tuple[int, int], str] = {}
        self.saved_workbooks = 0
        self.kompas_refresh_calls = 0

    def _cast(self, obj: Any, target_type: str) -> Any:
        if target_type == "IKompasDocument2D":
            return obj
        if target_type == "IDrawingContainer":
            return getattr(obj, "_container", None)
        if target_type == "IText":
            if hasattr(obj, "Str"):
                return obj
            text = getattr(obj, "Text", None)
            if text is not None and hasattr(text, "Str"):
                return text
        return None

    def _get_active_doc_context(self) -> tuple[Path, Any, Any] | None:
        return self._doc_path, self._doc2d, self._active_view

    def _ensure_workbook(self, workbook_path: Path) -> tuple[Any, Any]:  # noqa: ARG002
        return object(), self.ws

    def _close_previous_workbook_if_needed(self, workbook_path: Path) -> None:  # noqa: ARG002
        return

    def _get_doc_state(self, doc_path: Path, workbook_path: Path) -> tuple[dict[str, Any], Path]:
        if self._doc_state is None:
            self._doc_state = self._default_doc_state(doc_path, workbook_path)
            self._doc_state["_dirty"] = False
        return self._doc_state, workbook_path.with_suffix(".json")

    def _save_doc_state_if_needed(self, doc_state: dict[str, Any], state_file: Path, force: bool) -> None:  # noqa: ARG002
        doc_state["_dirty"] = False

    def _try_read_cell(self, ws: dict[tuple[int, int], str], row: int, col: int) -> tuple[str, bool]:
        return normalize_text(ws.get((row, col), "")), True

    def _read_cell(self, ws: dict[tuple[int, int], str], row: int, col: int) -> str:
        return normalize_text(ws.get((row, col), ""))

    def _write_cell(self, ws: dict[tuple[int, int], str], row: int, col: int, value: str) -> bool:
        ws[(row, col)] = normalize_text(value)
        return True

    def _set_red(self, ws: dict[tuple[int, int], str], row: int, col: int, mark: bool) -> bool:  # noqa: ARG002
        return False

    def _mark_unbound(self, doc_key: str, ws: Any, bindings: list[dict[str, Any]], force: bool = False) -> bool:  # noqa: ARG002
        return False

    def _auto_fit_bound_cells(self, ws: Any, bindings: list[dict[str, Any]], force: bool = False) -> bool:  # noqa: ARG002
        return False

    def _save_workbook(self, wb: Any) -> None:  # noqa: ARG002
        self.saved_workbooks += 1

    def _refresh_kompas_after_text_sync(self, active_view: Any, doc2d: Any) -> None:  # noqa: ARG002
        self.kompas_refresh_calls += 1

    def _is_workbook_alive(self, wb: Any) -> bool:
        return wb is not None

    def _is_sheet_alive(self, ws: Any) -> bool:
        return ws is not None


def test_e2e_sync_and_sorting_with_api7_drawingtext_accessor() -> None:
    dt_b = FakeDrawingText(reference=101, x=20.0, y=100.0, text="B")
    dt_a = FakeDrawingText(reference=102, x=10.0, y=100.2, text="A")
    dt_c = FakeDrawingText(reference=103, x=5.0, y=90.0, text="C")
    drawing_texts = FakeDrawingTextsOneBased([dt_b, dt_a, dt_c])

    view = SimpleNamespace()
    view._container = SimpleNamespace(DrawingTexts=drawing_texts)
    doc2d = SimpleNamespace(ViewsAndLayersManager=None)
    doc_path = Path(r"C:\tmp\demo.cdw")
    engine = InMemoryE2EEngine(doc_path=doc_path, doc2d=doc2d, active_view=view)

    first = engine.tick()
    assert first["state"] in {"syncing", "monitoring"}
    assert first["activity"] is True

    assert engine.ws[(1, 1)] == "A"
    assert engine.ws[(1, 2)] == "B"
    assert engine.ws[(2, 1)] == "C"
    assert engine.saved_workbooks == 1

    bindings = [
        (int(entry["row"]), int(entry["col"]), str(entry["text_id"]))
        for entry in engine._doc_state["bindings"]  # type: ignore[index]
    ]
    assert bindings == [(1, 1, "id:102"), (1, 2, "id:101"), (2, 1, "id:103")]

    engine.ws[(1, 2)] = "B_from_excel"
    second = engine.tick()
    assert second["state"] == "syncing"
    assert dt_b.Text.Str == "B_from_excel"
    assert engine.kompas_refresh_calls == 1

    dt_c.Text.Str = "C_from_kompas"
    third = engine.tick()
    assert third["state"] == "syncing"
    assert engine.ws[(2, 1)] == "C_from_kompas"
