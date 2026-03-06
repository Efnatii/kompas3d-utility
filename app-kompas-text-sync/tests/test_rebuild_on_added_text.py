from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine, TextRuntime  # noqa: E402


class DictWorksheetEngine(SyncEngine):
    def _read_cell(self, ws: dict[tuple[int, int], str], row: int, col: int) -> str:
        return str(ws.get((row, col), ""))

    def _write_cell(self, ws: dict[tuple[int, int], str], row: int, col: int, value: str) -> bool:
        ws[(row, col)] = value
        return True

    def _set_red(self, ws: dict[tuple[int, int], str], row: int, col: int, mark: bool) -> bool:  # noqa: ARG002
        return False


def test_full_excel_rebuild_when_new_text_elements_added() -> None:
    engine = DictWorksheetEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    ws: dict[tuple[int, int], str] = {
        (1, 1): "A_excel",
        (1, 2): "B_excel",
        (1, 3): "manual_note",
    }
    doc_state = {
        "signature": "old-signature",
        "corridor_mm": 1.0,
        "bindings": [
            {"row": 1, "col": 1, "text_id": "A", "last_excel": "A_excel", "last_kompas": "A_kompas"},
            {"row": 1, "col": 2, "text_id": "B", "last_excel": "B_excel", "last_kompas": "B_kompas"},
        ],
        "_dirty": False,
    }

    elements = [
        TextRuntime(text_id="A", text="A_kompas", x=10.0, y=100.0, item=None),
        TextRuntime(text_id="C", text="C_kompas", x=20.0, y=100.0, item=None),
        TextRuntime(text_id="B", text="B_kompas", x=10.0, y=90.0, item=None),
    ]
    by_id = {item.text_id: item for item in elements}

    result = engine._rebuild_map(doc_state, elements, by_id, ws, signature="new-signature")

    assert result["full_rebuild"] is True
    assert ws[(1, 1)] == "A_excel"
    assert ws[(1, 2)] == "C_kompas"
    assert ws[(2, 1)] == "B_excel"
    assert ws[(1, 3)] == "manual_note"

    rebuilt = {(int(item["row"]), int(item["col"]), str(item["text_id"])) for item in doc_state["bindings"]}
    assert rebuilt == {(1, 1, "A"), (1, 2, "C"), (2, 1, "B")}
    assert doc_state["signature"] == "new-signature"
