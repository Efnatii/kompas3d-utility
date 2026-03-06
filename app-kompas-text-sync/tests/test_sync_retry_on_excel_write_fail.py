from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine, TextRuntime  # noqa: E402


class FlakyKompasToExcelEngine(SyncEngine):
    def __init__(self) -> None:
        super().__init__(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
        self.write_calls = 0

    def _try_read_cell(self, ws: dict[tuple[int, int], str], row: int, col: int) -> tuple[str, bool]:
        return str(ws.get((row, col), "")), True

    def _write_cell(self, ws: dict[tuple[int, int], str], row: int, col: int, value: str) -> bool:
        self.write_calls += 1
        if self.write_calls == 1:
            return False
        ws[(row, col)] = value
        return True


def test_kompas_to_excel_retries_after_transient_write_failure() -> None:
    engine = FlakyKompasToExcelEngine()
    ws: dict[tuple[int, int], str] = {(1, 1): "E0"}
    doc_state = {
        "bindings": [
            {"row": 1, "col": 1, "text_id": "T1", "last_excel": "E0", "last_kompas": "K0"},
        ]
    }
    by_id = {
        "T1": TextRuntime(text_id="T1", text="K1", x=0.0, y=0.0, item=object()),
    }

    first = engine._sync_cells(doc_state, by_id, ws)
    binding = doc_state["bindings"][0]
    assert first["kompas_to_excel"] == 0
    assert ws[(1, 1)] == "E0"
    assert binding["last_excel"] == "E0"
    assert binding["last_kompas"] == "K0"

    second = engine._sync_cells(doc_state, by_id, ws)
    assert second["kompas_to_excel"] == 1
    assert ws[(1, 1)] == "K1"
    assert binding["last_excel"] == "K1"
    assert binding["last_kompas"] == "K1"
