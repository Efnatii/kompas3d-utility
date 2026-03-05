from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine  # noqa: E402


class FakeWorkbook:
    def __init__(self, full_name: str) -> None:
        self.FullName = full_name
        self.saved = False
        self.closed = False
        self.close_args: tuple[object, ...] = ()

    def Save(self) -> None:
        self.saved = True

    def Close(self, *args: object) -> None:
        self.closed = True
        self.close_args = args


class FakeWorkbooks:
    def __init__(self, items: list[FakeWorkbook]) -> None:
        self._items = items

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> FakeWorkbook:
        if 0 <= index < len(self._items):
            return self._items[index]
        raise IndexError(index)


class FakeExcel:
    def __init__(self, books: list[FakeWorkbook]) -> None:
        self.Workbooks = FakeWorkbooks(books)


class FakeSheet:
    def __init__(self, name: str = "SyncData", index: int = 1) -> None:
        self.Name = name
        self.Index = index


def test_close_previous_workbook_on_active_doc_switch() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    old_path = Path(r"C:\tmp\old.xlsx")
    new_path = Path(r"C:\tmp\new.xlsx")
    wb_old = FakeWorkbook(str(old_path))
    engine.excel = FakeExcel([wb_old])
    engine.last_workbook_path = engine._normalize_path(old_path)

    engine._close_previous_workbook_if_needed(new_path)

    assert wb_old.saved is True
    assert wb_old.closed is True
    assert wb_old.close_args in {(), (False,)}
    assert engine.last_workbook_path == ""


def test_do_not_close_when_workbook_path_unchanged() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    same_path = Path(r"C:\tmp\same.xlsx")
    wb = FakeWorkbook(str(same_path))
    engine.excel = FakeExcel([wb])
    engine.last_workbook_path = engine._normalize_path(same_path)

    engine._close_previous_workbook_if_needed(same_path)

    assert wb.saved is False
    assert wb.closed is False
    assert engine.last_workbook_path == engine._normalize_path(same_path)


def test_close_previous_is_noop_when_excel_instance_missing() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    engine.excel = None
    engine.last_workbook_path = engine._normalize_path(Path(r"C:\tmp\orphan.xlsx"))

    engine._close_previous_workbook_if_needed(Path(r"C:\tmp\next.xlsx"))

    assert engine.last_workbook_path == ""


def test_close_previous_workbook_uses_cached_reference_without_workbooks_collection() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    old_path = Path(r"C:\tmp\cached-old.xlsx")
    new_path = Path(r"C:\tmp\cached-new.xlsx")
    wb_old = FakeWorkbook(str(old_path))
    engine.excel = None
    engine.last_workbook_path = engine._normalize_path(old_path)
    engine.active_workbook_path = engine._normalize_path(old_path)
    engine.active_workbook = wb_old
    engine.active_sheet = FakeSheet()

    engine._close_previous_workbook_if_needed(new_path)

    assert wb_old.saved is True
    assert wb_old.closed is True
    assert engine.active_workbook is None
    assert engine.active_sheet is None
    assert engine.active_workbook_path == ""
    assert engine.last_workbook_path == ""


def test_workbook_and_sheet_alive_helpers() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    wb = FakeWorkbook(r"C:\tmp\alive.xlsx")
    ws = FakeSheet()

    assert engine._is_workbook_alive(wb) is True
    assert engine._is_sheet_alive(ws) is True
    assert engine._is_workbook_alive(None) is False
    assert engine._is_sheet_alive(None) is False
