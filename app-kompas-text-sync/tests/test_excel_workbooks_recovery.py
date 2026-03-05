from __future__ import annotations

import sys
from pathlib import Path

import pytest


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine  # noqa: E402


class FakeBooks:
    def __init__(self) -> None:
        self.open_calls: list[tuple[tuple[object, ...], dict[str, object]]] = []

    def Open(self, *args: object, **kwargs: object) -> object:
        self.open_calls.append((args, kwargs))
        return {"opened": True}


class FlakyExcelApp:
    def __init__(self, fail_times: int, books: FakeBooks) -> None:
        self._fail_times = fail_times
        self._calls = 0
        self._books = books

    @property
    def Workbooks(self) -> FakeBooks:
        self._calls += 1
        if self._calls <= self._fail_times:
            raise AttributeError("Excel.Application.Workbooks")
        return self._books


class BusyThenBooksApp:
    def __init__(self, busy_times: int, books: FakeBooks) -> None:
        self._busy_times = busy_times
        self._calls = 0
        self._books = books

    @property
    def Workbooks(self) -> FakeBooks:
        self._calls += 1
        if self._calls <= self._busy_times:
            raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")
        return self._books


class NoBooksEngine(SyncEngine):
    def _get_excel_workbooks(self, app: object, retries: int = 1, delay_sec: float = 0.0) -> object | None:  # noqa: ARG002
        return None


class AliveWorkbook:
    def __init__(self, full_name: str) -> None:
        self.FullName = full_name
        self.ReadOnly = False


class AliveSheet:
    def __init__(self, name: str = "SyncData") -> None:
        self.Name = name


class MissingBooksButAliveApp:
    def __init__(self) -> None:
        self.Visible = True
        self.DisplayAlerts = False
        self.Version = "16.0"

    @property
    def Workbooks(self) -> object:
        raise AttributeError("Excel.Application.Workbooks")


class FakeWin32Connector:
    def __init__(self, app: object) -> None:
        self.app = app
        self.get_calls = 0
        self.dispatch_calls = 0

    def GetObject(self, Class: str) -> object:  # noqa: N803
        self.get_calls += 1
        return self.app

    def Dispatch(self, prog_id: str) -> object:
        self.dispatch_calls += 1
        return self.app


class ActiveWorkbookOnly:
    def __init__(self, full_name: str, sheet: AliveSheet) -> None:
        self.FullName = full_name
        self.ReadOnly = False
        self._sheet = sheet

    def Worksheets(self, index_or_name: object) -> AliveSheet:
        if index_or_name in {"SyncData", 1}:
            return self._sheet
        raise KeyError(index_or_name)


class ActiveWorkbookApp:
    def __init__(self, workbook: ActiveWorkbookOnly) -> None:
        self.ActiveWorkbook = workbook
        self.Visible = True
        self.DisplayAlerts = False
        self.Version = "16.0"


class FlakySaveWorkbook:
    def __init__(self, busy_times: int) -> None:
        self._busy_times = busy_times
        self.calls = 0
        self.saved = False

    def Save(self) -> None:
        self.calls += 1
        if self.calls <= self._busy_times:
            raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")
        self.saved = True


class BusyAliveWorkbook:
    @property
    def FullName(self) -> str:
        raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")

    @property
    def Saved(self) -> bool:
        raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")


class BusyAliveSheet:
    @property
    def Name(self) -> str:
        raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")

    @property
    def Index(self) -> int:
        raise RuntimeError("(-2147418111, 'Call was rejected by callee.', None, None)")


def test_get_excel_workbooks_recovers_after_transient_missing() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    books = FakeBooks()
    app = FlakyExcelApp(fail_times=2, books=books)

    found = engine._get_excel_workbooks(app, retries=4, delay_sec=0.0)

    assert found is books


def test_get_excel_workbooks_recovers_after_busy_rejections() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    books = FakeBooks()
    app = BusyThenBooksApp(busy_times=2, books=books)

    found = engine._get_excel_workbooks(app, retries=5, delay_sec=0.0)

    assert found is books


def test_open_workbook_read_write_uses_books_collection() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    books = FakeBooks()
    app = FlakyExcelApp(fail_times=0, books=books)

    opened = engine._open_workbook_read_write(app, Path(r"C:\tmp\file.xlsx"))

    assert opened == {"opened": True}
    assert len(books.open_calls) == 1
    _, kwargs = books.open_calls[0]
    assert kwargs.get("Filename") == r"C:\tmp\file.xlsx"


def test_open_workbook_read_write_raises_when_books_unavailable() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    books = FakeBooks()
    app = FlakyExcelApp(fail_times=10, books=books)

    with pytest.raises(RuntimeError, match="Workbooks"):
        engine._open_workbook_read_write(app, Path(r"C:\tmp\file.xlsx"))


def test_ensure_workbook_uses_cached_handles_when_books_temporarily_unavailable() -> None:
    engine = NoBooksEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    target = Path(r"C:\tmp\cached.xlsx")
    wb = AliveWorkbook(str(target))
    ws = AliveSheet()
    engine.excel = object()
    engine.active_workbook = wb
    engine.active_sheet = ws
    engine.active_workbook_path = engine._normalize_path(target)

    result_wb, result_ws = engine._ensure_workbook(target)

    assert result_wb is wb
    assert result_ws is ws


def test_ensure_excel_keeps_app_when_workbooks_temporarily_unavailable() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    app = MissingBooksButAliveApp()
    win32 = FakeWin32Connector(app)
    engine.win32 = win32

    result = engine._ensure_excel()

    assert result is app
    assert engine.excel is app
    assert win32.get_calls == 1
    assert win32.dispatch_calls == 0


def test_ensure_workbook_uses_active_workbook_fallback_when_books_unavailable() -> None:
    engine = NoBooksEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    target = Path(r"C:\tmp\active.xlsx")
    sheet = AliveSheet()
    workbook = ActiveWorkbookOnly(str(target), sheet)
    engine.excel = ActiveWorkbookApp(workbook)

    result_wb, result_ws = engine._ensure_workbook(target)

    assert result_wb is workbook
    assert result_ws is sheet


def test_save_workbook_retries_on_busy_excel() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    wb = FlakySaveWorkbook(busy_times=2)

    engine._save_workbook(wb)

    assert wb.saved is True
    assert wb.calls == 3


def test_workbook_and_sheet_alive_keep_true_on_busy_excel() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")

    assert engine._is_workbook_alive(BusyAliveWorkbook()) is True
    assert engine._is_sheet_alive(BusyAliveSheet()) is True
