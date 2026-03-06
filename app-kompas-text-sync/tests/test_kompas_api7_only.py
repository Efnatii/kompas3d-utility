from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine  # noqa: E402


class FakeWin32:
    def __init__(self, app: object, fail_get_object: bool = False) -> None:
        self._app = app
        self._fail_get_object = fail_get_object
        self.get_calls: list[str] = []
        self.dispatch_calls: list[str] = []

    def GetObject(self, Class: str) -> object:
        self.get_calls.append(Class)
        if self._fail_get_object:
            raise RuntimeError("GetObject failed")
        if Class != "KOMPAS.Application.7":
            raise RuntimeError("unexpected ProgID")
        return self._app

    def Dispatch(self, prog_id: str) -> object:
        self.dispatch_calls.append(prog_id)
        if prog_id != "KOMPAS.Application.7":
            raise RuntimeError("unexpected ProgID")
        return self._app


class FakeDrawingTextsCollection:
    def __init__(self, items: list[object]) -> None:
        self._items = list(items)
        self.Count = len(self._items)

    def DrawingText(self, index: int) -> object:
        if 1 <= index <= len(self._items):
            return self._items[index - 1]
        raise IndexError(index)


class FakeLegacyCollection:
    def __init__(self, items: list[object]) -> None:
        self._items = list(items)
        self.Count = len(self._items)

    def Item(self, index: int) -> object:
        if 0 <= index < len(self._items):
            return self._items[index]
        raise IndexError(index)


class FakeDrawingText:
    def __init__(self, text: str, x: float, y: float, reference: int) -> None:
        self.Str = text
        self.X = x
        self.Y = y
        self.Reference = reference


class FakeDrawingContainer:
    def __init__(self, drawing_texts: FakeDrawingTextsCollection | None) -> None:
        self.DrawingTexts = drawing_texts


class FakeView:
    def __init__(self, container: FakeDrawingContainer | None, legacy_texts: FakeLegacyCollection | None = None) -> None:
        self._container = container
        self.Texts = legacy_texts


def test_connect_kompas_uses_api7_progid() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    app = SimpleNamespace(Visible=False)
    engine.win32 = FakeWin32(app)

    connected = engine._connect_kompas()

    assert connected is app
    assert app.Visible is True
    assert engine.win32.get_calls == ["KOMPAS.Application.7"]
    assert engine.win32.dispatch_calls == []


def test_connect_kompas_dispatch_after_getobject_failure() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    app = SimpleNamespace(Visible=False)
    engine.win32 = FakeWin32(app, fail_get_object=True)

    connected = engine._connect_kompas()

    assert connected is app
    assert app.Visible is True
    assert engine.win32.get_calls == ["KOMPAS.Application.7"]
    assert engine.win32.dispatch_calls == ["KOMPAS.Application.7"]


def test_collect_texts_reads_only_drawing_texts() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")

    def fake_cast(obj: object, target_type: str) -> object | None:
        if target_type == "IDrawingContainer":
            return getattr(obj, "_container", None)
        if target_type == "IText" and hasattr(obj, "Str"):
            return obj
        return None

    engine._cast = fake_cast  # type: ignore[method-assign]
    view = FakeView(
        container=FakeDrawingContainer(
            FakeDrawingTextsCollection(
                [
                    FakeDrawingText("BOTTOM", x=10.0, y=50.0, reference=12),
                    FakeDrawingText("TOP", x=5.0, y=70.0, reference=34),
                ]
            )
        )
    )
    doc2d = SimpleNamespace(ViewsAndLayersManager=None)

    items = engine._collect_texts(doc2d, view)

    assert [item.text_id for item in items] == ["id:34", "id:12"]
    assert [item.text for item in items] == ["TOP", "BOTTOM"]


def test_collect_texts_without_drawing_texts_returns_empty() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")

    def fake_cast(obj: object, target_type: str) -> object | None:
        if target_type == "IDrawingContainer":
            return getattr(obj, "_container", None)
        if target_type == "IText" and hasattr(obj, "Str"):
            return obj
        return None

    engine._cast = fake_cast  # type: ignore[method-assign]
    legacy = FakeLegacyCollection([SimpleNamespace(Str="legacy", X=1.0, Y=1.0)])
    view = FakeView(container=None, legacy_texts=legacy)
    doc2d = SimpleNamespace(ViewsAndLayersManager=None, Texts=legacy)

    items = engine._collect_texts(doc2d, view)

    assert items == []
