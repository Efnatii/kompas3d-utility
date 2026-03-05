from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from kompas_excel_text_sync import SyncEngine  # noqa: E402


class StrictText:
    def __init__(self) -> None:
        object.__setattr__(self, "Str", "old")
        object.__setattr__(self, "Color", 7)
        object.__setattr__(self, "Height", 3.5)
        object.__setattr__(self, "Style", "GOST")
        object.__setattr__(self, "updates", 0)

    def __setattr__(self, name: str, value: object) -> None:
        if name in {"Color", "Height", "Style"} and hasattr(self, name):
            raise AssertionError(f"Formatting field '{name}' must not be changed")
        object.__setattr__(self, name, value)

    def Update(self) -> None:
        self.updates += 1


class FormatResetText:
    def __init__(self) -> None:
        object.__setattr__(self, "Str", "old")
        object.__setattr__(self, "Color", 12)
        object.__setattr__(self, "Height", 4.2)
        object.__setattr__(self, "Style", "Corp")
        object.__setattr__(self, "Bold", True)
        object.__setattr__(self, "Italic", True)

    def __setattr__(self, name: str, value: object) -> None:
        if name == "Str":
            object.__setattr__(self, name, value)
            # Emulate hostile COM behavior that resets formatting when text changes.
            object.__setattr__(self, "Color", 0)
            object.__setattr__(self, "Height", 1.0)
            object.__setattr__(self, "Style", "Default")
            object.__setattr__(self, "Bold", False)
            object.__setattr__(self, "Italic", False)
            return
        object.__setattr__(self, name, value)


class FormatResetOnUpdateText:
    def __init__(self) -> None:
        object.__setattr__(self, "Str", "old")
        object.__setattr__(self, "Width", 25.0)
        object.__setattr__(self, "Height", 8.0)

    def __setattr__(self, name: str, value: object) -> None:
        object.__setattr__(self, name, value)

    def Update(self) -> None:
        # Emulate a KOMPAS object that mutates size during update call.
        object.__setattr__(self, "Width", 1.0)
        object.__setattr__(self, "Height", 1.0)


class TextStyleState:
    def __init__(self, height: float, width: float) -> None:
        self.Height = height
        self.Width = width


class ReplacingStyleOnStrText:
    def __init__(self) -> None:
        object.__setattr__(self, "Str", "old")
        object.__setattr__(self, "TextStyle", TextStyleState(height=12.0, width=19.0))

    def __setattr__(self, name: str, value: object) -> None:
        if name == "Str":
            object.__setattr__(self, name, value)
            # Emulate API behavior: style object gets replaced with default size=5.
            object.__setattr__(self, "TextStyle", TextStyleState(height=5.0, width=5.0))
            return
        object.__setattr__(self, name, value)


class DelayedApplyText:
    def __init__(self) -> None:
        object.__setattr__(self, "_actual", "old")
        object.__setattr__(self, "_pending", None)
        object.__setattr__(self, "Height", 6.0)
        object.__setattr__(self, "updates", 0)

    @property
    def Str(self) -> str:
        return self._actual

    @Str.setter
    def Str(self, value: str) -> None:
        object.__setattr__(self, "_pending", value)

    def Update(self) -> None:
        self.updates += 1
        if self._pending is not None:
            object.__setattr__(self, "_actual", self._pending)
            object.__setattr__(self, "_pending", None)


class ContainerWithoutStr:
    __slots__ = ("Text",)

    def __init__(self, text_obj: StrictText) -> None:
        self.Text = text_obj


def test_set_text_updates_only_str_direct() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    item = StrictText()
    updater = StrictText()

    assert engine._set_text(item, "NEW", updater) is True
    assert item.Str == "NEW"
    assert item.Color == 7
    assert item.Height == 3.5
    assert item.Style == "GOST"
    assert updater.updates == 0


def test_set_text_updates_only_nested_text_str() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    nested = StrictText()
    item = ContainerWithoutStr(nested)
    updater = StrictText()

    assert engine._set_text(item, "NESTED", updater) is True
    assert nested.Str == "NESTED"
    assert nested.Color == 7
    assert nested.Height == 3.5
    assert nested.Style == "GOST"
    assert updater.updates == 0


def test_set_text_returns_false_when_str_unavailable() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    item = object()
    updater = StrictText()

    assert engine._set_text(item, "VALUE", updater) is False
    assert updater.updates == 0


def test_set_text_restores_appearance_when_str_write_resets_it() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    item = FormatResetText()

    assert engine._set_text(item, "NEW-VALUE", None) is True
    assert item.Str == "NEW-VALUE"
    assert item.Color == 12
    assert item.Height == 4.2
    assert item.Style == "Corp"
    assert item.Bold is True
    assert item.Italic is True


def test_set_text_restores_width_height_when_update_resets_size() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    item = FormatResetOnUpdateText()
    updater = item

    assert engine._set_text(item, "VALUE", updater) is True
    assert item.Str == "VALUE"
    assert item.Width == 25.0
    assert item.Height == 8.0


def test_set_text_restores_size_when_style_object_replaced() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    item = ReplacingStyleOnStrText()

    assert engine._set_text(item, "REPLACED", None) is True
    assert item.Str == "REPLACED"
    assert item.TextStyle.Height == 12.0
    assert item.TextStyle.Width == 19.0


def test_set_text_calls_update_only_when_text_not_applied_without_it() -> None:
    engine = SyncEngine(corridor_mm=1.0, poll_ms=1000, sheet_name="SyncData")
    item = DelayedApplyText()

    assert engine._set_text(item, "APPLIED", item) is True
    assert item.Str == "APPLIED"
    assert item.updates == 1
    assert item.Height == 6.0
