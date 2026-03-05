from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from sync_logic import (  # noqa: E402
    TextElement,
    build_bindings,
    build_signature,
    choose_sync_action,
    group_vertical,
)


def test_group_vertical_and_sorting() -> None:
    elements = [
        TextElement(text_id="t1", text="A", x=30.0, y=100.0),
        TextElement(text_id="t2", text="B", x=10.0, y=100.4),
        TextElement(text_id="t3", text="C", x=15.0, y=89.9),
        TextElement(text_id="t4", text="D", x=5.0, y=89.2),
    ]

    rows = group_vertical(elements, corridor_mm=1.0)
    assert len(rows) == 2
    assert [item.text_id for item in rows[0]] == ["t2", "t1"]
    assert [item.text_id for item in rows[1]] == ["t4", "t3"]


def test_build_bindings() -> None:
    rows = [
        [TextElement("a", "A", 1, 10), TextElement("b", "B", 2, 10)],
        [TextElement("c", "C", 1, 5)],
    ]
    bindings = build_bindings(rows)
    assert [(b.row, b.col, b.text_id) for b in bindings] == [
        (1, 1, "a"),
        (1, 2, "b"),
        (2, 1, "c"),
    ]


def test_choose_sync_action_priority() -> None:
    assert choose_sync_action(
        last_excel="X", last_kompas="X", current_excel="Y", current_kompas="X"
    ) == "excel_to_kompas"
    assert choose_sync_action(
        last_excel="X", last_kompas="X", current_excel="X", current_kompas="Y"
    ) == "kompas_to_excel"
    assert choose_sync_action(
        last_excel="X", last_kompas="X", current_excel="Y", current_kompas="Y"
    ) == "state_only"
    assert choose_sync_action(
        last_excel="X", last_kompas="X", current_excel="Y", current_kompas="Z"
    ) == "excel_to_kompas"
    assert choose_sync_action(
        last_excel="X", last_kompas="X", current_excel="X", current_kompas="X"
    ) == "none"


def test_signature_ignores_text_changes() -> None:
    left = [
        TextElement("t1", "A", 1.0, 2.0),
        TextElement("t2", "B", 3.0, 4.0),
    ]
    right = [
        TextElement("t1", "changed", 1.0, 2.0),
        TextElement("t2", "new", 3.0, 4.0),
    ]
    assert build_signature(left) == build_signature(right)
