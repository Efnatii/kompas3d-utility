from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha1
from typing import Iterable, Literal


@dataclass(frozen=True)
class TextElement:
    text_id: str
    text: str
    x: float
    y: float


@dataclass(frozen=True)
class CellBinding:
    row: int
    col: int
    text_id: str


SyncAction = Literal["none", "excel_to_kompas", "kompas_to_excel", "state_only"]


def normalize_text(value: object | None) -> str:
    if value is None:
        return ""
    text = str(value)
    if text == "None":
        return ""
    return text


def group_vertical(elements: Iterable[TextElement], corridor_mm: float) -> list[list[TextElement]]:
    corridor = abs(float(corridor_mm))
    sorted_elements = sorted(elements, key=lambda item: (-item.y, item.x, item.text_id))

    groups: list[dict[str, object]] = []
    for element in sorted_elements:
        target_group: dict[str, object] | None = None
        for group in groups:
            y_ref = float(group["y_ref"])
            if abs(element.y - y_ref) <= corridor:
                target_group = group
                break

        if target_group is None:
            target_group = {"y_values": [element.y], "y_ref": element.y, "items": [element]}
            groups.append(target_group)
            continue

        y_values = target_group["y_values"]
        items = target_group["items"]
        assert isinstance(y_values, list)
        assert isinstance(items, list)
        y_values.append(element.y)
        items.append(element)
        target_group["y_ref"] = sum(y_values) / len(y_values)

    groups.sort(key=lambda group: float(group["y_ref"]), reverse=True)

    output: list[list[TextElement]] = []
    for group in groups:
        items = group["items"]
        assert isinstance(items, list)
        row = sorted(items, key=lambda item: (item.x, item.text_id))
        output.append(row)

    return output


def build_bindings(groups: Iterable[Iterable[TextElement]]) -> list[CellBinding]:
    bindings: list[CellBinding] = []
    row_index = 1
    for row in groups:
        col_index = 1
        for element in row:
            bindings.append(CellBinding(row=row_index, col=col_index, text_id=element.text_id))
            col_index += 1
        row_index += 1
    return bindings


def build_signature(elements: Iterable[TextElement]) -> str:
    normalized = [
        f"{item.text_id}|{item.x:.4f}|{item.y:.4f}"
        for item in sorted(elements, key=lambda v: (v.text_id, v.x, v.y))
    ]
    payload = "\n".join(normalized).encode("utf-8", errors="ignore")
    return sha1(payload).hexdigest()


def choose_sync_action(
    *,
    last_excel: str,
    last_kompas: str,
    current_excel: str,
    current_kompas: str,
) -> SyncAction:
    excel_old = normalize_text(last_excel)
    kompas_old = normalize_text(last_kompas)
    excel_new = normalize_text(current_excel)
    kompas_new = normalize_text(current_kompas)

    excel_changed = excel_new != excel_old
    kompas_changed = kompas_new != kompas_old

    if excel_changed and not kompas_changed:
        return "excel_to_kompas"

    if kompas_changed and not excel_changed:
        return "kompas_to_excel"

    if excel_changed and kompas_changed:
        if excel_new == kompas_new:
            return "state_only"
        return "excel_to_kompas"

    return "none"
