from __future__ import annotations

import sys
from pathlib import Path

import win32com.client


EXIT_OK = 0
EXIT_USAGE = 1
EXIT_KOMPAS_ERROR = 20
EXIT_NO_ACTIVE_DOC = 30
EXIT_NO_ACTIVE_VIEW = 31
EXIT_TABLE_CREATE = 40
EXIT_TABLE_FILL = 41
EXIT_TABLE_SAVE = 50

DEFAULT_ROW_HEIGHT_MM = 8.0
DEFAULT_COL_WIDTH_MM = 30.0
KS_TTL_NOT_CREATE = 0


def log(message: str) -> None:
    print(message)


def read_matrix_tsv(tsv_path: Path, row_count: int, col_count: int) -> list[list[str]]:
    lines = tsv_path.read_text(encoding="utf-16").splitlines()
    matrix: list[list[str]] = []

    for raw_line in lines:
        row = raw_line.split("\t")
        if len(row) < col_count:
            row.extend([""] * (col_count - len(row)))
        matrix.append(row[:col_count])

    while len(matrix) < row_count:
        matrix.append([""] * col_count)

    return matrix[:row_count]


def connect_kompas():
    try:
        kompas5 = win32com.client.GetObject(Class="Kompas.Application.5")
    except Exception:
        kompas5 = win32com.client.Dispatch("Kompas.Application.5")

    kompas5.Visible = True
    kompas7 = kompas5.ksGetApplication7
    return kompas5, kompas7


def get_cell_with_fallback(table, row: int, col: int):
    for r, c in ((row, col), (row + 1, col + 1)):
        try:
            return table.Cell(r, c)
        except Exception:
            continue
    return None


def parse_mm(raw_value: str) -> float:
    return float(raw_value.strip().replace(",", "."))


def main() -> int:
    if len(sys.argv) not in (5, 7):
        log(
            "Использование: python kompas_tbl_bridge.py <matrix.tsv> <out.tbl> <rows> <cols> [<row_height_mm> <col_width_mm>]"
        )
        return EXIT_USAGE

    matrix_path = Path(sys.argv[1]).resolve()
    out_tbl = Path(sys.argv[2]).resolve()
    row_count = int(sys.argv[3])
    col_count = int(sys.argv[4])

    row_height_mm = DEFAULT_ROW_HEIGHT_MM
    col_width_mm = DEFAULT_COL_WIDTH_MM

    if len(sys.argv) == 7:
        try:
            row_height_mm = parse_mm(sys.argv[5])
            col_width_mm = parse_mm(sys.argv[6])
        except Exception as exc:
            log(f"ERROR: Некорректные размеры ячеек для bridge: {exc}")
            return EXIT_USAGE

    log(f"INFO: Python bridge получил матрицу {row_count}x{col_count}: {matrix_path}")
    log(f"INFO: Размер ячейки (мм): width={col_width_mm}, height={row_height_mm}")
    matrix = read_matrix_tsv(matrix_path, row_count, col_count)

    try:
        _, kompas7 = connect_kompas()
        log("INFO: Python bridge подключился к КОМПАС.")
    except Exception as exc:
        log(f"ERROR: Python bridge не подключился к КОМПАС: {exc}")
        return EXIT_KOMPAS_ERROR

    doc = getattr(kompas7, "ActiveDocument", None)
    if doc is None:
        log("ERROR: Нет активного документа КОМПАС. Откройте Фрагмент/Чертёж.")
        return EXIT_NO_ACTIVE_DOC

    try:
        doc2d = win32com.client.CastTo(doc, "IKompasDocument2D")
    except Exception as exc:
        log(f"ERROR: Активный документ нельзя привести к IKompasDocument2D: {exc}")
        return EXIT_NO_ACTIVE_DOC

    try:
        active_view = doc2d.ViewsAndLayersManager.Views.ActiveView
    except Exception as exc:
        log(f"ERROR: Не удалось получить ActiveView: {exc}")
        return EXIT_NO_ACTIVE_VIEW

    if active_view is None:
        log("ERROR: ActiveView отсутствует.")
        return EXIT_NO_ACTIVE_VIEW

    try:
        symbols_container = win32com.client.CastTo(active_view, "ISymbols2DContainer")
        drawing_tables = symbols_container.DrawingTables
    except Exception as exc:
        log(f"ERROR: Не удалось получить ISymbols2DContainer.DrawingTables: {exc}")
        return EXIT_TABLE_CREATE

    try:
        drawing_table = drawing_tables.Add(
            row_count,
            col_count,
            row_height_mm,
            col_width_mm,
            KS_TTL_NOT_CREATE,
        )
    except Exception as exc:
        log(f"ERROR: Ошибка DrawingTables.Add(...): {exc}")
        return EXIT_TABLE_CREATE

    try:
        table = win32com.client.CastTo(drawing_table, "ITable")
    except Exception as exc:
        log(f"ERROR: Не удалось привести IDrawingTable к ITable: {exc}")
        return EXIT_TABLE_CREATE

    for r in range(row_count):
        for c in range(col_count):
            cell = get_cell_with_fallback(table, r, c)
            if cell is None:
                log(f"ERROR: Не удалось получить ячейку ({r}, {c}).")
                return EXIT_TABLE_FILL

            try:
                text_obj = win32com.client.CastTo(cell.Text, "IText")
                text_obj.Str = matrix[r][c]
            except Exception as exc:
                log(f"ERROR: Ошибка записи ячейки ({r + 1}, {c + 1}): {exc}")
                return EXIT_TABLE_FILL

    try:
        drawing_table.Save(str(out_tbl))
    except Exception as exc:
        log(f"ERROR: Ошибка сохранения таблицы в {out_tbl}: {exc}")
        return EXIT_TABLE_SAVE

    log(f"OK: Python bridge сохранил {out_tbl}")
    return EXIT_OK


if __name__ == "__main__":
    raise SystemExit(main())
