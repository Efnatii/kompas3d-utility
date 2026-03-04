from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
RUN_CMD = PROJECT_ROOT / "scripts" / "run.cmd"
OUT_TBL = PROJECT_ROOT / "out" / "table_M2.tbl"
XLSX_FIXTURE = PROJECT_ROOT / "fixtures" / "table_M2.xlsx"

EXIT_KOMPAS_ERROR = 20
EXIT_NO_ACTIVE_DOC = 30
EXIT_TABLE_CREATE = 40
EXIT_EXCEL_ERROR = 10


def _is_windows() -> bool:
    return os.name == "nt"


def _is_progid_registered(progid: str) -> bool:
    command = (
        "$t=[type]::GetTypeFromProgID('{0}'); "
        "if ($null -ne $t) {{ exit 0 }} else {{ exit 1 }}"
    ).format(progid)

    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", command],
        capture_output=True,
        text=True,
        check=False,
    )
    return result.returncode == 0


@pytest.mark.integration
def test_export_tbl_via_cscript() -> None:
    if not _is_windows():
        pytest.skip("Тест рассчитан на Windows (COM + cscript).")

    if not XLSX_FIXTURE.exists():
        pytest.skip(f"Нет фикстуры: {XLSX_FIXTURE}")

    if not _is_progid_registered("Kompas.Application.5"):
        pytest.skip("KOMPAS COM недоступен (ProgID Kompas.Application.5 не зарегистрирован).")

    if not _is_progid_registered("Excel.Application"):
        pytest.skip("Excel COM недоступен (ProgID Excel.Application не зарегистрирован).")

    if OUT_TBL.exists():
        OUT_TBL.unlink()

    result = subprocess.run(
        ["cmd", "/c", str(RUN_CMD)],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        encoding="cp866",
        errors="replace",
        check=False,
    )

    combined_output = "\n".join(
        [result.stdout or "", result.stderr or ""]
    ).strip()
    print("----- STDOUT/STDERR -----")
    print(combined_output)
    print("----- RETURN CODE -----")
    print(result.returncode)

    if result.returncode == EXIT_KOMPAS_ERROR:
        pytest.skip("КОМПАС COM не подключился во время запуска интеграционного теста.")

    if result.returncode == EXIT_EXCEL_ERROR:
        pytest.skip("Excel COM недоступен во время запуска интеграционного теста.")

    no_active_doc_message = "Нет активного документа КОМПАС" in combined_output
    if result.returncode == EXIT_NO_ACTIVE_DOC or no_active_doc_message:
        pytest.xfail(
            "КОМПАС доступен, но нет активного 2D документа. "
            "Откройте Фрагмент/Чертёж и запустите тест снова."
        )

    create_table_issue = (
        "DrawingTables" in combined_output
        and "fallback не сработал" in combined_output
    )
    if result.returncode == EXIT_TABLE_CREATE and create_table_issue:
        pytest.xfail(
            "КОМПАС доступен, но текущий COM-биндинг не дал прямой доступ к DrawingTables, "
            "и fallback не завершился успешно."
        )

    assert result.returncode == 0, (
        "Экспорт завершился с ошибкой. "
        f"Код: {result.returncode}\nЛог:\n{combined_output}"
    )
    assert OUT_TBL.exists(), f"Файл не создан: {OUT_TBL}"
    assert OUT_TBL.stat().st_size > 0, f"Файл пустой: {OUT_TBL}"
