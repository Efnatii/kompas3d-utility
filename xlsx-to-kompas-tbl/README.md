# xlsx-to-kompas-tbl

Утилита экспорта `Excel .xlsx -> KOMPAS-3D .tbl` через COM/API (ориентир: KOMPAS-3D v24).

Вход: первый лист `UsedRange` файла `.xlsx`.
Выход: `.tbl`, созданный в активном 2D документе КОМПАС (Фрагмент/Чертёж) и сохранённый в файл.

## Структура

- `src/create_tbl.vbs` - главный экспортёр (`cscript`).
- `src/kompas_tbl_bridge.py` - fallback-мост через `pywin32`, если в VBS недоступен `ActiveView.DrawingTables`.
- `config/table_layout.ini` - конфигурация размеров таблицы/ячеек.
- `fixtures/table_M2.xlsx` - тестовый фикстур (8x13).
- `fixtures/table_M2_expected_matrix.json` - ожидаемая матрица фикстуры.
- `scripts/run.cmd` - основной запуск экспорта.
- `scripts/selfcheck.ps1` - проверка окружения и кодировки VBS.
- `tests/test_xlsx_parse.py` - unit/smoke без КОМПАС.
- `tests/test_integration_kompas.py` - интеграционный тест с `cscript`.

## Быстрый старт

```bat
cd xlsx-to-kompas-tbl
python -m pip install -r requirements.txt
```

Проверка окружения:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\selfcheck.ps1
```

## Конфигурация размеров таблицы

Файл: `config/table_layout.ini`.

Поддерживаются два режима:

1. `mode=cell` - задаёте размеры ячейки напрямую.
2. `mode=table` - задаёте общие размеры таблицы, размеры ячейки вычисляются автоматически:
   - `cell_width_mm = table_width_mm / cols`
   - `cell_height_mm = table_height_mm / rows`

Пример:

```ini
mode=cell
cell_width_mm=30
cell_height_mm=8

; mode=table
; table_width_mm=390
; table_height_mm=64
```

CLI для VBS:

```bat
cscript //nologo src\create_tbl.vbs "fixtures\table_M2.xlsx" "out\table_M2.tbl" "config\table_layout.ini"
```

Третий параметр (путь к `.ini`) опционален.
По умолчанию используется `config/table_layout.ini`.
Если файл не найден, скрипт берёт безопасные значения по умолчанию: `30 x 8 мм`.

## Запуск утилиты

По умолчанию:

```bat
scripts\run.cmd
```

Кастомные пути:

```bat
scripts\run.cmd "C:\path\input.xlsx" "C:\path\out.tbl" "C:\path\table_layout.ini"
```

Ожидаемый результат:

- файл `out\table_M2.tbl` существует;
- размер файла больше 0.

## Тестирование

Одна команда для всего набора:

```bat
python -m pytest -v
```

Только smoke/unit:

```bat
python -m pytest -v tests\test_xlsx_parse.py
```

Только интеграционный:

```bat
python -m pytest -v tests\test_integration_kompas.py
```

Поведение интеграционного теста:

- `SKIP`, если COM ProgID KOMPAS/Excel недоступен.
- `XFAIL`, если КОМПАС доступен, но нет активного 2D документа.
- `PASS`, если создан `out\table_M2.tbl` размером больше 0.

## Кодировка create_tbl.vbs (важно)

`create_tbl.vbs` должен быть в `UTF-16 LE` (с BOM) или ANSI (`cp1251`).
`UTF-8 BOM` часто вызывает ошибку VBS: `Недопустимый знак (1,1)`.

Проверка выполняется в `scripts/selfcheck.ps1`.

Пересохранение в UTF-16 LE через PowerShell:

```powershell
$p = ".\src\create_tbl.vbs"
Get-Content -LiteralPath $p -Raw | Set-Content -LiteralPath $p -Encoding Unicode
```

## Типовые ошибки

- `ERROR: Входной файл не найден`
  Проверьте путь к `.xlsx`.

- `ERROR: Не удалось запустить Excel COM`
  Excel не установлен или COM недоступен.

- `ERROR: Не удалось подключиться к КОМПАС`
  KOMPAS-3D не установлен/не зарегистрирован COM ProgID.

- `ERROR: Нет активного документа КОМПАС`
  Откройте в КОМПАС активный 2D документ (Фрагмент/Чертёж).

- `ERROR: 'table' mode requires table_width_mm and table_height_mm`
  Для `mode=table` должны быть заданы оба размера таблицы.
