# app-kompas-text-sync

Утилита с GUI для двусторонней синхронизации текстов между активным чертежом KOMPAS-3D и Excel `.xlsx` через COM.

## Что делает

- В GUI есть кнопка-переключатель `Включить/Остановить синхронизацию`.
- Параметры (`corridor_mm`, `sheet_name`) сохраняются автоматически при изменении.
- При включении синхронизации:
  - определяется текущий активный документ KOMPAS;
  - рядом с ним создается/открывается Excel-файл с тем же именем (`<drawing>.xlsx`);
  - при переключении активной вкладки чертежа утилита автоматически переходит на соответствующий Excel-файл.
- Текстовые элементы группируются в режиме вертикальной синхронизации:
  - тексты в коридоре `± corridor_mm` по вертикали попадают в одну строку;
  - внутри строки сортировка слева-направо;
  - строки сортируются сверху-вниз.
- Изменения двусторонние:
  - Excel -> KOMPAS;
  - KOMPAS -> Excel.
- Если в Excel заполнена ячейка без привязки к тексту чертежа, она красится в красный фон до очистки.
- Для отказоустойчивости используется:
  - периодическое сохранение Excel;
  - атомарный JSON состояния рядом с таблицей (`<drawing>.json` рядом с `<drawing>.xlsx`), с автосохранением.
  - runtime-status JSON в `out/sync_runtime_status.json` для отображения фактического состояния синхронизации в GUI.

## Структура

- `scripts/gui_sync.ps1` - GUI.
- `scripts/kompas_excel_text_sync.py` - синхронизатор (COM KOMPAS + COM Excel).
- `scripts/sync_logic.py` - чистая логика группировки/решения конфликтов.
- `scripts/run_gui.cmd` / `scripts/run_gui.vbs` - запуск.
- `scripts/kompas_button_launcher.vbs` - запуск из кнопки KOMPAS.
- `scripts/build_launcher.ps1` + `scripts/AppLauncher.cs` - сборка `bin/app-kompas-text-sync.exe`.
- `scripts/bind_to_kompas.ps1` - локальная подготовка параметров биндинга.
- `scripts/selfcheck.ps1` - проверка окружения.

## Запуск

```bat
cd C:\__MY_PROJECTS__\git\kompas3d-utility\app-kompas-text-sync
.\scripts\run_gui.cmd
```

## Привязка кнопки в KOMPAS

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bind_to_kompas.ps1
```

Глобальный авто-биндер для всех `app-*`:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\bind_all_apps_to_kompas.ps1 -WritePerAppFiles
```

## Самопроверка

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\selfcheck.ps1
```

## Требования

- Windows
- KOMPAS-3D COM (`Kompas.Application.5`)
- Microsoft Excel COM (`Excel.Application`)
- Python 3 + `pywin32`
