# PWALL

PyRevit расширение **SplitLayers** добавляет инструмент "Разделить стену по слоям".

Основные возможности:

- выбор составной стены (тип Basic) и анализ её слоёв;
- выбор слоёв для разделения и сопоставление их существующим типам стен или автоматическое создание новых типов;
- создание отдельных однослойных стен со смещением по центрам слоёв;
- перенос окон и дверей на выбранную новую стену;
- копирование параметров исходной стены и объединение новых стен (Join Geometry);
- удаление исходной стены после успешного разделения.

Инструмент рассчитан на pyRevit (Python 3.11) и Revit 2022.

## Настройка структуры расширения

Чтобы pyRevit увидел кнопку **SplitLayers** на вкладке `LayerTools`, структура папок должна выглядеть так:

```
SplitLayers.extension/
└── LayerTools.tab/
    └── LayerTools.panel/
        └── SplitLayers.pushbutton/
            └── script.py
```

Каталог `SplitLayers.extension/Latest/SplitLayers.pushbutton` содержит "эталонный" набор файлов кнопки. Скрипт развёртывания всегда копирует именно его.

### Автоматическая настройка (рекомендуется)

1. Откройте PowerShell и перейдите в корень репозитория `SplitLayers.extension`:
   ```powershell
   Set-Location "C:\Users\<имя>\PycharmProjects\PWALL\SplitLayers.extension"
   ```
2. Запустите скрипт, который пересоберёт вкладку и скопирует свежий pushbutton:
   ```powershell
   .\EnsureLayerToolsTab.ps1
   ```
   Скрипт делает всё сам:
   - создаёт (или заново создаёт) каталоги `LayerTools.tab` и `LayerTools.panel` с флагом `-Force`;
   - удаляет возможную старую версию `LayerTools.panel\SplitLayers.pushbutton`;
   - копирует содержимое `Latest\SplitLayers.pushbutton` в нужное место.

   Если нужно разложить кнопку из другого источника (например, скачали архив с новой версией), передайте путь параметром:
   ```powershell
   .\EnsureLayerToolsTab.ps1 -SourcePushbuttonPath "D:\Temp\SplitLayers.pushbutton"
   ```

### Ручная настройка (для понимания процесса)

Хотите проделать всё вручную — действуйте так:

```powershell
New-Item -ItemType Directory -Path "LayerTools.tab" -Force | Out-Null
New-Item -ItemType Directory -Path "LayerTools.tab\LayerTools.panel" -Force | Out-Null
Remove-Item -Path "LayerTools.tab\LayerTools.panel\SplitLayers.pushbutton" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -Path "Latest\SplitLayers.pushbutton" -Destination "LayerTools.tab\LayerTools.panel" -Recurse -Force
```

Такая последовательность безоговорочно сформирует правильную структуру и закинет актуальный скрипт. После этого перезагрузите pyRevit (или заново добавьте расширение) и убедитесь, что вкладка `LayerTools` появилась с панелью и кнопкой **SplitLayers**.

Обновление скрипта теперь сводится к замене файлов в каталоге `Latest\SplitLayers.pushbutton` и повторному запуску `EnsureLayerToolsTab.ps1`.
