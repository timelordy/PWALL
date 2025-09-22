<#
.SYNOPSIS
    Создаёт структуру вкладки LayerTools и обновляет панель SplitLayers.
.DESCRIPTION
    Формирует структуру
        LayerTools.tab/LayerTools.panel/SplitLayers.pushbutton
    и копирует внутрь неё актуальный pushbutton из каталога Latest.
.PARAMETER SourcePushbuttonPath
    Необязательный путь к каталогу SplitLayers.pushbutton, который нужно
    разложить во вкладку LayerTools. По умолчанию берётся папка
    Latest/SplitLayers.pushbutton, расположенная рядом со скриптом.
#>

param(
    [string]
    $SourcePushbuttonPath = (Join-Path -Path $PSScriptRoot -ChildPath "Latest\SplitLayers.pushbutton")
)

$extensionRoot = $PSScriptRoot
$tabPath = Join-Path -Path $extensionRoot -ChildPath "LayerTools.tab"
$panelPath = Join-Path -Path $tabPath -ChildPath "LayerTools.panel"
$targetPushbuttonPath = Join-Path -Path $panelPath -ChildPath "SplitLayers.pushbutton"

New-Item -ItemType Directory -Path $tabPath -Force | Out-Null
New-Item -ItemType Directory -Path $panelPath -Force | Out-Null

Remove-Item -Path $targetPushbuttonPath -Recurse -Force -ErrorAction SilentlyContinue

if (-not (Test-Path -Path $SourcePushbuttonPath)) {
    throw "Не найден источник pushbutton по пути $SourcePushbuttonPath. " +
          "Убедитесь, что рядом со скриптом есть каталог Latest\\SplitLayers.pushbutton или передайте свой путь через параметр -SourcePushbuttonPath."
}

Copy-Item -Path $SourcePushbuttonPath -Destination $panelPath -Recurse -Force

Write-Host "Готово! Каталог с кнопкой разложен по пути:`n$targetPushbuttonPath"
