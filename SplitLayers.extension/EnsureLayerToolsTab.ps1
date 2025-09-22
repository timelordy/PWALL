<#
.SYNOPSIS
    Создаёт каталог вкладки LayerTools и при необходимости переносит панель LayerTools.panel внутрь него.
.DESCRIPTION
    Скрипт аккуратно приводит структуру расширения pyRevit к виду
    SplitLayers.extension/LayerTools.tab/LayerTools.panel/...
    Он безопасно выполняется повторно: если элементы уже на месте, просто выводит сообщение.
#>

$extensionRoot = $PSScriptRoot
$tabPath = Join-Path -Path $extensionRoot -ChildPath "LayerTools.tab"
$panelOldPath = Join-Path -Path $extensionRoot -ChildPath "LayerTools.panel"
$panelNewPath = Join-Path -Path $tabPath -ChildPath "LayerTools.panel"

if (-not (Test-Path -Path $tabPath)) {
    New-Item -ItemType Directory -Path $tabPath -Force | Out-Null
    Write-Host "Создан каталог вкладки:`n$tabPath"
} else {
    Write-Host "Каталог вкладки уже существует:`n$tabPath"
}

if ((Test-Path -Path $panelOldPath) -and -not (Test-Path -Path $panelNewPath)) {
    Move-Item -Path $panelOldPath -Destination $tabPath
    Write-Host "Панель LayerTools.panel перенесена в каталог вкладки."
} elseif (Test-Path -Path $panelNewPath) {
    Write-Host "Панель уже находится внутри LayerTools.tab:`n$panelNewPath"
} else {
    Write-Warning "Каталог панели не найден. Убедитесь, что рядом со скриптом есть LayerTools.panel или уже создана правильная структура."
}
