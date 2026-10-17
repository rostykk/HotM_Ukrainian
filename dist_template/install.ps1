# Встановлення українізатора Heart of the Machine (Windows)
param([string]$GameDir = "")
$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ltRel = "GameData\Configuration\LocalizationType\LocalizationTypes.xml"

function Test-Game([string]$p) {
    return ($p -and (Test-Path -LiteralPath (Join-Path $p $ltRel)))
}

function Find-Game {
    # 1) архів розпаковано в теку гри (або поруч)
    foreach ($c in @($here, (Split-Path -Parent $here))) { if (Test-Game $c) { return $c } }
    # 2) бібліотеки Steam
    $roots = @()
    foreach ($k in @("HKCU:\Software\Valve\Steam", "HKLM:\SOFTWARE\WOW6432Node\Valve\Steam", "HKLM:\SOFTWARE\Valve\Steam")) {
        try {
            $v = Get-ItemProperty -Path $k -ErrorAction Stop
            if ($v.SteamPath) { $roots += $v.SteamPath }
            if ($v.InstallPath) { $roots += $v.InstallPath }
        } catch {}
    }
    $libs = @()
    foreach ($r in $roots) {
        $r = $r -replace '/', '\'
        $libs += $r
        $vdf = Join-Path $r "steamapps\libraryfolders.vdf"
        if (Test-Path -LiteralPath $vdf) {
            $raw = Get-Content -LiteralPath $vdf -Raw
            foreach ($m in [regex]::Matches($raw, '"path"\s+"([^"]+)"')) {
                $libs += ($m.Groups[1].Value -replace '\\\\', '\')
            }
        }
    }
    foreach ($l in ($libs | Select-Object -Unique)) {
        $c = Join-Path $l "steamapps\common\HeartOfTheMachine"
        if (Test-Game $c) { return $c }
    }
    return $null
}

Write-Host "Українізатор Heart of the Machine" -ForegroundColor Cyan
if (-not (Test-Game $GameDir)) { $GameDir = Find-Game }
if (-not $GameDir) {
    Write-Host "Не вдалося знайти гру автоматично."
    $GameDir = (Read-Host "Вставте шлях до теки HeartOfTheMachine (наприклад D:\SteamLibrary\steamapps\common\HeartOfTheMachine)").Trim().Trim('"')
}
if (-not (Test-Game $GameDir)) {
    Write-Host "У теці «$GameDir» не знайдено гру (немає $ltRel)." -ForegroundColor Red
    exit 1
}
Write-Host "Гра: $GameDir"

# 1. файл перекладу
$xlsx = Join-Path $here "Ukrainian.xlsx"
$dstDir = Join-Path $GameDir "GameData\LocalizationImports"
Copy-Item -LiteralPath $xlsx -Destination (Join-Path $dstDir "Ukrainian.xlsx") -Force
Write-Host "Скопійовано Ukrainian.xlsx -> GameData\LocalizationImports"

# 2. мова в LocalizationTypes.xml
$lt = Join-Path $GameDir $ltRel
$utf8 = New-Object System.Text.UTF8Encoding($false)
$text = [System.IO.File]::ReadAllText($lt, $utf8)
$block = [System.IO.File]::ReadAllText((Join-Path $here "localization_type_uk.xml"), $utf8).Trim("`r", "`n", [char]0xFEFF)
$nl = "`n"
if ($text.Contains("`r`n")) { $nl = "`r`n" }
$block = ($block -replace "`r`n", "`n") -replace "`n", $nl
$rx = [regex]'(?s)[ \t]*<localization_type\s+id="Ukrainian".*?/>'
if ($rx.IsMatch($text)) {
    $text = $rx.Replace($text, [System.Text.RegularExpressions.MatchEvaluator] { param($m) $block }, 1)
    Write-Host "Запис мови в LocalizationTypes.xml оновлено."
} else {
    Copy-Item -LiteralPath $lt -Destination "$lt.bak_before_ukrainian" -Force
    $i = $text.LastIndexOf("</root>")
    if ($i -lt 0) { Write-Host "LocalizationTypes.xml має неочікуваний формат." -ForegroundColor Red; exit 1 }
    $text = $text.Substring(0, $i).TrimEnd() + $nl + $nl + $block + $nl + $nl + $text.Substring($i)
    Write-Host "Українську мову додано в LocalizationTypes.xml (резервна копія: LocalizationTypes.xml.bak_before_ukrainian)."
}
[System.IO.File]::WriteAllText($lt, $text, $utf8)

Write-Host ""
Write-Host "Готово! Запустіть гру й виберіть «Українська» у списку мов унизу головного меню." -ForegroundColor Green
Write-Host "Після кожного оновлення гри запустіть «Встановити.bat» ще раз: Steam перезаписує LocalizationTypes.xml."
