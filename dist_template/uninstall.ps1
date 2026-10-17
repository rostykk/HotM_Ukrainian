# Видалення українізатора Heart of the Machine (Windows)
param([string]$GameDir = "")
$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ltRel = "GameData\Configuration\LocalizationType\LocalizationTypes.xml"

function Test-Game([string]$p) {
    return ($p -and (Test-Path -LiteralPath (Join-Path $p $ltRel)))
}

function Find-Game {
    foreach ($c in @($here, (Split-Path -Parent $here))) { if (Test-Game $c) { return $c } }
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

Write-Host "Видалення українізатора Heart of the Machine" -ForegroundColor Cyan
if (-not (Test-Game $GameDir)) { $GameDir = Find-Game }
if (-not $GameDir) {
    $GameDir = (Read-Host "Вставте шлях до теки HeartOfTheMachine").Trim().Trim('"')
}
if (-not (Test-Game $GameDir)) {
    Write-Host "У теці «$GameDir» не знайдено гру." -ForegroundColor Red
    exit 1
}

$xlsx = Join-Path $GameDir "GameData\LocalizationImports\Ukrainian.xlsx"
if (Test-Path -LiteralPath $xlsx) { Remove-Item -LiteralPath $xlsx -Force; Write-Host "Видалено Ukrainian.xlsx" }

$lt = Join-Path $GameDir $ltRel
$utf8 = New-Object System.Text.UTF8Encoding($false)
$text = [System.IO.File]::ReadAllText($lt, $utf8)
$rx = [regex]'(?s)(\r?\n)*[ \t]*<localization_type\s+id="Ukrainian".*?/>'
if ($rx.IsMatch($text)) {
    $text = $rx.Replace($text, "", 1)
    [System.IO.File]::WriteAllText($lt, $text, $utf8)
    Write-Host "Українську мову прибрано з LocalizationTypes.xml"
}
Write-Host "Готово. Якщо в грі була вибрана українська, гра повернеться до англійської." -ForegroundColor Green
