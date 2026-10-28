# Builds the UkrVoice prototype and installs it into <game>\XMLMods\UkrVoice.
#   powershell -ExecutionPolicy Bypass -File voice_mod\install.ps1
#   powershell -ExecutionPolicy Bypass -File voice_mod\install.ps1 -TtsPython D:\zp-tts\venv\Scripts\python.exe -TtsEnv D:\zp-tts\env.ps1
# With a Python that has StyleTTS2 (styletts2-inference, ipa-uk, ukrainian-word-stress, torch) the mod speaks
# Ukrainian live with the StyleTTS2-RosKud voice; otherwise it falls back to the Windows voice (English original).
param(
    [string]$TtsPython = 'D:\zp-tts\venv\Scripts\python.exe',
    [string]$TtsEnv = 'D:\zp-tts\env.ps1'
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$proj = Split-Path -Parent $here
$game = Split-Path -Parent $proj
$dest = Join-Path $game 'XMLMods\UkrVoice'

$msb = Get-ChildItem 'C:\Program Files (x86)\Microsoft Visual Studio', 'C:\Program Files\Microsoft Visual Studio' -Recurse -Filter MSBuild.exe -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch 'amd64|arm64' } | Sort-Object FullName -Descending | Select-Object -First 1
if (-not $msb) { throw 'MSBuild not found (install Visual Studio Build Tools)' }
& $msb.FullName (Join-Path $here 'UkrVoice.csproj') /p:Configuration=Release /v:m /nologo
if ($LASTEXITCODE -ne 0) { throw 'build failed' }

New-Item -ItemType Directory -Force (Join-Path $dest 'ModdableLogicDLLs') | Out-Null
Copy-Item (Join-Path $here 'mod_files\*') $dest -Recurse -Force
Copy-Item (Join-Path $here 'bin\UkrVoice.dll') (Join-Path $dest 'ModdableLogicDLLs\UkrVoice.dll') -Force
foreach ($f in 'stressed_uk.tsv', 'stress_fixes.txt', 'speech_replace.tsv') {
    $src = Join-Path $here "stress\$f"
    if (Test-Path $src) { Copy-Item $src (Join-Path $dest $f) -Force }
}
python -I -X utf8 (Join-Path $here 'make_en_lookup.py') (Join-Path $dest 'en_lookup.tsv')

$cmd = Join-Path $dest 'helper_cmd.txt'
$model = Join-Path $proj 'StyleTTS2-RosKud'
if ((Test-Path $TtsPython) -and (Test-Path (Join-Path $model 'pytorch_model.bin'))) {
    $args2 = "-u `"{mod}\tts_helper.py`" --model `"$model`""
    if (Test-Path $TtsEnv) { $args2 += " --envfile `"$TtsEnv`"" }
    [IO.File]::WriteAllLines($cmd, @($TtsPython, $args2), (New-Object Text.UTF8Encoding($false)))
    Write-Host "Voice: StyleTTS2-RosKud via $TtsPython"
} else {
    if (Test-Path $cmd) { Remove-Item $cmd }
    Write-Host 'Voice: Windows TTS fallback (English original)'
}
Write-Host "Installed to $dest"
