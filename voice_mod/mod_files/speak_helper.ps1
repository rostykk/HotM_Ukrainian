# UkrVoice prototype helper (fallback without StyleTTS2). The game mod starts this script and writes:
#   SAY<TAB>voice<TAB><Table>|<ID><TAB>base64 text   e.g. key NPCDialog|Ch1_Foo, NPCEvent|Cont_X
#   STOP                                              window closed, stop speaking
# For each key we speak the ENGLISH original (Windows has no Ukrainian voice) so you can check
# by ear that the text on screen and the detected key match.
# Test without the game:  echo NPCDialog^|SomeId | powershell -File speak_helper.ps1

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Speech
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$log = Join-Path $here 'speak_helper_log.txt'
function Log($m) { Add-Content -Path $log -Value ((Get-Date -Format 'HH:mm:ss') + "`t" + $m) -Encoding UTF8 }

$en = @{}
$tsv = Join-Path $here 'en_lookup.tsv'
if (Test-Path $tsv) {
    foreach ($l in [IO.File]::ReadAllLines($tsv, [Text.Encoding]::UTF8)) {
        $i = $l.IndexOf("`t")
        if ($i -gt 0) { $en[$l.Substring(0, $i)] = $l.Substring($i + 1).Replace('\n', "`n") }
    }
}
Log ("helper started, english lines: " + $en.Count)

$s = New-Object System.Speech.Synthesis.SpeechSynthesizer
try { $s.SetOutputToDefaultAudioDevice() } catch {}
$s.Rate = 1

function Field($k) { if ($en.ContainsKey($k)) { return $en[$k] } return '' }
function Join-Parts($parts) { (($parts | Where-Object { $_ }) -join "`n") }

function Compose($table, $id) {
    switch ($table) {
        'NPCEvent' {
            $t = Join-Parts @((Field "NPCEvent|$id|high_text"), (Field "NPCEvent|$id|description"), (Field "NPCEvent|$id|strategy_tip"))
            if (-not $t -and $id.StartsWith('Cont_')) {
                $c = $id.Substring(5)
                $t = Join-Parts @((Field "ContemplationType|$c|description"), (Field "ContemplationType|$c|strategy_tip"))
            }
            return $t
        }
        'NPCDialog' { return Join-Parts @((Field "NPCDialog|$id|description"), (Field "NPCDialog|$id|strategy_tip")) }
        'OtherKeyMessage' { return Join-Parts @((Field "OtherKeyMessage|$id|high_text"), (Field "OtherKeyMessage|$id|description"), (Field "OtherKeyMessage|$id|strategy_tip")) }
        'MachineHandbookEntry' { return Join-Parts @((Field "MachineHandbookEntry|$id|description"), (Field "MachineHandbookEntry|$id|strategy_tip")) }
        'PROVIDER' { return "Unmapped window: $id" }
        'RESULT' { return "Choice result." }
    }
    return ''
}

function Clean($t) {
    $t = $t -replace '<[^>]+>', ' '
    $t = $t -replace '\{[^}]*\}', ' '
    $t = $t -replace '\s+', ' '
    return $t.Trim()
}

while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }      # game closed our stdin
    $line = $line.Trim()
    if (-not $line) { continue }
    if ($line -eq 'STOP' -or $line -eq 'STOP_HOVER') { $s.SpeakAsyncCancelAll(); continue }
    if ($line.StartsWith("SAY`t")) { $line = $line.Split("`t")[2] }   # SAY <voice> <key> <base64 text>
    $p = $line.Split('|', 2)
    $table = $p[0]; $id = if ($p.Count -gt 1) { $p[1] } else { '' }
    $text = Compose $table $id
    if (-not $text) { $text = "No English text for $table $id" }
    Log ("$line`t" + $text.Length + " chars")
    $s.SpeakAsyncCancelAll()
    [void]$s.SpeakAsync((Clean $text))
}
Log 'helper stopped'
