# ============================================================================
# gen_emitter.ps1 -- generate vba\modZipEmit.bas from the three reference files
#                    that Excel has to be able to write at run time:
#
#     worker\ZipWorker.cs        the C# source
#     worker\run_worker.bat      the launcher
#     worker\build_worker.ps1    the builder
#
# The "emitted" worker variant means Excel writes those three files itself, with
# no exe and nothing pre-installed. Hand-maintaining a second copy of ~1000
# lines of C# inside VBA string literals would drift within a day, so the copy
# is generated from the originals instead. Run this after touching any of them.
#
# Everything must survive a round trip through CP932, because .bas files are
# imported into the VBA project in the system ANSI code page. check_cp932.ps1
# is run as part of this script and fails loudly rather than silently mangling
# a character.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$sources = [ordered]@{
  'WorkerCs'   = @{ path = "$Root\worker\ZipWorker.cs";     desc = 'C# ソース (UTF-8 BOM つきで書き出す)' }
  'RunBat'     = @{ path = "$Root\worker\run_worker.bat";   desc = '起動用 BAT (ASCII のみ)' }
  'BuildPs1'   = @{ path = "$Root\worker\build_worker.ps1"; desc = 'ビルド用 PS1 (ASCII のみ)' }
}

$cp932 = [Text.Encoding]::GetEncoding(932)
$sb = New-Object System.Text.StringBuilder
$LINES_PER_FUNC = 110

function Add-Line([string] $s) { [void]$sb.AppendLine($s) }

Add-Line 'Attribute VB_Name = "modZipEmit"'
Add-Line "'=============================================================================="
Add-Line "' modZipEmit -- 自動生成。手で直さないこと。"
Add-Line "'"
Add-Line "' build\gen_emitter.ps1 が worker\ 配下の 3 つの原本から作る:"
foreach ($k in $sources.Keys) {
  Add-Line ("'     " + (Split-Path $sources[$k].path -Leaf).PadRight(20) + $sources[$k].desc)
}
Add-Line "'"
Add-Line "' 「emitted」版のワーカーは、この 3 つを Excel がその場で書き出して建てる。"
Add-Line "' 原本を書き換えたら gen_emitter.ps1 を流し直すこと。そうしないと"
Add-Line "' 参照コピーと Excel が吐くものがずれる。"
Add-Line "'=============================================================================="
Add-Line 'Option Explicit'
Add-Line ''

foreach ($key in $sources.Keys) {
  $path = $sources[$key].path
  if (-not (Test-Path -LiteralPath $path)) { throw "missing source: $path" }
  $text = [IO.File]::ReadAllText($path)

  $rt = $cp932.GetString($cp932.GetBytes($text))
  if ($text -ne $rt) {
    $bad = @{}
    for ($i = 0; $i -lt [Math]::Min($text.Length, $rt.Length); $i++) {
      if ($text[$i] -ne $rt[$i]) { $bad[("U+{0:X4}" -f [int]$text[$i])] = 1 }
    }
    throw ("$path contains characters that cannot survive CP932: " + ($bad.Keys -join ', ') +
           " -- escape them (C# \uXXXX) before generating the emitter")
  }

  $lines = $text -split "`r`n|`n"
  # a trailing newline produces one empty final element; drop it, WriteLine adds it back
  if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') {
    $lines = $lines[0..($lines.Count - 2)]
  }

  $nFuncs = [Math]::Ceiling($lines.Count / $LINES_PER_FUNC)
  if ($nFuncs -lt 1) { $nFuncs = 1 }

  Add-Line ("'--- " + (Split-Path $path -Leaf) + " : " + $lines.Count + " 行 ---")
  Add-Line ("Public Function " + $key + "Text() As String")
  Add-Line ('    Dim s As String')
  for ($f = 1; $f -le $nFuncs; $f++) {
    Add-Line ('    s = s & ' + $key + '_' + $f.ToString('00') + '()')
  }
  Add-Line ('    ' + $key + 'Text = s')
  Add-Line 'End Function'
  Add-Line ''

  for ($f = 1; $f -le $nFuncs; $f++) {
    $from = ($f - 1) * $LINES_PER_FUNC
    $to = [Math]::Min($lines.Count - 1, $from + $LINES_PER_FUNC - 1)
    $count = $to - $from + 1
    Add-Line ("Private Function " + $key + '_' + $f.ToString('00') + '() As String')
    Add-Line ('    Dim a(0 To ' + ($count - 1) + ') As String')
    for ($i = $from; $i -le $to; $i++) {
      $esc = $lines[$i].Replace('"', '""')
      Add-Line ('    a(' + ($i - $from) + ') = "' + $esc + '"')
    }
    Add-Line ('    ' + $key + '_' + $f.ToString('00') + ' = Join(a, vbCrLf) & vbCrLf')
    Add-Line 'End Function'
    Add-Line ''
  }
}

$out = "$Root\vba\modZipEmit.bas"
[IO.File]::WriteAllText($out, $sb.ToString(), [Text.UTF8Encoding]::new($false))

# final safety net: the generated module itself must round-trip through CP932
$g = [IO.File]::ReadAllText($out)
if ($g -ne $cp932.GetString($cp932.GetBytes($g))) { throw "generated $out is not CP932-safe" }

Write-Output ("wrote {0}  ({1} bytes)" -f $out, (Get-Item -LiteralPath $out).Length)
foreach ($k in $sources.Keys) {
  Write-Output ("  {0,-12} <- {1}" -f ($k + 'Text()'), $sources[$k].path)
}
