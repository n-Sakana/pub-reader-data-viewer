# ---------------------------------------------------------------------------
# boot-hybrid.ps1 -- method 3, C# engine + Excel front.
# The workbook is the other half of the distribution and must sit beside this
# .cmd. It is opened read-only; the app writes cells for display and can never
# save over the file.
# ---------------------------------------------------------------------------
$RdvBook = Join-Path $RdvHere 'ReaderDataViewer-Hybrid.xlsm'
foreach ($t in $RdvTokens) {
  if ($t -like '*.xlsm') { $RdvBook = $t }
}
if (-not (Test-Path -LiteralPath $RdvBook)) {
  Rdv-Fail ("ReaderDataViewer-Hybrid.xlsm was not found beside this .cmd:`r`n  " + $RdvBook)
}
$RdvBook = (Resolve-Path -LiteralPath $RdvBook).Path
# see boot-csharp.ps1: exit does not evaluate a method call in its argument
$RdvRc = [RdvProgramHybrid]::Run($RdvData, $RdvBook, $RdvCompileMs, $RdvLog)
exit $RdvRc
