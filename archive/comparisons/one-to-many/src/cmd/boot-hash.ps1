# ---------------------------------------------------------------------------
# boot-hash.ps1 -- one-to-many build, hand-built hash index.
# ---------------------------------------------------------------------------
$RdvPick = ''
for ($i = 0; $i -lt $RdvTokens.Count; $i++) {
  if ($RdvTokens[$i] -eq '-autopick' -and ($i + 1) -lt $RdvTokens.Count) { $RdvPick = $RdvTokens[$i + 1] }
}
# see boot-csharp.ps1: exit does not evaluate a method call in its argument
$RdvRc = [Rdv2Program]::Run($RdvData, $RdvCompileMs, $RdvLog, $RdvPick)
exit $RdvRc
