# ---------------------------------------------------------------------------
# boot-csharp.ps1 -- method 2, C# only. WinForms shows the record.
# ---------------------------------------------------------------------------
# NOT "exit [RdvProgramForms]::Run(...)": PowerShell's exit statement does not
# evaluate a method call in its argument, it just leaves. Measured: an exit whose
# argument was a 9-second Sleep returned instantly. Call first, then exit.
$RdvRc = [RdvProgramForms]::Run($RdvData, $RdvCompileMs, $RdvLog)
exit $RdvRc
