' ============================================================================
'  @@TITLE@@
'
'  Reader Data Viewer -- the entry point that never shows a console.
'
'  A console window is not a setting: it comes from the SUBSYSTEM of whatever
'  program Windows starts. cmd.exe and powershell.exe are console programs, so
'  a .cmd always gets a window before any of our own code runs. wscript.exe,
'  which runs this file, is a GUI program and is never given one. That is the
'  whole difference between this file and the .cmd beside it -- the payload,
'  the compile and the app are identical.
'
'  usage:  @@USAGE@@
'          the data directory defaults to the "data" folder next to this file
'
'  The payload below is commented out line by line, because VBScript parses the
'  whole file before it runs a single statement. The bootstrap takes the leading
'  quote back off before running anything.
' ============================================================================
Option Explicit

Dim sh, fso, q, i, args, self, psExe, ps, cmdline
Set sh = CreateObject("WScript.Shell")
q = Chr(34)
self = WScript.ScriptFullName

args = ""
For i = 0 To WScript.Arguments.Count - 1
  args = args & " " & q & WScript.Arguments(i) & q
Next

' the bootstrap reads these two, exactly as the .cmd version does. They live in
' this process only: nothing is written to the registry or to any profile, and
' nothing has to be set up beforehand.
sh.Environment("PROCESS")("RDV_SELF") = self
sh.Environment("PROCESS")("RDV_ARGS") = args

psExe = sh.ExpandEnvironmentStrings("%SystemRoot%") & _
        "\System32\WindowsPowerShell\v1.0\powershell.exe"
If Not CreateObject("Scripting.FileSystemObject").FileExists(psExe) Then
  psExe = "powershell.exe"
End If

' Only single quotes in here, so the whole command can be wrapped in double
' quotes without escaping anything. '(?m)^''' is the regex that strips the
' leading quote from each payload line.
ps = "$a='#RDV'+'-PS-BEGIN'; $b='#RDV'+'-CS-BEGIN'; " & _
     "$t=[IO.File]::ReadAllText($env:RDV_SELF,[Text.Encoding]::UTF8); " & _
     "$i=$t.IndexOf($a); $j=$t.IndexOf($b); " & _
     "if($i -lt 0 -or $j -lt 0){ exit 9 }; " & _
     "$k='(?m)^'''; " & _
     "$g=$t.Substring($i+$a.Length,$j-$i-$a.Length) -replace $k,''; " & _
     "$global:RdvCs=$t.Substring($j+$b.Length) -replace $k,''; " & _
     "Invoke-Expression $g"

cmdline = q & psExe & q & " -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -Command " & q & ps & q

' 0 = hidden, False = do not wait. wscript exits immediately; the app owns the
' session from here.
sh.Run cmdline, 0, False
