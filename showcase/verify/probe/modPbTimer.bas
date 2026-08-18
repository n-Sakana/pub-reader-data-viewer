Attribute VB_Name = "modPbTimer"
'==============================================================================
' modPbTimer -- VERIFICATION ONLY.
'
' The requirement document asks for a 0.5 s poll and says (section 10) to
' measure what Application.OnTime actually does and to include "go back to 1 s"
' in the decision. A first pass showed that OnTime Now + 0.5 s fires IMMEDIATELY
' -- but that is ambiguous: Now truncates to whole seconds, so the target was
' usually already in the past. This tells the two apart.
'
'   phase 1  one chain at Now + 1.5 s. Gaps near 1.5 s mean sub-second targets
'            are honoured; gaps near 1.0 / 2.0 mean OnTime works in whole
'            seconds and 0.5 s is not reachable by asking for it.
'   phase 2  TWO chains, each at Now + 1 s, started half a second apart. If
'            sub-second targets are honoured they stay interleaved and the
'            combined rate is 2 Hz -- a 0.5 s pump with no blocking anywhere.
'==============================================================================
Option Explicit

Private m_res As String
Private m_n As Long
Private m_t0 As Double
Private m_last As Double
Private m_nA As Long
Private m_nB As Long
Private m_lastAny As Double
Private m_t0Any As Double

Private Sub Say(ByVal s As String)
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
    For i = 1 To 60
        Err.Clear
        f = FreeFile
        Open m_res For Append As #f
        If Err.Number = 0 Then
            Print #f, s
            Close #f
            Exit For
        End If
        DoEvents
    Next i
End Sub

Private Function Qual(ByVal procName As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

Public Sub PbTmPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG  project compiles"
End Sub

Public Sub PbTmArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, Qual("PbTmRun")
End Sub

Public Sub PbTmRun()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "=== timer probe start ==="
    Say "START phase1_1500ms"
    m_n = 0
    m_t0 = Timer
    m_last = m_t0
    Application.OnTime Now + 1.5 / 86400#, Qual("PbTmP1")
End Sub

Public Sub PbTmP1()
    Dim t As Double
    t = Timer
    m_n = m_n + 1
    Say "TICK  p1 n=" & m_n & " gap=" & Format$(t - m_last, "0.000") & _
        " total=" & Format$(t - m_t0, "0.000")
    m_last = t
    If m_n < 6 Then
        Application.OnTime Now + 1.5 / 86400#, Qual("PbTmP1")
        Exit Sub
    End If
    Say "OK    phase1_1500ms = 6 ticks in " & Format$(t - m_t0, "0.000") & "s"
    StartPhase2
End Sub

' two chains, half a second apart. The one-off half-second offset is a spin
' with DoEvents -- acceptable ONCE, at set-up, and never in the pump itself.
Private Sub StartPhase2()
    Dim t As Double
    Say "START phase2_two_chains"
    m_nA = 0
    m_nB = 0
    m_t0Any = Timer
    m_lastAny = m_t0Any
    Application.OnTime Now + 1# / 86400#, Qual("PbTmA")
    t = Timer
    Do While Timer - t < 0.5
        DoEvents
    Loop
    Application.OnTime Now + 1# / 86400#, Qual("PbTmB")
End Sub

Public Sub PbTmA()
    Dim t As Double
    t = Timer
    m_nA = m_nA + 1
    Say "TICK  A n=" & m_nA & " gapAny=" & Format$(t - m_lastAny, "0.000") & _
        " total=" & Format$(t - m_t0Any, "0.000")
    m_lastAny = t
    If m_nA < 6 Then Application.OnTime Now + 1# / 86400#, Qual("PbTmA")
    MaybeDone
End Sub

Public Sub PbTmB()
    Dim t As Double
    t = Timer
    m_nB = m_nB + 1
    Say "TICK  B n=" & m_nB & " gapAny=" & Format$(t - m_lastAny, "0.000") & _
        " total=" & Format$(t - m_t0Any, "0.000")
    m_lastAny = t
    If m_nB < 6 Then Application.OnTime Now + 1# / 86400#, Qual("PbTmB")
    MaybeDone
End Sub

Private Sub MaybeDone()
    If m_nA >= 6 And m_nB >= 6 Then
        Say "OK    phase2_two_chains = " & (m_nA + m_nB) & " ticks in " & _
            Format$(Timer - m_t0Any, "0.000") & "s"
        Say "=== timer probe done ==="
    End If
End Sub
