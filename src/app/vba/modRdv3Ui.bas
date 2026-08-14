Attribute VB_Name = "modRdv3Ui"
'==============================================================================
' modRdv3Ui -- everything the UI sheet shows. Rendering only; no decisions.
'
' The screen is cells: fills, borders, fonts and merges the builder painted.
' No UserForm, no ActiveX, no Forms controls, no ListBox, no Shape buttons.
' The buttons are hyperlink cells (ScreenTip = the hover explanation) and the
' clicks arrive in modRdv3App via Worksheet_FollowHyperlink.
'
' The two performance figures the spec allows on screen -- merge time and
' search time -- are written here and nowhere else. Every finer figure goes to
' the execution log.
'
' The "checking for updates" animation is a dot cycle on the state cell.
' Rdv3UiAnimPump is called from inside every chunked loop, so the dots move
' whether the wait is an OnTime poll or a chunked read.
'==============================================================================
Option Explicit

Private m_animBase As String
Private m_animStep As Long
Private m_animLast As Double
Private m_animOn As Boolean

Private Function UiWs() As Object
    Set UiWs = ThisWorkbook.Worksheets(RDV3_SHEET_UI)
End Function

Public Sub Rdv3UiState(ByVal text As String)
    On Error Resume Next
    UiWs().Range(RDV3_C_STATE).Value2 = text
End Sub

Public Sub Rdv3UiNotepad(ByVal text As String)
    On Error Resume Next
    UiWs().Range(RDV3_C_NOTEPAD).Value2 = text
End Sub

Public Sub Rdv3UiLedgerInfo(ByVal text As String)
    On Error Resume Next
    UiWs().Range(RDV3_C_LEDINFO).Value2 = text
End Sub

Public Sub Rdv3UiMergeMs(ByVal ms As Double)
    On Error Resume Next
    If ms >= 0 Then
        UiWs().Range(RDV3_C_MERGEMS).Value2 = Rdv3FmtMs(ms)
    Else
        UiWs().Range(RDV3_C_MERGEMS).Value2 = "--"
    End If
End Sub

Public Sub Rdv3UiSearchMs(ByVal ms As Double)
    On Error Resume Next
    If ms >= 0 Then
        UiWs().Range(RDV3_C_SEARCHMS).Value2 = Rdv3FmtMs(ms)
    Else
        UiWs().Range(RDV3_C_SEARCHMS).Value2 = "--"
    End If
End Sub

Public Sub Rdv3UiError(ByVal text As String)
    On Error Resume Next
    UiWs().Range(RDV3_C_ERROR).Value2 = text
End Sub

Public Sub Rdv3UiWatchButton(ByVal running As Boolean)
    On Error Resume Next
    If running Then
        UiWs().Range(RDV3_BTN_WATCH).Value2 = "監視停止"
    Else
        UiWs().Range(RDV3_BTN_WATCH).Value2 = "監視再開"
    End If
End Sub

Public Function Rdv3UiInputKey() As String
    On Error Resume Next
    Rdv3UiInputKey = Trim$(CStr(UiWs().Range(RDV3_C_INPUT).Value2))
End Function

'------------------------------------------------------------------------------
' the checking animation
'------------------------------------------------------------------------------
Public Sub Rdv3UiAnimStart(ByVal baseText As String)
    m_animBase = baseText
    m_animStep = 0
    m_animLast = 0
    m_animOn = True
    Rdv3UiState baseText
End Sub

Public Sub Rdv3UiAnimStop()
    m_animOn = False
End Sub

' cheap enough to call from every chunk: repaints at most ~3 times a second
Public Sub Rdv3UiAnimPump()
    If Not m_animOn Then Exit Sub
    If Timer - m_animLast < 0.3 And Timer >= m_animLast Then Exit Sub
    m_animLast = Timer
    m_animStep = (m_animStep + 1) Mod 4
    Rdv3UiState m_animBase & String$(m_animStep, "・")
End Sub

'------------------------------------------------------------------------------
' result region
'------------------------------------------------------------------------------
Public Sub Rdv3UiShowRecord(ByVal key As String, ByVal verdict As String, _
                            ByRef vA() As String, ByRef vB() As String, ByRef vC() As String, _
                            ByVal procText As String)
    Dim ws As Object
    Dim rec(1 To 10, 1 To 7) As Variant
    Dim i As Long

    Set ws = UiWs()
    For i = 1 To RDV3_FIELDS
        rec(i, 1) = i
        rec(i, 2) = Rdv3NameOfA(i - 1)
        rec(i, 3) = vA(i - 1)
        rec(i, 4) = Rdv3NameOfB(i - 1)
        rec(i, 5) = vB(i - 1)
        rec(i, 6) = Rdv3NameOfC(i - 1)
        rec(i, 7) = vC(i - 1)
    Next i
    ws.Range(RDV3_C_KEY).Value2 = key
    ws.Range(RDV3_C_VERDICT).Value2 = verdict
    ws.Range(RDV3_C_PROCSTATE).Value2 = procText
    ws.Range("B" & RDV3_REC_TOP & ":H" & (RDV3_REC_TOP + RDV3_FIELDS - 1)).Value2 = rec
    Rdv3UiClearCandidates
End Sub

Public Sub Rdv3UiShowEmpty(ByVal key As String, ByVal verdict As String)
    Dim ws As Object
    Set ws = UiWs()
    ws.Range(RDV3_C_KEY).Value2 = key
    ws.Range(RDV3_C_VERDICT).Value2 = verdict
    ws.Range(RDV3_C_PROCSTATE).Value2 = ""
    ws.Range("B" & RDV3_REC_TOP & ":H" & (RDV3_REC_TOP + RDV3_FIELDS - 1)).ClearContents
    Rdv3UiClearCandidates
End Sub

' rows(1..n, 1..8) -- the candidate slots below the record area. The 選択
' hyperlink cells are static; only their row text changes.
Public Sub Rdv3UiShowCandidates(ByVal key As String, ByVal verdict As String, _
                                ByRef rows() As Variant, ByVal n As Long, ByVal note As String)
    Dim ws As Object
    Dim i As Long, c As Long
    Dim arr() As Variant

    Set ws = UiWs()
    ws.Range(RDV3_C_KEY).Value2 = key
    ws.Range(RDV3_C_VERDICT).Value2 = verdict
    ws.Range(RDV3_C_PROCSTATE).Value2 = ""
    ws.Range("B" & RDV3_REC_TOP & ":H" & (RDV3_REC_TOP + RDV3_FIELDS - 1)).ClearContents
    Rdv3UiClearCandidates
    If Len(note) > 0 Then
        ws.Range(RDV3_C_CANDNOTE).Value2 = RDV3_CAND_LABEL & "   " & note
    End If
    If n <= 0 Then Exit Sub
    ReDim arr(1 To n, 1 To 8)
    For i = 1 To n
        For c = 1 To 8
            arr(i, c) = rows(i, c)
        Next c
        ws.Cells(RDV3_CAND_TOP + i - 1, 2).Value2 = "選択"
    Next i
    ws.Range(ws.Cells(RDV3_CAND_TOP, 3), ws.Cells(RDV3_CAND_TOP + n - 1, 10)).Value2 = arr
End Sub

Public Sub Rdv3UiClearCandidates()
    Dim ws As Object
    Set ws = UiWs()
    ws.Range(RDV3_C_CANDNOTE).Value2 = RDV3_CAND_LABEL
    ws.Range(ws.Cells(RDV3_CAND_TOP, 2), ws.Cells(RDV3_CAND_TOP + RDV3_MAXSHOW - 1, 10)).ClearContents
End Sub

' クリア: the input and the current result/candidates only. The ledger, the
' processed flags and the two timing figures stay.
Public Sub Rdv3UiClearResult()
    Dim ws As Object
    Set ws = UiWs()
    ws.Range(RDV3_C_INPUT).Value2 = ""
    ws.Range(RDV3_C_KEY).Value2 = ""
    ws.Range(RDV3_C_VERDICT).Value2 = ""
    ws.Range(RDV3_C_PROCSTATE).Value2 = ""
    ws.Range("B" & RDV3_REC_TOP & ":H" & (RDV3_REC_TOP + RDV3_FIELDS - 1)).ClearContents
    Rdv3UiClearCandidates
End Sub
