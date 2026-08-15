Attribute VB_Name = "modRdv3Save"
'==============================================================================
' modRdv3Save -- the ways one processed flag can reach the ledger file, and the
' switch that decides which one the BE uses.
'
'   book  the shipped method. The BE keeps the ledger workbook open in its own
'         invisible Excel, writes the one cell and calls Workbook.Save.
'   ado   the workbook stays CLOSED. An ADO/ACE connection updates the one cell.
'         Several shapes of that are available so they can be measured against
'         each other: a connection per record, a held connection, a prepared
'         command, and an explicit transaction.
'   zip   the workbook stays CLOSED and no Excel, no external program and no
'         Win32 is involved: modRdv3Zip changes ONE BYTE of the original sheet
'         XML in place and leaves every other byte of the package alone.
'
' The distribution default is NOT changed here: the BE uses book unless a
' "<ledger>.savemethod" file sits next to the ledger and names another one, and
' no distribution carries that file.
'
' NO Declare AND NO Shell. ADO and ACE are COM.
'==============================================================================
Option Explicit

Public Const RDV3_SAVE_BOOK As Long = 0
Public Const RDV3_SAVE_ADO As Long = 1
Public Const RDV3_SAVE_ZIP As Long = 2

Private m_adoProvider As String
Private m_lastNote As String
Private m_stageMs As String

' held-connection shapes, for the comparison
Private m_cn As Object
Private m_cmd As Object
Private m_adoMode As String
Private m_adoPath As String
Private m_inTx As Boolean

'==============================================================================
' method selection
'==============================================================================
Public Function Rdv3SaveMethodName(ByVal id As Long) As String
    Select Case id
        Case RDV3_SAVE_ADO
            Rdv3SaveMethodName = "ado"
        Case RDV3_SAVE_ZIP
            Rdv3SaveMethodName = "zip"
        Case Else
            Rdv3SaveMethodName = "book"
    End Select
End Function

Public Function Rdv3SaveMethodId(ByVal nm As String) As Long
    Select Case LCase$(Trim$(nm))
        Case "ado"
            Rdv3SaveMethodId = RDV3_SAVE_ADO
        Case "zip"
            Rdv3SaveMethodId = RDV3_SAVE_ZIP
        Case Else
            Rdv3SaveMethodId = RDV3_SAVE_BOOK
    End Select
End Function

' "<ledger>.savemethod" next to the ledger, holding book / ado / zip. Absent is
' the shipped case and means book: no distribution carries this file.
Public Function Rdv3SaveMethodFromFile(ByVal ledgerPath As String) As Long
    Dim p As String
    Dim f As Integer
    Dim s As String
    Rdv3SaveMethodFromFile = RDV3_SAVE_BOOK
    p = ledgerPath & ".savemethod"
    If Dir$(p) = "" Then Exit Function
    On Error GoTo Fail
    f = FreeFile
    Open p For Input As #f
    If Not EOF(f) Then Line Input #f, s
    Close #f
    Rdv3SaveMethodFromFile = Rdv3SaveMethodId(s)
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

Public Function Rdv3SaveLastNote() As String
    Rdv3SaveLastNote = m_lastNote
End Function

Public Function Rdv3SaveStageMs() As String
    Rdv3SaveStageMs = m_stageMs
End Function

Public Sub Rdv3SaveReset()
    On Error Resume Next
    If Not m_cn Is Nothing Then m_cn.Close
    Set m_cmd = Nothing
    Set m_cn = Nothing
    m_adoMode = ""
    m_inTx = False
    On Error GoTo 0
End Sub

'==============================================================================
' ADO / ACE
'==============================================================================
' ACE is the only Excel ISAM there is; 16.0 and 12.0 are the same driver under
' two names. Trying both is not a fallback to another METHOD -- if neither is
' installed the method is reported unavailable, never silently swapped.
Private Function AdoConnString(ByVal ledgerPath As String) As String
    Dim tryProv As Variant
    Dim p As Variant
    Dim cn As Object
    Dim lastErr As String

    If Len(m_adoProvider) > 0 Then
        AdoConnString = ConnFor(m_adoProvider, ledgerPath)
        Exit Function
    End If
    tryProv = Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")
    For Each p In tryProv
        On Error Resume Next
        Err.Clear
        Set cn = CreateObject("ADODB.Connection")
        cn.Open ConnFor(CStr(p), ledgerPath)
        If Err.Number = 0 Then
            cn.Close
            Set cn = Nothing
            On Error GoTo 0
            m_adoProvider = CStr(p)
            AdoConnString = ConnFor(m_adoProvider, ledgerPath)
            Exit Function
        End If
        lastErr = CStr(p) & " -> " & Err.Number & ": " & Err.Description
        Set cn = Nothing
        On Error GoTo 0
    Next p
    Err.Raise vbObjectError + 2001, "modRdv3Save", _
        "ADO/ACE の Excel プロバイダを開けません (" & lastErr & ")"
End Function

Private Function ConnFor(ByVal prov As String, ByVal ledgerPath As String) As String
    ConnFor = "Provider=" & prov & ";Data Source=" & ledgerPath & _
              ";Extended Properties=""Excel 12.0 Xml;HDR=NO;"";"
End Function

Public Function Rdv3SaveAdoProbe(ByVal ledgerPath As String) As String
    Dim s As String
    On Error GoTo Fail
    s = AdoConnString(ledgerPath)
    Rdv3SaveAdoProbe = "ok provider=" & m_adoProvider
    Exit Function
Fail:
    Rdv3SaveAdoProbe = "ERR " & Err.Number & ": " & Err.Description
End Function

' one cell, addressed directly: no scan over 100,000 rows, and no type guessing
' on any column but this one. The processed column is a real boolean cell in the
' ledger (Excel writes <c t="b">), so ACE types it boolean and a quoted 'TRUE'
' is a type mismatch -- measured.
Private Function UpdateSql(ByVal sheetRow As Long, ByVal newVal As Boolean) As String
    UpdateSql = "UPDATE [" & RDV3_LB_SHEET & "$A" & CStr(sheetRow) & ":A" & CStr(sheetRow) & _
                "] SET F1=" & IIf(newVal, "TRUE", "FALSE")
End Function

' the shipped shape: a connection per record, closed before the call returns
Public Function Rdv3SaveOneAdo(ByVal ledgerPath As String, ByVal sheetRow As Long, _
                               ByVal newVal As Boolean, ByRef errMsg As String) As Boolean
    Dim cn As Object
    Dim opened As Boolean

    errMsg = ""
    On Error GoTo Fail
    Set cn = CreateObject("ADODB.Connection")
    cn.Open AdoConnString(ledgerPath)
    opened = True
    cn.Execute UpdateSql(sheetRow, newVal)
    cn.Close                     ' closing the connection is what commits it
    opened = False
    Set cn = Nothing
    Rdv3SaveOneAdo = True
    Exit Function
Fail:
    errMsg = "ADO 保存に失敗 " & Err.Number & ": " & Err.Description & _
        IIf(Len(m_adoProvider) > 0, " (provider=" & m_adoProvider & ")", " (provider 未決定)")
    On Error Resume Next
    If opened Then cn.Close
    Set cn = Nothing
End Function

' the shapes worth measuring against that one. mode:
'   peritem   a new connection for every record (what ships)
'   hold      one connection kept open across records
'   prepared  one connection, one prepared command with a parameter
'   tx        one connection inside an explicit transaction
Public Function Rdv3SaveAdoBegin(ByVal ledgerPath As String, ByVal mode As String, _
                                 ByRef errMsg As String) As Boolean
    errMsg = ""
    Rdv3SaveReset
    m_adoPath = ledgerPath
    m_adoMode = LCase$(Trim$(mode))
    On Error GoTo Fail
    If m_adoMode = "peritem" Then
        AdoConnString ledgerPath      ' settle the provider outside the records
        Rdv3SaveAdoBegin = True
        Exit Function
    End If
    Set m_cn = CreateObject("ADODB.Connection")
    m_cn.Open AdoConnString(ledgerPath)
    If m_adoMode = "prepared" Then
        Set m_cmd = CreateObject("ADODB.Command")
        Set m_cmd.ActiveConnection = m_cn
        m_cmd.Prepared = True
    ElseIf m_adoMode = "tx" Then
        m_cn.BeginTrans
        m_inTx = True
    End If
    Rdv3SaveAdoBegin = True
    Exit Function
Fail:
    errMsg = "ADO begin (" & mode & ") 失敗 " & Err.Number & ": " & Err.Description
    Rdv3SaveReset
End Function

Public Function Rdv3SaveAdoStep(ByVal sheetRow As Long, ByVal newVal As Boolean, _
                                ByRef errMsg As String) As Boolean
    errMsg = ""
    On Error GoTo Fail
    Select Case m_adoMode
        Case "peritem"
            Rdv3SaveAdoStep = Rdv3SaveOneAdo(m_adoPath, sheetRow, newVal, errMsg)
        Case "hold", "tx"
            m_cn.Execute UpdateSql(sheetRow, newVal)
            Rdv3SaveAdoStep = True
        Case "prepared"
            m_cmd.CommandText = UpdateSql(sheetRow, newVal)
            m_cmd.Execute
            Rdv3SaveAdoStep = True
        Case Else
            errMsg = "no ADO run is open"
    End Select
    Exit Function
Fail:
    errMsg = "ADO step 失敗 " & Err.Number & ": " & Err.Description
End Function

Public Function Rdv3SaveAdoEnd(ByRef errMsg As String) As Boolean
    errMsg = ""
    On Error GoTo Fail
    If m_inTx Then
        m_cn.CommitTrans
        m_inTx = False
    End If
    If Not m_cn Is Nothing Then m_cn.Close
    Set m_cmd = Nothing
    Set m_cn = Nothing
    m_adoMode = ""
    Rdv3SaveAdoEnd = True
    Exit Function
Fail:
    errMsg = "ADO end 失敗 " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Set m_cmd = Nothing
    Set m_cn = Nothing
    m_adoMode = ""
End Function

'==============================================================================
' zip: one byte of the original sheet XML, nothing else (modRdv3Zip)
'==============================================================================
Public Function Rdv3SaveOneZip(ByVal ledgerPath As String, ByVal sheetRow As Long, _
                               ByVal rowsTotal As Long, ByVal newVal As Boolean, _
                               ByRef errMsg As String) As Boolean
    Dim tmp As String
    Dim info As String

    errMsg = ""
    m_lastNote = ""
    m_stageMs = ""
    tmp = ledgerPath & ".ziptmp"
    On Error GoTo Fail
    If Len(Dir$(tmp)) > 0 Then Kill tmp
    If Not Rdv3ZipSetProcessed(ledgerPath, tmp, sheetRow, rowsTotal, newVal, info, errMsg) Then
        Exit Function
    End If
    ' the same rule the channel and the sidecar use: park the live file by
    ' rename, put the new one in place, and only then drop the parked copy --
    ' every failure leaves both generations on disk
    If Not Rdv3ChReplaceFile(tmp, ledgerPath) Then
        errMsg = "台帳の置換に失敗 (err " & CStr(Rdv3ChLastReplaceErr()) & ")"
        Exit Function
    End If
    m_lastNote = info
    m_stageMs = info
    Rdv3SaveOneZip = True
    Exit Function
Fail:
    errMsg = "ZIP 保存に失敗 " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Len(tmp) > 0 Then
        If Len(Dir$(tmp)) > 0 Then Kill tmp
    End If
End Function

'==============================================================================
' build-time probe: forces this module and its dependencies to compile
'==============================================================================
Public Function Rdv3SaveSelfTest() As String
    Rdv3SaveSelfTest = "save ok methods=" & Rdv3SaveMethodName(RDV3_SAVE_BOOK) & "," & _
        Rdv3SaveMethodName(RDV3_SAVE_ADO) & "," & Rdv3SaveMethodName(RDV3_SAVE_ZIP) & _
        " | " & Rdv3ZipSelfTest()
End Function
