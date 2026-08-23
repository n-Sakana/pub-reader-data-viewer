VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRdv2Idx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsRdv2Idx (Dictionary flavour) -- index method 2 of 2 on the VBA side.
'
' clsRdv2IdxHash.cls declares a class with the SAME name and the same four
' members. build\build_workbooks2.ps1 imports exactly one of them into a given
' workbook.
'
' The stock associative array: Scripting.Dictionary, late bound through
' CreateObject so the workbook needs no reference set and opens on a machine
' where the reference was never added. The value under each key is a
' Collection of row numbers -- the readable VBA counterpart of
' Dictionary<string, List<int>> - so a key that owns several rows keeps all of
' them and nothing is overwritten.
'
' What it costs, and what the comparison is really about: a key has to be a
' String. Every row makes one on the way in and every probe makes one on the
' way out, because the engine only ever has bytes at an offset. That is not a
' handicap invented for the benchmark; it is what using the stock dictionary on
' a byte buffer costs, and clsRdv2IdxHash exists to show the size of it.
'
' Twin of src\v2\csharp\Rdv2IdxDict.cs.
'==============================================================================
Option Explicit

Private m_d As Object
Private m_probes As Double

Public Property Get Kind() As String
    Kind = "dict"
End Property

Public Property Get Probes() As Double
    Probes = m_probes
End Property

' s is the file widened one byte to one character, so Mid$(s, st(i) + 1, 8) is
' the key of data row i. b() is ignored here; the hash flavour needs it and the
' two must take the same arguments.
Public Sub Build(ByRef b() As Byte, ByRef s As String, ByRef st() As Long, ByVal nrows As Long)
    Dim i As Long
    Dim k As String
    Dim col As Collection

    Set m_d = CreateObject("Scripting.Dictionary")
    m_d.CompareMode = 0                    ' binary: 8 digits, case is not a question
    m_probes = 0
    For i = 0 To nrows - 1
        k = Mid$(s, st(i) + 1, RDV2_KEYLEN)
        If m_d.Exists(k) Then
            m_d.Item(k).Add i
        Else
            Set col = New Collection
            col.Add i
            m_d.Add k, col
        End If
    Next i
End Sub

Public Function Find(ByRef pb() As Byte, ByRef ps As String, ByVal p As Long, ByRef result() As Long) As Long
    Dim k As String
    Dim n As Long, j As Long
    Dim col As Collection

    k = Mid$(ps, p + 1, RDV2_KEYLEN)
    If Not m_d.Exists(k) Then
        m_probes = m_probes + 1
        Find = 0
        Exit Function
    End If
    Set col = m_d.Item(k)
    n = col.Count
    m_probes = m_probes + n
    If n > RDV2_MAXCAND Then
        n = RDV2_MAXCAND
    End If
    For j = 1 To n
        result(j - 1) = col.Item(j)
    Next j
    Find = n
End Function
