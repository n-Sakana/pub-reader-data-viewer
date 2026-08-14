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
' clsRdv2Idx (hash flavour) -- index method 1 of 2 on the VBA side.
'
' clsRdv2IdxDict.cls declares a class with the SAME name and the same four
' members. build\build_workbooks2.ps1 imports exactly one of them into a given
' workbook, so modRdv2Engine says New clsRdv2Idx either way and nothing in the
' engine knows which structure it got.
'
' Structure: bucket chaining. head() holds the first row of a bucket and next()
' the row after it, so a key that owns several rows simply has several rows on
' its chain. One-to-many falls out of the design; nothing had to be added for
' it, and no row can overwrite another.
'
' The eight key bytes are copied into two Longs at build time (k0, k1). That is
' what makes a comparison two Long compares instead of eight byte compares, and
' it is also the honest counterpart of what the Dictionary flavour does when it
' turns each key into a String -- both keep their own copy of the key.
'
' Twin of src\v2\csharp\Rdv2IdxHash.cs.
'==============================================================================
Option Explicit

Private m_head() As Long
Private m_next() As Long
Private m_k0() As Long
Private m_k1() As Long
Private m_mask As Long
Private m_rows As Long
Private m_probes As Double

Public Property Get Kind() As String
    Kind = "hash"
End Property

Public Property Get Probes() As Double
    Probes = m_probes
End Property

' st(i) is the byte offset of data row i's key. s is ignored here; the
' Dictionary flavour needs it and the two must take the same arguments.
Public Sub Build(ByRef b() As Byte, ByRef s As String, ByRef st() As Long, ByVal nrows As Long)
    Dim i As Long, h As Long, idx As Long, p As Long, cap As Long

    m_rows = nrows
    cap = Rdv2Capacity(nrows)
    m_mask = cap - 1
    ReDim m_head(0 To cap - 1)
    ReDim m_next(0 To nrows - 1)
    ReDim m_k0(0 To nrows - 1)
    ReDim m_k1(0 To nrows - 1)
    m_probes = 0

    For i = 0 To nrows - 1
        p = st(i)
        ' --- Rdv2HashBytes, inlined ---
        h = b(p)
        h = ((h * 131) + b(p + 1)) And RDV2_HASHMASK
        h = ((h * 131) + b(p + 2)) And RDV2_HASHMASK
        h = ((h * 131) + b(p + 3)) And RDV2_HASHMASK
        h = ((h * 131) + b(p + 4)) And RDV2_HASHMASK
        h = ((h * 131) + b(p + 5)) And RDV2_HASHMASK
        h = ((h * 131) + b(p + 6)) And RDV2_HASHMASK
        h = ((h * 131) + b(p + 7)) And RDV2_HASHMASK
        h = (h Xor (h \ 1024)) And RDV2_HASHMASK
        ' --- end inline ---
        idx = h And m_mask
        ' 256& forces Long arithmetic: Byte * 256 is Integer arithmetic in VBA
        ' and overflows above 127
        m_k0(i) = b(p) + b(p + 1) * 256& + b(p + 2) * 65536 + b(p + 3) * 16777216
        m_k1(i) = b(p + 4) + b(p + 5) * 256& + b(p + 6) * 65536 + b(p + 7) * 16777216
        m_next(i) = m_head(idx)
        m_head(idx) = i + 1
    Next i
End Sub

' every row whose key equals the eight bytes at pb(p). Returns how many were
' found and fills result with the row numbers.
Public Function Find(ByRef pb() As Byte, ByRef ps As String, ByVal p As Long, ByRef result() As Long) As Long
    Dim h As Long, r As Long, row As Long, n As Long
    Dim q0 As Long, q1 As Long
    Dim steps As Double

    ' --- Rdv2HashBytes, inlined ---
    h = pb(p)
    h = ((h * 131) + pb(p + 1)) And RDV2_HASHMASK
    h = ((h * 131) + pb(p + 2)) And RDV2_HASHMASK
    h = ((h * 131) + pb(p + 3)) And RDV2_HASHMASK
    h = ((h * 131) + pb(p + 4)) And RDV2_HASHMASK
    h = ((h * 131) + pb(p + 5)) And RDV2_HASHMASK
    h = ((h * 131) + pb(p + 6)) And RDV2_HASHMASK
    h = ((h * 131) + pb(p + 7)) And RDV2_HASHMASK
    h = (h Xor (h \ 1024)) And RDV2_HASHMASK
    ' --- end inline ---
    q0 = pb(p) + pb(p + 1) * 256& + pb(p + 2) * 65536 + pb(p + 3) * 16777216
    q1 = pb(p + 4) + pb(p + 5) * 256& + pb(p + 6) * 65536 + pb(p + 7) * 16777216

    n = 0
    steps = 0
    r = m_head(h And m_mask)
    Do While r <> 0
        steps = steps + 1
        row = r - 1
        If m_k0(row) = q0 Then
            If m_k1(row) = q1 Then
                If n < RDV2_MAXCAND Then
                    result(n) = row
                End If
                n = n + 1
            End If
        End If
        r = m_next(row)
    Loop
    m_probes = m_probes + steps
    If n > RDV2_MAXCAND Then
        n = RDV2_MAXCAND
    End If
    Find = n
End Function
