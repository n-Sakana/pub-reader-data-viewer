Attribute VB_Name = "modRdv3Zip"
'==============================================================================
' modRdv3Zip -- DEFLATE, in VBA, for the smallest possible edit of a workbook
' that is already on disk.
'
' The job this exists for: one cell of a 100,000-row sheet holds "0" or "1",
' and the sheet part is a 128 MB XML compressed to 18.5 MB. Nothing else in the
' package may change. So the original XML is the authority -- it is never
' regenerated -- and the edit is one byte in it.
'
' A deflate stream cannot be seeked: block boundaries only exist once you have
' decoded everything before them. But finding the target does NOT require
' rebuilding 128 MB of output. Two observations do the work:
'
'   1. SKIPPING is far cheaper than inflating. To walk past a block you must
'      decode every symbol, but a match only has to ADD its length to a
'      counter -- the bytes it would copy are never written. Measured on this
'      ledger's sheet part: 1,988 ms to skip all of it, 5,723 ms to
'      materialise all of it.
'   2. The window only matters near the target. Matches reach at most 32 KB
'      back, so the walk can skip almost to the row and only then start
'      writing bytes.
'
' THE TRAP THIS DESIGN EXISTS TO AVOID
' Changing one byte of the OUTPUT is not one byte of the STREAM. Every later
' row's "<c r=... t="b"><v>0</v></c>" is encoded as a match that copies from
' the row before it, so a naive patch of the target block changed the flag of
' 58,033 following rows -- measured, not feared. A block can only reach 32 KB
' back, so every block that starts within 32 KB after the target is re-emitted
' too, and inside them any match that would copy the changed byte is split so
' that byte goes out as a literal with its ORIGINAL value. Blocks that start
' further on cannot see it and are copied bit for bit.
'
' Each re-emitted block keeps its OWN Huffman header, copied verbatim, and its
' own symbols; only the symbols that touch the changed byte differ. The
' compressed size therefore moves by a few bits.
'
' NO Declare AND NO Shell.
'==============================================================================
Option Explicit

Public Const RDV3_Z_STORED As Long = 0
Public Const RDV3_Z_FIXED As Long = 1
Public Const RDV3_Z_DYNAMIC As Long = 2

Private Const WIN_SIZE As Long = 32768
Private Const WIN_MASK As Long = 32767
Private Const FAST_BITS As Long = 9
Private Const FAST_SIZE As Long = 512
Private Const FAST_MASK As Long = 511
Private Const MAX_DIST As Long = 32768

' ---- the compressed input and the bit reader --------------------------------
Private m_buf() As Byte
Private m_bufLen As Long
Private m_bytePos As Long
Private m_hold As Long
Private m_cnt As Long

' ---- decoded Huffman tables (literal/length and distance) -------------------
Private m_lCount(0 To 15) As Long
Private m_lSym() As Long
Private m_lFast(0 To FAST_SIZE - 1) As Long      ' sym * 16 + len, 0 = miss
Private m_dCount(0 To 15) As Long
Private m_dSym() As Long
Private m_dFast(0 To FAST_SIZE - 1) As Long
Private m_lLen(0 To 287) As Long                 ' code lengths, for re-emitting
Private m_dLen(0 To 29) As Long
Private m_lCode(0 To 287) As Long                ' reversed codes, for writing
Private m_dCode(0 To 29) As Long

' ---- output bookkeeping -----------------------------------------------------
Private m_out As Long                  ' uncompressed bytes produced so far
Private m_win() As Byte                ' 32 KB circular window
Private m_materialise As Boolean
Private m_watch As Boolean             ' are we near enough to look at bytes?
Private m_blocks As Long

' ---- pattern search ---------------------------------------------------------
Private m_pat() As Byte
Private m_patLen As Long
Private m_patHit As Long               ' output offset of the match, -1 = none
Private m_patAt As Long
Private m_pTarget As Long              ' the byte to change, -1 until found
Private m_newByte As Byte
Private m_oldByte As Byte

' ---- the span that has to be re-emitted -------------------------------------
Private m_spanBuf() As Byte            ' original output bytes of that span
Private m_spanStart As Long            ' output offset of m_spanBuf(0)
Private m_spanN As Long
Private m_spanOn As Boolean

Private m_capKind() As Byte            ' 0 = literal, 1 = match
Private m_capA() As Long               ' literal byte / match length
Private m_capB() As Long               ' 0 / match distance
Private m_capOut() As Long             ' output offset this symbol starts at
Private m_capN As Long
Private m_capture As Boolean

Private m_spliceStartBit As Long       ' first bit of the first re-emitted block
Private m_spliceEndBit As Long         ' first bit after the last one
Private m_patching As Boolean
Private m_alignBlocks As Long          ' blocks re-emitted purely to land on a byte
Private m_spanToEnd As Boolean         ' the span ran to the end of the stream

' ---- the bit writer ---------------------------------------------------------
Private m_wBuf() As Byte
Private m_wAt As Long
Private m_wHold As Long
Private m_wCnt As Long
Private m_wBits As Long

Private m_pow2(0 To 31) As Long
Private m_ready As Boolean
Private m_err As String
Private m_stage As String              ' where a raised error came from
Private m_info As String               ' the numbers the caller reports
Private m_wrInfo As String             ' where the write phase spent its time

' ---- CRC-32 (reflected, 0xEDB88320) and its one-byte-difference update ------
Private m_crcTbl(0 To 255) As Long
Private m_crcReady As Boolean
Private m_crcM(0 To 31) As Long
Private m_crcMReady As Boolean

Private Const LBASE As String = "3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258"
Private Const LEXT As String = "0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0"
Private Const DBASE As String = "1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577"
Private Const DEXT As String = "0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13"
Private Const CLORD As String = "16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15"

Private m_lbase(0 To 28) As Long
Private m_lext(0 To 28) As Long
Private m_dbase(0 To 29) As Long
Private m_dext(0 To 29) As Long
Private m_clord(0 To 18) As Long

'==============================================================================
Public Sub Rdv3ZipInit()
    Dim i As Long
    Dim p As Variant
    If m_ready Then Exit Sub
    For i = 0 To 30
        m_pow2(i) = 2 ^ i
    Next i
    m_pow2(31) = &H80000000
    p = Split(LBASE, ",")
    For i = 0 To 28
        m_lbase(i) = CLng(p(i))
    Next i
    p = Split(LEXT, ",")
    For i = 0 To 28
        m_lext(i) = CLng(p(i))
    Next i
    p = Split(DBASE, ",")
    For i = 0 To 29
        m_dbase(i) = CLng(p(i))
    Next i
    p = Split(DEXT, ",")
    For i = 0 To 29
        m_dext(i) = CLng(p(i))
    Next i
    p = Split(CLORD, ",")
    For i = 0 To 18
        m_clord(i) = CLng(p(i))
    Next i
    ReDim m_win(0 To WIN_SIZE - 1)
    m_ready = True
End Sub

Public Function Rdv3ZipLastErr() As String
    Rdv3ZipLastErr = m_err
End Function

Public Function Rdv3ZipLastInfo() As String
    Rdv3ZipLastInfo = m_info
End Function

Public Function Rdv3ZipOutCount() As Long
    Rdv3ZipOutCount = m_out
End Function

Public Function Rdv3ZipBlocks() As Long
    Rdv3ZipBlocks = m_blocks
End Function

Public Function Rdv3ZipPatternAt() As Long
    Rdv3ZipPatternAt = m_patHit
End Function

'------------------------------------------------------------------------------
' the bit reader
'------------------------------------------------------------------------------
Private Sub BitStart(ByVal atByte As Long)
    m_bytePos = atByte
    m_hold = 0
    m_cnt = 0
End Sub

Private Function BitPos() As Long
    BitPos = m_bytePos * 8& - m_cnt
End Function

Private Sub Refill()
    Do While m_cnt <= 23
        If m_bytePos >= m_bufLen Then
            m_cnt = m_cnt + 8
        Else
            m_hold = m_hold Or (CLng(m_buf(m_bytePos)) * m_pow2(m_cnt))
            m_bytePos = m_bytePos + 1
            m_cnt = m_cnt + 8
        End If
    Loop
End Sub

Private Function Bits(ByVal n As Long) As Long
    If n = 0 Then Exit Function
    If m_cnt < n Then Refill
    Bits = m_hold And (m_pow2(n) - 1)
    m_hold = m_hold \ m_pow2(n)
    m_cnt = m_cnt - n
End Function

'------------------------------------------------------------------------------
' canonical Huffman
'------------------------------------------------------------------------------
Private Function BuildTable(ByRef lengths() As Long, ByVal n As Long, _
                            ByRef cnt() As Long, ByRef sym() As Long, _
                            ByRef fast() As Long) As Boolean
    Dim i As Long, len_ As Long, code As Long
    Dim offs(0 To 15) As Long
    Dim nextCode(0 To 15) As Long
    Dim rev As Long, step_ As Long, idx As Long, c As Long

    For i = 0 To 15
        cnt(i) = 0
    Next i
    For i = 0 To n - 1
        cnt(lengths(i)) = cnt(lengths(i)) + 1
    Next i
    cnt(0) = 0
    offs(1) = 0
    For i = 1 To 14
        offs(i + 1) = offs(i) + cnt(i)
    Next i
    ReDim sym(0 To n)
    For i = 0 To n - 1
        If lengths(i) <> 0 Then
            sym(offs(lengths(i))) = i
            offs(lengths(i)) = offs(lengths(i)) + 1
        End If
    Next i

    For i = 0 To FAST_SIZE - 1
        fast(i) = 0
    Next i
    code = 0
    For len_ = 1 To 15
        nextCode(len_) = code
        code = (code + cnt(len_)) * 2
    Next len_
    For i = 0 To n - 1
        len_ = lengths(i)
        If len_ > 0 Then
            c = nextCode(len_)
            nextCode(len_) = nextCode(len_) + 1
            If len_ <= FAST_BITS Then
                rev = RevBits(c, len_)
                step_ = m_pow2(len_)
                idx = rev
                Do While idx < FAST_SIZE
                    fast(idx) = i * 16& + len_
                    idx = idx + step_
                Loop
            End If
        End If
    Next i
    BuildTable = True
End Function

' Huffman codes are written MSB first, but the bit stream hands them over LSB
' first, so the root table is indexed by the REVERSED code. Taking the input
' bits from the bottom up and appending them at the top is the reversal;
' walking them from the top down would just rebuild the same number.
Private Function RevBits(ByVal v As Long, ByVal n As Long) As Long
    Dim i As Long, r As Long
    For i = 0 To n - 1
        r = r * 2
        If (v And m_pow2(i)) <> 0 Then r = r Or 1
    Next i
    RevBits = r
End Function

Private Function DecodeSlow(ByRef cnt() As Long, ByRef sym() As Long) As Long
    Dim code As Long, first As Long, index As Long, len_ As Long, count As Long
    For len_ = 1 To 15
        code = code Or Bits(1)
        count = cnt(len_)
        If code - count < first Then
            DecodeSlow = sym(index + (code - first))
            Exit Function
        End If
        index = index + count
        first = (first + count) * 2
        code = code * 2
    Next len_
    DecodeSlow = -1
End Function

Private Function DecodeWith(ByRef cnt() As Long, ByRef sym() As Long, ByRef fast() As Long) As Long
    Dim e As Long
    If m_cnt < 15 Then Refill
    e = fast(m_hold And FAST_MASK)
    If e <> 0 Then
        DecodeWith = e \ 16&
        m_hold = m_hold \ m_pow2(e And 15&)
        m_cnt = m_cnt - (e And 15&)
    Else
        DecodeWith = DecodeSlow(cnt, sym)
    End If
End Function

Private Sub BuildFixed()
    Dim l(0 To 287) As Long
    Dim d(0 To 29) As Long
    Dim i As Long
    For i = 0 To 143
        l(i) = 8
    Next i
    For i = 144 To 255
        l(i) = 9
    Next i
    For i = 256 To 279
        l(i) = 7
    Next i
    For i = 280 To 287
        l(i) = 8
    Next i
    For i = 0 To 29
        d(i) = 5
    Next i
    For i = 0 To 287
        m_lLen(i) = l(i)
    Next i
    For i = 0 To 29
        m_dLen(i) = d(i)
    Next i
    BuildTable l, 288, m_lCount, m_lSym, m_lFast
    BuildTable d, 30, m_dCount, m_dSym, m_dFast
End Sub

Private Function BuildDynamic() As Boolean
    Dim hlit As Long, hdist As Long, hclen As Long
    Dim cl(0 To 18) As Long
    Dim clCount(0 To 15) As Long
    Dim clSym() As Long
    Dim clFast(0 To FAST_SIZE - 1) As Long
    Dim lens(0 To 319) As Long
    Dim i As Long, idx As Long, s As Long, r As Long, prev As Long
    Dim ll(0 To 287) As Long
    Dim dd(0 To 29) As Long

    m_stage = "dyn-header"
    hlit = Bits(5) + 257
    hdist = Bits(5) + 1
    hclen = Bits(4) + 4
    For i = 0 To 18
        cl(i) = 0
    Next i
    For i = 0 To hclen - 1
        cl(m_clord(i)) = Bits(3)
    Next i
    BuildTable cl, 19, clCount, clSym, clFast

    m_stage = "dyn-lengths"
    idx = 0
    Do While idx < hlit + hdist
        s = DecodeWith(clCount, clSym, clFast)
        If s < 0 Or s > 18 Then
            m_err = "bad code-length code (" & CStr(s) & ") at idx=" & CStr(idx)
            Exit Function
        End If
        If s < 16 Then
            lens(idx) = s
            idx = idx + 1
        Else
            If s = 16 Then
                If idx = 0 Then
                    m_err = "repeat with no previous length"
                    Exit Function
                End If
                prev = lens(idx - 1)
                r = 3 + Bits(2)
            ElseIf s = 17 Then
                prev = 0
                r = 3 + Bits(3)
            Else
                prev = 0
                r = 11 + Bits(7)
            End If
            If idx + r > hlit + hdist Then
                m_err = "code-length run overruns the table"
                Exit Function
            End If
            Do While r > 0
                lens(idx) = prev
                idx = idx + 1
                r = r - 1
            Loop
        End If
    Loop

    m_stage = "dyn-build"
    For i = 0 To 287
        If i < hlit Then
            ll(i) = lens(i)
        Else
            ll(i) = 0
        End If
        m_lLen(i) = ll(i)
    Next i
    For i = 0 To 29
        If i < hdist Then
            dd(i) = lens(hlit + i)
        Else
            dd(i) = 0
        End If
        m_dLen(i) = dd(i)
    Next i
    BuildTable ll, 288, m_lCount, m_lSym, m_lFast
    BuildTable dd, 30, m_dCount, m_dSym, m_dFast
    BuildDynamic = True
End Function

'==============================================================================
' the walk. One pass: skip, then materialise, then re-emit the affected blocks.
'==============================================================================
' watchFrom: the output offset from which the bytes are examined. The window is
' maintained from the very start regardless -- that is what correctness needs.
Public Function Rdv3ZipWalk(ByRef comp() As Byte, ByVal compLen As Long, _
                            ByVal pattern As String, ByVal watchFrom As Long, _
                            ByVal patching As Boolean, ByVal newByte As Byte) As Boolean
    Dim last As Boolean
    Dim btype As Long
    Dim i As Long
    Dim n As Long
    Dim blkBit As Long
    Dim blkOut As Long
    Dim dataBit As Long
    Dim endBit As Long
    Dim storedAt As Long
    Dim pad As Long

    Rdv3ZipInit
    m_err = ""
    m_stage = "start"
    m_buf = comp
    m_bufLen = compLen
    m_out = 0
    m_blocks = 0
    m_patHit = -1
    m_patAt = 0
    m_pTarget = -1
    m_newByte = newByte
    m_patching = patching
    m_materialise = True
    m_watch = (watchFrom <= 0)
    m_spliceStartBit = -1
    m_spliceEndBit = -1
    m_alignBlocks = 0
    m_spanToEnd = False
    m_spanOn = False
    m_spanN = 0
    m_capN = 0
    m_capture = False
    If Len(pattern) > 0 Then
        m_pat = StrConv(pattern, vbFromUnicode)
        m_patLen = UBound(m_pat) + 1
    Else
        m_patLen = 0
    End If
    If patching Then
        WStart 262144
        ReDim m_capKind(0 To 65535)
        ReDim m_capA(0 To 65535)
        ReDim m_capB(0 To 65535)
        ReDim m_capOut(0 To 65535)
        ReDim m_spanBuf(0 To 393215)
    End If
    BitStart 0

    On Error GoTo Fail
    Do While Not last
        blkBit = BitPos()
        blkOut = m_out
        ' Two things have to be true before the rest of the stream can be copied
        ' as it is. A block that starts more than a window past the target
        ' cannot see the changed byte -- and the rest has to resume on a BYTE
        ' boundary. Deflate is a bit stream, but a stored block inside it is
        ' byte-aligned, so resuming at a shifted alignment turns its LEN/NLEN
        ' pair into nonsense: measured, the third patch of a file produced
        ' "Block length does not match with its complement" while the first two
        ' (whose spans happened to come out a whole number of bytes) were fine.
        ' Re-emitting more blocks cannot fix that -- an untouched block comes
        ' back bit for bit -- so the span is extended to a block that STARTS on
        ' a byte, and this side is padded to a byte with an empty stored block.
        ' The rest of the stream is then a straight byte copy.
        If m_pTarget >= 0 And blkOut > m_pTarget + MAX_DIST Then
            If (blkBit Mod 8&) = 0 Then
                m_spliceEndBit = blkBit
                Exit Do
            End If
            m_alignBlocks = m_alignBlocks + 1
            If m_alignBlocks > 256 Then
                m_err = "no byte-aligned block boundary within 256 blocks of the target"
                Exit Function
            End If
        End If
        m_blocks = m_blocks + 1
        m_capture = (patching And m_watch)
        m_capN = 0
        If patching And m_watch And m_pTarget < 0 Then
            ' until the target is found, only this block's bytes can matter
            m_spanStart = blkOut
            m_spanN = 0
            m_spanOn = True
        End If
        last = (Bits(1) = 1)
        btype = Bits(2)
        If btype = RDV3_Z_STORED Then
            ' Skip to the next byte boundary. The reader runs up to 31 bits
            ' ahead, so the whole bytes it has buffered must be handed back
            ' first -- dropping them outright would skip up to three bytes.
            m_stage = "stored"
            m_bytePos = m_bytePos - (m_cnt \ 8)
            m_hold = 0
            m_cnt = 0
            ' the stream can end on a stored block whose header is the last
            ' thing in it; anything past that is padding inside the entry
            If m_bytePos + 4 > m_bufLen Then
                If m_pTarget >= 0 And m_spliceEndBit < 0 Then
                    m_spliceEndBit = compLen * 8&
                    m_spanToEnd = True
                End If
                Exit Do
            End If
            n = CLng(m_buf(m_bytePos)) + CLng(m_buf(m_bytePos + 1)) * 256&
            If m_bytePos + 4 + n > m_bufLen Then
                m_err = "stored block runs past the end of the stream"
                Exit Function
            End If
            m_bytePos = m_bytePos + 4
            storedAt = m_bytePos
            For i = 1 To n
                EmitByte m_buf(m_bytePos)
                m_bytePos = m_bytePos + 1
            Next i
            endBit = BitPos()
            If m_pTarget >= 0 Then
                If m_pTarget >= blkOut And m_pTarget < m_out Then
                    m_err = "the target byte sits in a stored block; not supported"
                    Exit Function
                End If
                ' A stored block is byte-aligned by definition, so it is written
                ' out fresh rather than copied bit for bit: whatever alignment
                ' this side is on, the block itself pads to the next byte.
                If m_spliceStartBit < 0 Then m_spliceStartBit = blkBit
                WBits IIf(last, 1, 0), 1
                WBits 0, 2
                pad = (8& - ((m_spliceStartBit + m_wBits) Mod 8&)) Mod 8&
                If pad > 0 Then WBits 0, pad
                WBits n And &HFFFF&, 16
                WBits (Not n) And &HFFFF&, 16
                For i = 0 To n - 1
                    WBits CLng(m_buf(storedAt + i)), 8
                Next i
            End If
        Else
            If btype = RDV3_Z_FIXED Then
                BuildFixed
            ElseIf btype = RDV3_Z_DYNAMIC Then
                If Not BuildDynamic() Then GoTo Fail
            Else
                m_err = "invalid deflate block type"
                Exit Function
            End If
            dataBit = BitPos()
            If Not RunBlock() Then GoTo Fail
            endBit = BitPos()
            If patching And m_pTarget >= 0 Then
                If m_spliceStartBit < 0 Then m_spliceStartBit = blkBit
                If Not EmitBlock(comp, blkBit, dataBit, m_err) Then Exit Function
            End If
        End If
        ' the span reached the end: there is no rest to copy, so the alignment
        ' the tail would have needed does not apply
        If last And m_pTarget >= 0 And m_spliceEndBit < 0 Then
            m_spliceEndBit = compLen * 8&
            m_spanToEnd = True
        End If
        If Not m_watch Then
            If m_out >= watchFrom Then m_watch = True
        End If
    Loop
    Rdv3ZipWalk = True
    Exit Function
Fail:
    If Len(m_err) = 0 Then
        m_err = "inflate error " & Err.Number & ": " & Err.Description & _
            " at " & m_stage & " blk=" & CStr(m_blocks) & " out=" & CStr(m_out)
    End If
End Function

' One Huffman block.
'
' The bit reader, both Huffman lookups and the window write are INLINED here.
' A call per symbol is the dominant cost of walking 16 million of them in VBA:
' measured on this ledger's 128 MB sheet part, the same loop with
' DecodeLit/DecodeDist/Bits as calls took 3,121 ms to skip and 10,852 ms to
' materialise; inlined it takes 1,988 ms and 5,723 ms.
Private Function RunBlock() As Boolean
    Dim s As Long
    Dim e As Long
    Dim k As Long
    Dim ln As Long
    Dim ds As Long
    Dim dv As Long
    Dim i As Long
    Dim src As Long
    Dim dst As Long
    Dim b As Byte

    RunBlock = True
    m_stage = "block-body"
    Do
        If m_cnt < 15 Then
            Do While m_cnt <= 23
                If m_bytePos >= m_bufLen Then
                    m_cnt = m_cnt + 8
                Else
                    m_hold = m_hold Or (CLng(m_buf(m_bytePos)) * m_pow2(m_cnt))
                    m_bytePos = m_bytePos + 1
                    m_cnt = m_cnt + 8
                End If
            Loop
        End If
        e = m_lFast(m_hold And FAST_MASK)
        If e <> 0 Then
            s = e \ 16&
            k = e And 15&
            m_hold = m_hold \ m_pow2(k)
            m_cnt = m_cnt - k
        Else
            s = DecodeSlow(m_lCount, m_lSym)
            If s < 0 Then
                m_err = "bad literal code"
                RunBlock = False
                Exit Function
            End If
        End If

        If s < 256 Then
            If m_capture Then CapSym 0, s, 0
            If m_materialise Then
                b = CByte(s)
                m_win(m_out And WIN_MASK) = b
                If m_watch Then WatchByte b
                m_out = m_out + 1
            Else
                m_out = m_out + 1
            End If
        ElseIf s = 256 Then
            Exit Function
        Else
            s = s - 257
            If s > 28 Then
                m_err = "bad length code"
                RunBlock = False
                Exit Function
            End If
            ln = m_lbase(s)
            k = m_lext(s)
            If k > 0 Then
                If m_cnt < k Then Refill
                ln = ln + (m_hold And (m_pow2(k) - 1))
                m_hold = m_hold \ m_pow2(k)
                m_cnt = m_cnt - k
            End If
            If m_cnt < 15 Then
                Do While m_cnt <= 23
                    If m_bytePos >= m_bufLen Then
                        m_cnt = m_cnt + 8
                    Else
                        m_hold = m_hold Or (CLng(m_buf(m_bytePos)) * m_pow2(m_cnt))
                        m_bytePos = m_bytePos + 1
                        m_cnt = m_cnt + 8
                    End If
                Loop
            End If
            e = m_dFast(m_hold And FAST_MASK)
            If e <> 0 Then
                ds = e \ 16&
                k = e And 15&
                m_hold = m_hold \ m_pow2(k)
                m_cnt = m_cnt - k
            Else
                ds = DecodeSlow(m_dCount, m_dSym)
            End If
            If ds < 0 Or ds > 29 Then
                m_err = "bad distance code"
                RunBlock = False
                Exit Function
            End If
            dv = m_dbase(ds)
            k = m_dext(ds)
            If k > 0 Then
                If m_cnt < k Then Refill
                dv = dv + (m_hold And (m_pow2(k) - 1))
                m_hold = m_hold \ m_pow2(k)
                m_cnt = m_cnt - k
            End If

            If m_capture Then CapSym 1, ln, dv
            If m_materialise Then
                src = (m_out - dv) And WIN_MASK
                dst = m_out And WIN_MASK
                If m_watch Then
                    For i = 1 To ln
                        b = m_win(src)
                        m_win(dst) = b
                        WatchByte b
                        m_out = m_out + 1
                        dst = (dst + 1) And WIN_MASK
                        src = (src + 1) And WIN_MASK
                    Next i
                ElseIf src + ln <= WIN_SIZE And dst + ln <= WIN_SIZE Then
                    ' nothing is looking at the bytes yet and neither end of the
                    ' copy wraps: the cheapest this loop can be
                    For i = 0 To ln - 1
                        m_win(dst + i) = m_win(src + i)
                    Next i
                    m_out = m_out + ln
                Else
                    For i = 1 To ln
                        m_win(dst) = m_win(src)
                        dst = (dst + 1) And WIN_MASK
                        src = (src + 1) And WIN_MASK
                    Next i
                    m_out = m_out + ln
                End If
            Else
                ' the whole point: a skipped match costs one addition
                m_out = m_out + ln
            End If
        End If
    Loop
End Function

' Called only while the walk is close enough to the row to care about the
' bytes: it records the span and runs the pattern search. The window itself is
' maintained the whole way, because a match can reach 32 KB back and a window
' filled in from a cold start produces bytes that are simply wrong -- measured:
' materialising from any mid-stream offset never finds the row at all.
Private Sub WatchByte(ByVal b As Byte)
    If m_spanOn Then
        If m_spanN > UBound(m_spanBuf) Then
            ReDim Preserve m_spanBuf(0 To m_spanN + 262143)
        End If
        m_spanBuf(m_spanN) = b
        m_spanN = m_spanN + 1
    End If
    If m_patHit < 0 Then
        If b = m_pat(m_patAt) Then
            m_patAt = m_patAt + 1
            If m_patAt = m_patLen Then
                m_patHit = m_out - m_patLen + 1
                m_pTarget = m_patHit + m_patLen      ' the value byte follows the tag
            End If
        ElseIf b = m_pat(0) Then
            m_patAt = 1
        Else
            m_patAt = 0
        End If
    ElseIf m_out = m_pTarget Then
        m_oldByte = b
    End If
End Sub

Private Sub EmitByte(ByVal b As Byte)
    m_win(m_out And WIN_MASK) = b
    If m_watch Then WatchByte b
    m_out = m_out + 1
End Sub

Private Sub CapSym(ByVal kind As Byte, ByVal a As Long, ByVal b As Long)
    If m_capN > UBound(m_capKind) Then
        ReDim Preserve m_capKind(0 To m_capN + 65535)
        ReDim Preserve m_capA(0 To m_capN + 65535)
        ReDim Preserve m_capB(0 To m_capN + 65535)
        ReDim Preserve m_capOut(0 To m_capN + 65535)
    End If
    m_capKind(m_capN) = kind
    m_capA(m_capN) = a
    m_capB(m_capN) = b
    m_capOut(m_capN) = m_out
    m_capN = m_capN + 1
End Sub

'==============================================================================
' re-emitting an affected block: its own header, its own symbols, with only the
' ones that touch the changed byte adjusted
'==============================================================================
Private Function EmitBlock(ByRef comp() As Byte, ByVal blkBit As Long, _
                           ByVal dataBit As Long, ByRef errMsg As String) As Boolean
    Dim i As Long
    Dim q As Long
    Dim ln As Long
    Dim dv As Long

    m_stage = "reemit"
    If Not BuildCodes(errMsg) Then Exit Function
    CopyBits comp, blkBit, dataBit - blkBit          ' the Huffman header, verbatim

    For i = 0 To m_capN - 1
        q = m_capOut(i)
        If m_capKind(i) = 0 Then
            If q = m_pTarget Then
                If Not WLit(CLng(m_newByte), errMsg) Then Exit Function
            Else
                If Not WLit(m_capA(i), errMsg) Then Exit Function
            End If
        Else
            ln = m_capA(i)
            dv = m_capB(i)
            If Not EmitMatch(q, ln, dv, errMsg) Then Exit Function
        End If
    Next i
    If Not WLit(256, errMsg) Then Exit Function      ' end of block
    EmitBlock = True
End Function

' A match is left alone unless the changed byte is inside what it WRITES or
' inside what it READS. Either way the affected output positions are emitted as
' literals -- with the new value at the target itself, and with the ORIGINAL
' value where a copy of it would otherwise land.
Private Function EmitMatch(ByVal q As Long, ByVal ln As Long, ByVal dv As Long, _
                           ByRef errMsg As String) As Boolean
    Dim cut1 As Long
    Dim cut2 As Long
    Dim cuts(0 To 1) As Long
    Dim nCuts As Long
    Dim i As Long
    Dim at As Long
    Dim seg As Long
    Dim t As Long

    cut1 = -1
    cut2 = -1
    If m_pTarget >= q And m_pTarget < q + ln Then cut1 = m_pTarget
    If m_pTarget >= q - dv And m_pTarget < q - dv + ln Then cut2 = m_pTarget + dv
    If cut1 < 0 And cut2 < 0 Then
        EmitMatch = WMatch(ln, dv, errMsg)
        Exit Function
    End If

    nCuts = 0
    If cut1 >= 0 Then
        cuts(nCuts) = cut1
        nCuts = nCuts + 1
    End If
    If cut2 >= 0 And cut2 <> cut1 And cut2 >= q And cut2 < q + ln Then
        cuts(nCuts) = cut2
        nCuts = nCuts + 1
    End If
    If nCuts = 2 Then
        If cuts(0) > cuts(1) Then
            t = cuts(0)
            cuts(0) = cuts(1)
            cuts(1) = t
        End If
    End If

    at = q
    For i = 0 To nCuts - 1
        seg = cuts(i) - at
        If seg > 0 Then
            If Not EmitRun(at, seg, dv, errMsg) Then Exit Function
            at = at + seg
        End If
        If cuts(i) = m_pTarget Then
            If Not WLit(CLng(m_newByte), errMsg) Then Exit Function
        Else
            ' a copy of the target byte would land here: write what was there
            If Not WLit(CLng(m_oldByte), errMsg) Then Exit Function
        End If
        at = at + 1
    Next i
    seg = q + ln - at
    If seg > 0 Then
        If Not EmitRun(at, seg, dv, errMsg) Then Exit Function
    End If
    EmitMatch = True
End Function

' a piece of a split match: still a match if deflate allows the length, and
' literals out of the recorded span if it does not
Private Function EmitRun(ByVal at As Long, ByVal n As Long, ByVal dv As Long, _
                         ByRef errMsg As String) As Boolean
    Dim k As Long
    If n >= 3 Then
        EmitRun = WMatch(n, dv, errMsg)
        Exit Function
    End If
    For k = 0 To n - 1
        If Not WLit(CLng(SpanByte(at + k, errMsg)), errMsg) Then Exit Function
        If Len(errMsg) > 0 Then Exit Function
    Next k
    EmitRun = True
End Function

Private Function SpanByte(ByVal outPos As Long, ByRef errMsg As String) As Byte
    Dim idx As Long
    idx = outPos - m_spanStart
    If idx < 0 Or idx >= m_spanN Then
        errMsg = "output byte " & CStr(outPos) & " is outside the recorded span"
        Exit Function
    End If
    SpanByte = m_spanBuf(idx)
End Function

Private Function BuildCodes(ByRef errMsg As String) As Boolean
    Dim i As Long
    Dim len_ As Long
    Dim code As Long
    Dim cnt(0 To 15) As Long
    Dim nextCode(0 To 15) As Long

    For i = 0 To 15
        cnt(i) = 0
    Next i
    For i = 0 To 287
        cnt(m_lLen(i)) = cnt(m_lLen(i)) + 1
    Next i
    cnt(0) = 0
    code = 0
    For len_ = 1 To 15
        nextCode(len_) = code
        code = (code + cnt(len_)) * 2
    Next len_
    For i = 0 To 287
        If m_lLen(i) > 0 Then
            m_lCode(i) = RevBits(nextCode(m_lLen(i)), m_lLen(i))
            nextCode(m_lLen(i)) = nextCode(m_lLen(i)) + 1
        Else
            m_lCode(i) = -1
        End If
    Next i

    For i = 0 To 15
        cnt(i) = 0
    Next i
    For i = 0 To 29
        cnt(m_dLen(i)) = cnt(m_dLen(i)) + 1
    Next i
    cnt(0) = 0
    code = 0
    For len_ = 1 To 15
        nextCode(len_) = code
        code = (code + cnt(len_)) * 2
    Next len_
    For i = 0 To 29
        If m_dLen(i) > 0 Then
            m_dCode(i) = RevBits(nextCode(m_dLen(i)), m_dLen(i))
            nextCode(m_dLen(i)) = nextCode(m_dLen(i)) + 1
        Else
            m_dCode(i) = -1
        End If
    Next i
    BuildCodes = True
End Function

'------------------------------------------------------------------------------
' the bit writer
'------------------------------------------------------------------------------
Private Sub WStart(ByVal capBytes As Long)
    ReDim m_wBuf(0 To capBytes + 1024)
    m_wAt = 0
    m_wHold = 0
    m_wCnt = 0
    m_wBits = 0
End Sub

Private Sub WBits(ByVal v As Long, ByVal n As Long)
    If n = 0 Then Exit Sub
    m_wHold = m_wHold Or (v * m_pow2(m_wCnt))
    m_wCnt = m_wCnt + n
    m_wBits = m_wBits + n
    Do While m_wCnt >= 8
        If m_wAt > UBound(m_wBuf) Then ReDim Preserve m_wBuf(0 To m_wAt + 262143)
        m_wBuf(m_wAt) = m_wHold And &HFF&
        m_wAt = m_wAt + 1
        m_wHold = m_wHold \ 256&
        m_wCnt = m_wCnt - 8
    Loop
End Sub

Private Function WLit(ByVal sym As Long, ByRef errMsg As String) As Boolean
    If sym < 0 Or sym > 287 Then
        errMsg = "literal out of range: " & CStr(sym)
        Exit Function
    End If
    If m_lLen(sym) = 0 Then
        errMsg = "symbol " & CStr(sym) & " has no code in this block's table"
        Exit Function
    End If
    WBits m_lCode(sym), m_lLen(sym)
    WLit = True
End Function

Private Function WMatch(ByVal ln As Long, ByVal dv As Long, ByRef errMsg As String) As Boolean
    Dim lc As Long
    Dim dc As Long
    Dim ex As Long
    If ln < 3 Or ln > 258 Then
        errMsg = "match length out of range: " & CStr(ln)
        Exit Function
    End If
    lc = LenCode(ln)
    If m_lLen(257 + lc) = 0 Then
        errMsg = "length code " & CStr(257 + lc) & " has no code in this block"
        Exit Function
    End If
    WBits m_lCode(257 + lc), m_lLen(257 + lc)
    ex = m_lext(lc)
    If ex > 0 Then WBits ln - m_lbase(lc), ex
    dc = DistCode(dv)
    If m_dLen(dc) = 0 Then
        errMsg = "distance code " & CStr(dc) & " has no code in this block"
        Exit Function
    End If
    WBits m_dCode(dc), m_dLen(dc)
    ex = m_dext(dc)
    If ex > 0 Then WBits dv - m_dbase(dc), ex
    WMatch = True
End Function

Private Function LenCode(ByVal ln As Long) As Long
    Dim i As Long
    For i = 28 To 0 Step -1
        If ln >= m_lbase(i) Then
            LenCode = i
            Exit Function
        End If
    Next i
End Function

Private Function DistCode(ByVal dv As Long) As Long
    Dim i As Long
    For i = 29 To 0 Step -1
        If dv >= m_dbase(i) Then
            DistCode = i
            Exit Function
        End If
    Next i
End Function

Private Sub CopyBits(ByRef src() As Byte, ByVal atBit As Long, ByVal n As Long)
    Dim i As Long
    Dim take As Long
    Dim byteAt As Long
    Dim bitAt As Long
    Dim v As Long
    Do While i < n
        byteAt = (atBit + i) \ 8
        bitAt = (atBit + i) And 7
        take = 8 - bitAt
        If take > n - i Then take = n - i
        v = (CLng(src(byteAt)) \ m_pow2(bitAt)) And (m_pow2(take) - 1)
        WBits v, take
        i = i + take
    Loop
End Sub

'==============================================================================
' the patch
'==============================================================================
Public Function Rdv3ZipPatchByte(ByVal srcPath As String, ByVal dstPath As String, _
                                 ByVal entryName As String, ByVal pattern As String, _
                                 ByVal newByte As Byte, ByVal estOut As Long, _
                                 ByVal margin As Long, ByRef info As String, _
                                 ByRef errMsg As String) As Boolean
    Dim dataAt As Long
    Dim compLen As Long
    Dim uncompLen As Long
    Dim comp() As Byte
    Dim f As Integer
    Dim matFrom As Long
    Dim t As Double
    Dim walkMs As Double
    Dim ok As Boolean
    Dim refound As Long
    Dim readMs As Double

    errMsg = ""
    info = ""
    m_info = ""
    m_wrInfo = ""
    Rdv3ZipInit
    If Not Rdv3ZipEntryInfo(srcPath, entryName, dataAt, compLen, uncompLen, errMsg) Then Exit Function

    On Error GoTo Fail
    t = Rdv3Ticks()
    f = FreeFile
    Open srcPath For Binary Access Read As #f
    ReDim comp(0 To compLen - 1)
    Get #f, dataAt + 1, comp
    Close #f
    f = 0
    readMs = Rdv3MsSince(t)

    ' the row is looked for from a little before where it should be; the window
    ' is built the whole way regardless
    matFrom = estOut - margin
    If matFrom < 0 Then matFrom = 0
    t = Rdv3Ticks()
    ok = Rdv3ZipWalk(comp, compLen, pattern, matFrom, True, newByte)
    If ok And m_pTarget < 0 And matFrom > 0 Then
        ' the guess was wrong: pay for the honest full pass rather than report
        ' a byte that was never found
        refound = 1
        ok = Rdv3ZipWalk(comp, compLen, pattern, 0, True, newByte)
    End If
    walkMs = Rdv3MsSince(t)
    If Not ok Then
        errMsg = "walk failed: " & m_err
        Exit Function
    End If
    If m_pTarget < 0 Then
        errMsg = "target cell not found: " & pattern
        Exit Function
    End If
    If m_spliceStartBit < 0 Or m_spliceEndBit < 0 Then
        errMsg = "the affected span was not closed"
        Exit Function
    End If

    ' pad this side to a byte so the untouched rest can be copied as bytes. An
    ' empty stored block decodes to nothing, so the content is untouched. When
    ' the span ran to the end there is nothing to line up with, and the last
    ' partial byte is simply flushed.
    If Not m_spanToEnd Then
        If ((m_spliceStartBit + m_wBits) Mod 8&) <> 0 Then
            If Not WPadToByte(errMsg) Then Exit Function
        End If
    End If
    WFlushTail
    t = Rdv3Ticks()
    If Not SpliceAndWrite(srcPath, dstPath, entryName, dataAt, compLen, uncompLen, _
                          comp, errMsg) Then Exit Function

    info = "p=" & CStr(m_pTarget) & ";old=" & CStr(m_oldByte) & ";new=" & CStr(newByte) & _
        ";walked=" & CStr(m_out) & ";blocks=" & CStr(m_blocks) & _
        ";span_bits_old=" & CStr(m_spliceEndBit - m_spliceStartBit) & _
        ";span_bits_new=" & CStr(m_wBits) & ";align_blocks=" & CStr(m_alignBlocks) & _
        ";refound=" & CStr(refound) & _
        ";read_ms=" & Format$(readMs, "0.0") & _
        ";walk_ms=" & Format$(walkMs, "0.0") & _
        ";write_ms=" & Format$(Rdv3MsSince(t), "0.0") & _
        ";" & m_wrInfo
    m_info = info
    Rdv3ZipPatchByte = True
    Exit Function
Fail:
    errMsg = "patch error " & Err.Number & ": " & Err.Description & " at " & m_stage
    On Error Resume Next
    If f <> 0 Then Close #f
End Function

' an empty stored block: three header bits, padding to the byte, then LEN = 0
' and its complement. It produces no output and leaves us byte-aligned.
Private Function WPadToByte(ByRef errMsg As String) As Boolean
    Dim pad As Long
    WBits 0, 1                                   ' BFINAL = 0
    WBits 0, 2                                   ' BTYPE  = 00, stored
    pad = (8& - ((m_spliceStartBit + m_wBits) Mod 8&)) Mod 8&
    If pad > 0 Then WBits 0, pad
    WBits 0, 16                                  ' LEN  = 0
    WBits &HFFFF&, 16                            ' NLEN = ~0
    WPadToByte = True
End Function

Private Sub WFlushTail()
    If m_wCnt > 0 Then
        If m_wAt > UBound(m_wBuf) Then ReDim Preserve m_wBuf(0 To m_wAt + 16)
        m_wBuf(m_wAt) = m_wHold And &HFF&
        m_wAt = m_wAt + 1
        m_wHold = 0
        m_wCnt = 0
    End If
End Sub

' Writes the package out with the target entry's stream spliced and everything
' else copied byte for byte -- headers, extra fields, comments, order, times,
' flags and compression methods included. Entries are copied as contiguous
' regions, so anything between them survives as well.
Private Function SpliceAndWrite(ByVal srcPath As String, ByVal dstPath As String, _
                                ByVal entryName As String, ByVal dataAt As Long, _
                                ByVal compLen As Long, ByVal uncompLen As Long, _
                                ByRef comp() As Byte, ByRef errMsg As String) As Boolean
    Dim fb() As Byte
    Dim f As Integer
    Dim n As Long
    Dim eocd As Long
    Dim cnt As Long
    Dim cdOff As Long
    Dim i As Long, k As Long, q As Long
    Dim cRec() As Long
    Dim cLen() As Long
    Dim lOff() As Long
    Dim nameLen As Long
    Dim nmB() As Byte
    Dim target As Long
    Dim newCompLen As Long
    Dim oldCrc As Long
    Dim newCrc As Long
    Dim seg() As Byte
    Dim at As Long
    Dim tailStart As Long
    Dim delta As Long
    Dim fs As Integer
    Dim preBytes As Long
    Dim preBits As Long
    Dim tailByte As Long
    Dim midBuf() As Byte
    Dim midLen As Long
    Dim headLen As Long

    Dim t0 As Double
    Dim tRead As Double
    Dim tSplice As Double
    Dim tAssemble As Double

    errMsg = ""
    m_stage = "splice"
    On Error GoTo Fail
    ' only the directory is needed in memory here; the bulk moves with Get/Put
    t0 = Rdv3Ticks()
    f = FreeFile
    Open srcPath For Binary Access Read As #f
    n = LOF(f)
    ReDim fb(0 To n - 1)
    Get #f, 1, fb
    Close #f
    f = 0
    tRead = Rdv3MsSince(t0)

    eocd = -1
    For i = n - 22 To 0 Step -1
        If fb(i) = &H50 And fb(i + 1) = &H4B And fb(i + 2) = &H5 And fb(i + 3) = &H6 Then
            eocd = i
            Exit For
        End If
    Next i
    If eocd < 0 Then
        errMsg = "no end-of-central-directory"
        Exit Function
    End If
    cnt = CLng(fb(eocd + 10)) + CLng(fb(eocd + 11)) * 256&
    cdOff = U32(fb, eocd + 16)

    ReDim cRec(0 To cnt - 1)
    ReDim cLen(0 To cnt - 1)
    ReDim lOff(0 To cnt - 1)
    target = -1
    q = cdOff
    For i = 0 To cnt - 1
        cRec(i) = q
        nameLen = CLng(fb(q + 28)) + CLng(fb(q + 29)) * 256&
        cLen(i) = 46 + nameLen + (CLng(fb(q + 30)) + CLng(fb(q + 31)) * 256&) + _
                  (CLng(fb(q + 32)) + CLng(fb(q + 33)) * 256&)
        lOff(i) = U32(fb, q + 42)
        ReDim nmB(0 To nameLen - 1)
        For k = 0 To nameLen - 1
            nmB(k) = fb(q + 46 + k)
        Next k
        If StrConv(nmB, vbUnicode) = entryName Then
            target = i
            oldCrc = U32(fb, q + 16)
        End If
        q = q + cLen(i)
    Next i
    If target < 0 Then
        errMsg = "entry vanished between passes: " & entryName
        Exit Function
    End If

    ' The new stream is: the original bits up to the span, the re-emitted span
    ' (which ends on a byte), and the original bytes from the byte-aligned
    ' boundary the span stopped at. Only the middle is built here; the two big
    ' pieces are moved with Get/Put, which copies natively.
    t0 = Rdv3Ticks()
    preBytes = m_spliceStartBit \ 8
    preBits = m_spliceStartBit And 7
    tailByte = m_spliceEndBit \ 8
    If (m_spliceEndBit And 7) <> 0 And Not m_spanToEnd Then
        errMsg = "the span did not stop on a byte boundary"
        Exit Function
    End If
    If m_spanToEnd Then tailByte = compLen
    If Not BuildMiddle(comp, preBytes, preBits, midBuf, midLen, errMsg) Then Exit Function
    newCompLen = preBytes + midLen + (compLen - tailByte)
    tSplice = Rdv3MsSince(t0)
    ' CRC-32 covers the UNCOMPRESSED bytes, and exactly one of them moved
    newCrc = oldCrc Xor Rdv3ZipCrcDelta(CLng(m_oldByte) Xor CLng(m_newByte), _
                                        uncompLen - m_pTarget - 1)
    t0 = Rdv3Ticks()

    delta = newCompLen - compLen
    tAssemble = Rdv3MsSince(t0)

    t0 = Rdv3Ticks()
    If Len(Dir$(dstPath)) > 0 Then Kill dstPath
    fs = FreeFile
    Open srcPath For Binary Access Read As #fs
    f = FreeFile
    Open dstPath For Binary Access Write As #f

    ' A: everything up to and including the untouched head of the stream
    headLen = dataAt + preBytes
    ReDim seg(0 To headLen - 1)
    Get #fs, 1, seg
    PutU32 seg, lOff(target) + 14, newCrc
    PutU32 seg, lOff(target) + 18, newCompLen
    Put #f, 1, seg

    ' B: the re-emitted span
    Put #f, headLen + 1, midBuf

    ' C: the rest of the file, with the directory offsets moved
    tailStart = dataAt + tailByte
    ReDim seg(0 To n - tailStart - 1)
    Get #fs, tailStart + 1, seg
    For i = 0 To cnt - 1
        If lOff(i) > lOff(target) Then
            PutU32 seg, cRec(i) - tailStart + 42, lOff(i) + delta
        End If
        If i = target Then
            PutU32 seg, cRec(i) - tailStart + 16, newCrc
            PutU32 seg, cRec(i) - tailStart + 20, newCompLen
        End If
    Next i
    PutU32 seg, eocd - tailStart + 16, cdOff + delta
    Put #f, headLen + midLen + 1, seg
    Close #f
    Close #fs
    f = 0
    fs = 0
    at = n + delta
    m_wrInfo = "wr_read=" & Format$(tRead, "0.0") & ";wr_splice=" & Format$(tSplice, "0.0") & _
        ";wr_assemble=" & Format$(tAssemble, "0.0") & ";wr_put=" & Format$(Rdv3MsSince(t0), "0.0") & _
        ";bytes=" & CStr(at)
    SpliceAndWrite = True
    Exit Function
Fail:
    errMsg = "splice/write error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If f <> 0 Then Close #f
End Function
' The only piece built by hand: the bits before the span that share a byte with
' it, the re-emitted span itself, and nothing else. It ends on a byte, because
' the span was chosen so that it does.
Private Function BuildMiddle(ByRef comp() As Byte, ByVal preBytes As Long, _
                             ByVal preBits As Long, ByRef midBuf() As Byte, _
                             ByRef midLen As Long, ByRef errMsg As String) As Boolean
    Dim total As Long
    Dim hold As Long
    Dim cnt As Long
    Dim at As Long
    Dim i As Long
    Dim nFull As Long
    Dim nRem As Long

    errMsg = ""
    total = preBits + m_wBits
    If (total Mod 8&) <> 0 And Not m_spanToEnd Then
        errMsg = "the re-emitted span does not end on a byte"
        Exit Function
    End If
    midLen = total \ 8
    If (total Mod 8&) <> 0 Then midLen = midLen + 1
    ReDim midBuf(0 To midLen - 1)
    If preBits > 0 Then
        hold = CLng(comp(preBytes)) And (m_pow2(preBits) - 1)
        cnt = preBits
    End If
    nFull = m_wBits \ 8
    nRem = m_wBits And 7
    For i = 0 To nFull - 1
        hold = hold Or (CLng(m_wBuf(i)) * m_pow2(cnt))
        midBuf(at) = hold And &HFF&
        at = at + 1
        hold = hold \ 256&
    Next i
    If nRem > 0 Then
        hold = hold Or ((CLng(m_wBuf(nFull)) And (m_pow2(nRem) - 1)) * m_pow2(cnt))
        cnt = cnt + nRem
        Do While cnt >= 8
            midBuf(at) = hold And &HFF&
            at = at + 1
            hold = hold \ 256&
            cnt = cnt - 8
        Loop
    End If
    If cnt > 0 Then
        midBuf(at) = hold And &HFF&
        at = at + 1
    End If
    If at <> midLen Then
        errMsg = "middle length mismatch: " & CStr(at) & " vs " & CStr(midLen)
        Exit Function
    End If
    BuildMiddle = True
End Function

'==============================================================================
' what the ledger asks for: set one processed flag, keep everything else
'==============================================================================
Public Function Rdv3ZipSetProcessed(ByVal path As String, ByVal tmpPath As String, _
                                    ByVal sheetRow As Long, ByVal rowsTotal As Long, _
                                    ByVal newVal As Boolean, ByRef info As String, _
                                    ByRef errMsg As String) As Boolean
    Dim entryName As String
    Dim pat As String
    Dim est As Long
    Dim uncompLen As Long
    Dim dataAt As Long
    Dim compLen As Long
    Dim nb As Byte

    errMsg = ""
    entryName = Rdv3ZipLedgerEntry(path)
    If Left$(entryName, 3) = "ERR" Then
        errMsg = entryName
        Exit Function
    End If
    If Not Rdv3ZipEntryInfo(path, entryName, dataAt, compLen, uncompLen, errMsg) Then Exit Function

    pat = "<c r=""A" & CStr(sheetRow) & """ s=""1"" t=""b""><v>"
    If rowsTotal > 0 Then
        est = CLng((CDbl(sheetRow - 1) / CDbl(rowsTotal + 1)) * CDbl(uncompLen))
    End If
    If newVal Then
        nb = 49
    Else
        nb = 48
    End If
    Rdv3ZipSetProcessed = Rdv3ZipPatchByte(path, tmpPath, entryName, pat, nb, est, 2097152, info, errMsg)
End Function

'==============================================================================
' ZIP directory helpers
'==============================================================================
Public Function Rdv3ZipEntryInfo(ByVal path As String, ByVal entryName As String, _
                                 ByRef dataAt As Long, ByRef compLen As Long, _
                                 ByRef uncompLen As Long, ByRef errMsg As String) As Boolean
    Dim f As Integer
    Dim n As Long, tailLen As Long, i As Long, k As Long, p As Long
    Dim tail() As Byte
    Dim eocd As Long, cnt As Long, cdOff As Long, cdSize As Long
    Dim cd() As Byte
    Dim nameLen As Long
    Dim nmB() As Byte
    Dim lh(0 To 29) As Byte

    errMsg = ""
    On Error GoTo Fail
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    tailLen = 66000
    If tailLen > n Then tailLen = n
    ReDim tail(0 To tailLen - 1)
    Get #f, n - tailLen + 1, tail
    eocd = -1
    For i = tailLen - 22 To 0 Step -1
        If tail(i) = &H50 And tail(i + 1) = &H4B And tail(i + 2) = &H5 And tail(i + 3) = &H6 Then
            eocd = i
            Exit For
        End If
    Next i
    If eocd < 0 Then
        errMsg = "no end-of-central-directory"
        GoTo CloseFail
    End If
    cnt = CLng(tail(eocd + 10)) + CLng(tail(eocd + 11)) * 256&
    cdSize = U32(tail, eocd + 12)
    cdOff = U32(tail, eocd + 16)
    ReDim cd(0 To cdSize - 1)
    Get #f, cdOff + 1, cd
    p = 0
    For i = 0 To cnt - 1
        nameLen = CLng(cd(p + 28)) + CLng(cd(p + 29)) * 256&
        ReDim nmB(0 To nameLen - 1)
        For k = 0 To nameLen - 1
            nmB(k) = cd(p + 46 + k)
        Next k
        If StrConv(nmB, vbUnicode) = entryName Then
            Get #f, U32(cd, p + 42) + 1, lh
            dataAt = U32(cd, p + 42) + 30 + _
                     (CLng(lh(26)) + CLng(lh(27)) * 256&) + (CLng(lh(28)) + CLng(lh(29)) * 256&)
            compLen = U32(cd, p + 20)
            uncompLen = U32(cd, p + 24)
            Close #f
            Rdv3ZipEntryInfo = True
            Exit Function
        End If
        p = p + 46 + nameLen + (CLng(cd(p + 30)) + CLng(cd(p + 31)) * 256&) + _
            (CLng(cd(p + 32)) + CLng(cd(p + 33)) * 256&)
    Next i
    errMsg = "entry not found: " & entryName
CloseFail:
    On Error Resume Next
    Close #f
    Exit Function
Fail:
    errMsg = "zip read error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #f
End Function

' the largest worksheet part, which is the ledger sheet
Public Function Rdv3ZipLedgerEntry(ByVal path As String) As String
    Dim f As Integer
    Dim n As Long
    Dim tail() As Byte
    Dim tailLen As Long
    Dim i As Long, k As Long, p As Long
    Dim eocd As Long, cnt As Long, cdOff As Long, cdSize As Long
    Dim cd() As Byte
    Dim nameLen As Long
    Dim nmB() As Byte
    Dim nm As String
    Dim best As String
    Dim bestLen As Long

    On Error GoTo Fail
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    tailLen = 66000
    If tailLen > n Then tailLen = n
    ReDim tail(0 To tailLen - 1)
    Get #f, n - tailLen + 1, tail
    eocd = -1
    For i = tailLen - 22 To 0 Step -1
        If tail(i) = &H50 And tail(i + 1) = &H4B And tail(i + 2) = &H5 And tail(i + 3) = &H6 Then
            eocd = i
            Exit For
        End If
    Next i
    If eocd < 0 Then
        Rdv3ZipLedgerEntry = "ERR no eocd"
        Close #f
        Exit Function
    End If
    cnt = CLng(tail(eocd + 10)) + CLng(tail(eocd + 11)) * 256&
    cdSize = U32(tail, eocd + 12)
    cdOff = U32(tail, eocd + 16)
    ReDim cd(0 To cdSize - 1)
    Get #f, cdOff + 1, cd
    Close #f
    bestLen = -1
    p = 0
    For i = 0 To cnt - 1
        nameLen = CLng(cd(p + 28)) + CLng(cd(p + 29)) * 256&
        ReDim nmB(0 To nameLen - 1)
        For k = 0 To nameLen - 1
            nmB(k) = cd(p + 46 + k)
        Next k
        nm = StrConv(nmB, vbUnicode)
        If LCase$(Left$(nm, 15)) = "xl/worksheets/s" Then
            If U32(cd, p + 24) > bestLen Then
                bestLen = U32(cd, p + 24)
                best = nm
            End If
        End If
        p = p + 46 + nameLen + (CLng(cd(p + 30)) + CLng(cd(p + 31)) * 256&) + _
            (CLng(cd(p + 32)) + CLng(cd(p + 33)) * 256&)
    Next i
    Rdv3ZipLedgerEntry = best
    Exit Function
Fail:
    Rdv3ZipLedgerEntry = "ERR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #f
End Function

Private Function U32(ByRef b() As Byte, ByVal at As Long) As Long
    Dim hi As Long
    Dim v As Double
    hi = CLng(b(at + 3))
    v = CDbl(b(at)) + CDbl(b(at + 1)) * 256# + CDbl(b(at + 2)) * 65536#
    If hi > 127 Then
        U32 = CLng(v + CDbl(hi - 256) * 16777216#)
    Else
        U32 = CLng(v + CDbl(hi) * 16777216#)
    End If
End Function

Private Sub PutU32(ByRef b() As Byte, ByVal at As Long, ByVal v As Long)
    b(at) = v And &HFF&
    b(at + 1) = Shr(v, 8) And &HFF&
    b(at + 2) = Shr(v, 16) And &HFF&
    b(at + 3) = Shr(v, 24) And &HFF&
End Sub

Private Function Shr(ByVal v As Long, ByVal n As Long) As Long
    Dim r As Long
    If n <= 0 Then
        Shr = v
        Exit Function
    End If
    If n >= 32 Then Exit Function
    If v >= 0 Then
        Shr = v \ m_pow2(n)
    Else
        r = (v And &H7FFFFFFF) \ m_pow2(n)
        Shr = r Or m_pow2(31 - n)
    End If
End Function

'==============================================================================
' CRC-32 and the one-byte-difference update. The checksum is linear over GF(2),
' so changing one byte of a 128 MB part costs a 32x32 matrix power, not a
' rescan.
'==============================================================================
Public Sub Rdv3ZipCrcInit()
    Dim nn As Long, k As Long, c As Long
    If m_crcReady Then Exit Sub
    Rdv3ZipInit
    For nn = 0 To 255
        c = nn
        For k = 1 To 8
            If (c And 1&) <> 0 Then
                c = &HEDB88320 Xor Shr(c, 1)
            Else
                c = Shr(c, 1)
            End If
        Next k
        m_crcTbl(nn) = c
    Next nn
    m_crcReady = True
End Sub

Public Function Rdv3ZipCrcBytes(ByVal c As Long, ByRef b() As Byte, ByVal n As Long) As Long
    Dim i As Long
    Rdv3ZipCrcInit
    For i = 0 To n - 1
        If c < 0 Then
            c = m_crcTbl((c Xor CLng(b(i))) And &HFF&) Xor (((c And &H7FFFFFFF) \ &H100&) Or &H800000)
        Else
            c = m_crcTbl((c Xor CLng(b(i))) And &HFF&) Xor (c \ &H100&)
        End If
    Next i
    Rdv3ZipCrcBytes = c
End Function

Private Sub CrcMatrixInit()
    Dim j As Long, c As Long
    If m_crcMReady Then Exit Sub
    Rdv3ZipCrcInit
    For j = 0 To 31
        c = m_pow2(j)
        m_crcM(j) = m_crcTbl(c And &HFF&) Xor Shr(c, 8)
    Next j
    m_crcMReady = True
End Sub

Private Function GfApply(ByRef mat() As Long, ByVal v As Long) As Long
    Dim j As Long, r As Long
    For j = 0 To 31
        If (v And m_pow2(j)) <> 0 Then r = r Xor mat(j)
    Next j
    GfApply = r
End Function

Private Sub GfMul(ByRef dst() As Long, ByRef a() As Long, ByRef b() As Long)
    Dim j As Long
    Dim t(0 To 31) As Long
    For j = 0 To 31
        t(j) = GfApply(a, b(j))
    Next j
    For j = 0 To 31
        dst(j) = t(j)
    Next j
End Sub

Public Function Rdv3ZipCrcDelta(ByVal delta As Long, ByVal zerosAfter As Long) As Long
    Dim op(0 To 31) As Long
    Dim acc(0 To 31) As Long
    Dim j As Long, nn As Long, v As Long
    CrcMatrixInit
    v = m_crcTbl(delta And &HFF&)
    For j = 0 To 31
        op(j) = m_crcM(j)
        acc(j) = m_pow2(j)
    Next j
    nn = zerosAfter
    Do While nn > 0
        If (nn And 1&) <> 0 Then GfMul acc, op, acc
        nn = nn \ 2
        If nn > 0 Then GfMul op, op, op
    Loop
    Rdv3ZipCrcDelta = GfApply(acc, v)
End Function

' measurement entry: walk only, no patch, so the cost of finding a row and the
' effect of where materialising starts can be read directly
Public Function Rdv3ZipProbeWalk(ByVal path As String, ByVal pattern As String, _
                                 ByVal matFrom As Long) As String
    Dim entryName As String
    Dim dataAt As Long, compLen As Long, uncompLen As Long
    Dim e As String
    Dim comp() As Byte
    Dim f As Integer
    Dim t As Double
    Dim ok As Boolean

    entryName = Rdv3ZipLedgerEntry(path)
    If Not Rdv3ZipEntryInfo(path, entryName, dataAt, compLen, uncompLen, e) Then
        Rdv3ZipProbeWalk = "ERR " & e
        Exit Function
    End If
    On Error GoTo Fail
    f = FreeFile
    Open path For Binary Access Read As #f
    ReDim comp(0 To compLen - 1)
    Get #f, dataAt + 1, comp
    Close #f
    t = Rdv3Ticks()
    ok = Rdv3ZipWalk(comp, compLen, pattern, matFrom, False, 49)
    Rdv3ZipProbeWalk = IIf(ok, "ok", "ERR " & m_err) & " matFrom=" & CStr(matFrom) & _
        " hit=" & CStr(m_patHit) & " p=" & CStr(m_pTarget) & " out=" & CStr(m_out) & _
        " blocks=" & CStr(m_blocks) & " ms=" & Format$(Rdv3MsSince(t), "0.0")
    Exit Function
Fail:
    Rdv3ZipProbeWalk = "ERR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #f
End Function

'==============================================================================
' build-time self test: the arithmetic this module rests on, checked against
' published values
'==============================================================================
Public Function Rdv3ZipSelfTest() As String
    Dim b(0 To 8) As Byte
    Dim i As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim d As Long
    Rdv3ZipInit
    Rdv3ZipCrcInit
    For i = 0 To 8
        b(i) = Asc(Mid$("123456789", i + 1, 1))
    Next i
    c1 = Rdv3ZipCrcBytes(-1, b, 9) Xor -1
    d = CLng(Asc("5")) Xor CLng(Asc("X"))
    b(4) = Asc("X")
    c2 = Rdv3ZipCrcBytes(-1, b, 9) Xor -1
    Rdv3ZipSelfTest = "crc(123456789)=" & Hex$(c1) & " expect=CBF43926" & _
        " delta_ok=" & CStr((c1 Xor Rdv3ZipCrcDelta(d, 9 - 4 - 1)) = c2) & _
        " rev(1,3)=" & CStr(RevBits(1, 3)) & " expect=4" & _
        " lencode(258)=" & CStr(LenCode(258)) & " expect=28" & _
        " distcode(32768)=" & CStr(DistCode(32768)) & " expect=29"
End Function
