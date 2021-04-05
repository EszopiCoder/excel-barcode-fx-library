Attribute VB_Name = "modAztec"
Option Explicit
Dim typ As Long, enc(1665) As Integer
Dim md As Long, eb As Long, el As Long, b As Long

' Aztec bar code symbol creation according ISO/IEC 24778:2008
'  param text to encode
'  param security optional: percentage of checkwords used for security 2%-90% (23%)
'  param layers optional: number of layers (size), default autodetect, 0 - Aztec rune
'   creates Actec and compact Aztec bar code symbol as shape in Excel cell.
Public Function Aztec(text As String, Optional security As Integer, Optional layers As Integer = 1) As String
Attribute Aztec.VB_Description = "Draw Aztec barcode"
Attribute Aztec.VB_ProcData.VB_Invoke_Func = " \n18"
Dim fColor As Long, bColor As Long, line As Long, shp As Shape, txt As String
Dim x As Long, y As Long, dx As Long, dy As Long, ctr As Long, ec As Long
Dim c As Long, i As Long, j As Long, k As Long, l As Long, m As Long

On Error GoTo failed
If Not TypeOf Application.Caller Is Range Then Err.Raise 513, "Aztec code", "Call only from sheet"
fColor = vbBlack: bColor = vbBlack: line = xlHairline ' redraw graphic ?
For Each shp In Application.Caller.Parent.Shapes
    If shp.Name = Application.Caller.Address Then
        If shp.Title = text Then Exit Function ' same as prev ?
        fColor = shp.Fill.ForeColor.RGB  ' remember format
        bColor = shp.line.ForeColor.RGB
        line = shp.line.Weight
        shp.Delete
    End If
Next shp
If security < 1 Then security = 23 Else If security > 90 Then security = 90
txt = IIf(text = "", " ", utf16to8(text)) ' at least 1 char
el = Len(txt): x = 4: typ = 0
Do ' compute word size b: 6/8/10/12 bits
    i = Int(el * 100 / (100 - security) + 3) * x ' needed bits, at least 3 checkwords
    If i > l Then l = i
    b = IIf(l <= 240, 6, IIf(l <= 1920, 8, IIf(l <= 10208, 10, 12))) ' bit capacity -> word size
    i = IIf(layers < 3, 6, IIf(layers < 9, 8, IIf(layers < 23, 10, 12))) ' layer paramerter
    If i > b Then b = i
    If x >= b Then Exit Do
    el = 0: md = 0: eb = 0: enc(0) = 0 ' clr bit stream
    For i = 1 To Len(txt) ' scan text
        c = Asc(Mid(txt, i, 1)): k = 0
        If i < Len(txt) Then k = Asc(Mid(txt, i + 1, 1))
        If c = 32 Then ' space
            If md = 3 Then push 31: md = 0 ' punct: latch to upper
            c = 1 ' space in all other modes
        ElseIf md = 4 And c = 44 Then
            c = 12 ' , in digit mode
        ElseIf md = 4 And c = 46 Then
            c = 13 ' . in digit mode
        ElseIf ((c = 44 Or c = 46 Or c = 58) And k = 32) Or (c = 13 And k = 10) Then
            If md <> 3 Then push (0) ' shift to punct
            push IIf(c = 46, 3, IIf(c = 44, 4, IIf(c = 58, 5, 2))) ' two char encoding
            i = i + 1: GoTo continue
        Else
            c = IIf(c = 13 And modeOf(k) \ 32 = md, 97, modeOf(c))
            If c < 0 Then ' binary
                If md > 2 Then push IIf(md = 3, 31, 14): md = 0 ' latch to upper
                j = 0: push 31 ' shift to binary
                For k = 0 To Len(txt) - i - 1 ' calc binary length
                    If modeOf(Asc(Mid(txt, k + i, 1))) < 0 Then
                        j = 0
                    Else
                        j = j + 1
                        If j > 5 Then Exit For ' look for at least 5 consecutive non binary chars
                    End If
                Next k
                k = k - j
                If k > 30 Then
                    push 0: push k - 30, 11
                Else
                    push k + 1
                End If
                For j = 0 To k  ' encode binary data
                    push Asc(Mid(txt, i + j, 1)), 8
                    If el > 1660 Then Exit For
                Next j
                i = i + k: GoTo continue
            End If
            m = c \ 32 ' needed mode
            If m = 4 And md = 2 Then push 29: md = 0 ' mixed to upper (to digit)
            If m <> 3 And md = 3 Then push 31: md = 0 ' exit punct: to upper
            If m <> 4 And md = 4 Then ' exit digit
                If (m = 3 Or m = 0) And modeOf(k) > 129 Then
                    push (3 - m) * 5: push c And 31, 5 ' shift to punct/upper
                    GoTo continue
                End If
                push 14: md = 0 ' latch to upper
            End If
            If md <> m Then ' mode change needed
                If m = 3 Then ' to punct
                    If md <> 4 And modeOf(k) \ 32 = 3 Then ' 2x punct, latch to punkt
                        If md <> 2 Then push 29 ' latch to mixed
                        push 30 ' latch to punct
                        md = 3 ' mode punct
                    Else
                        push 0 ' shift to punct
                    End If
                ElseIf md = 1 And m = 0 Then ' lower to upper
                    If modeOf(k) \ 32 = 1 Then
                        push 28 ' shift
                    Else
                        push 30: push 14, 4 ' latch
                        md = 0
                    End If
                Else ' latch to ..
                    push Array(29, 28, 29, 30, 30)(m)
                    md = m
                End If
            End If
        End If
        push c And 31 ' add char
        If el > 1660 Then Exit For
continue:
    Next i
    push 2 ^ (b - eb) - 1, b - eb ' add padding bits
    x = b
Loop
If el > 1660 Then Err.Raise 514, "Aztec code", "Message too long."
typ = IIf(l > 608 Or el > 64, 14, 11) ' full or compact Aztec
md = val(Left(txt, 3)) ' Aztec rune possible ?
If md < 0 Or md > 255 Or md & "" <> txt Or layers > 0 Then
    i = -Int((typ - Sqr(l + typ * typ)) / 4) ' needed layers
    If i > layers Then layers = i
    If layers > 32 Then layers = 32
End If
ec = (8 * layers * (typ + 2 * layers)) \ b - el ' # of checkwords
typ = typ \ 2: ctr = typ + 2 * layers: ctr = ctr + (ctr - 1) \ 15 ' center position
security = 100 * ec / (el + ec)

With Application.Caller.Parent.Shapes ' layout Aztec barcode
    m = .Count + 1
    For y = 1 - typ To typ - 1 ' layout central finder
        For x = 1 - typ To typ - 1
            If (IIf(Abs(x) > Abs(y), x, y) And 1) = 0 Then
                .AddShape(msoShapeRectangle, ctr + x, ctr + y, 1, 1).Name = Application.Caller.Address
            End If
        Next x
    Next y
    For i = 0 To 5 ' orientation marks
        x = Array(-typ, -typ, 1 - typ, typ, typ, typ)(i)
        y = Array(1 - typ, -typ, -typ, typ - 1, 1 - typ, -typ)(i)
        .AddShape(msoShapeRectangle, ctr + x, ctr + y, 1, 1).Name = Application.Caller.Address
    Next i
    If layers > 0 Then ' layout data
        addCheck ec, 2 ^ b - 1, Array(67, 301, 1033, 4201)(b / 2 - 3) ' error correction, generator polynomial
        x = -typ: y = x - 1 ' start of layer 1 at top left
        j = (3 * typ + 11) / 2: l = j ' length of inner side
        dx = 1: dy = 0 ' direction right
        For ec = ec + el - 1 To 0 Step -1 ' layout codeword
            c = enc(ec) ' data in reversed order inside to outside
            For i = 1 To b / 2
                If c And 1 Then ' odd bit
                    .AddShape(msoShapeRectangle, ctr + x, ctr + y, 1, 1).Name = Application.Caller.Address
                End If
                move x, y, dy, -dx ' move across
                If c And 2 Then ' even bit
                    .AddShape(msoShapeRectangle, ctr + x, ctr + y, 1, 1).Name = Application.Caller.Address
                End If
                move x, y, dx - dy, dx + dy ' move ahead
                j = j - 1
                If j = 0 Then ' spiral turn
                    move x, y, dy, -dx ' move across
                    j = dx: dx = -dy: dy = j ' rotate clockwise
                    If dx < 1 Then
                        move x, y, dx - dy, dx + dy ' move ahead
                        move x, y, dx - dy, dx + dy ' move ahead
                    Else
                        l = l + 4 ' full turn -> next layer
                    End If
                    j = l ' start new side
                End If
                c = c \ 4
            Next i
        Next ec
        If typ = 7 Then ' layout reference grid
            For x = (15 - ctr) And -16 To ctr Step 16
                For y = (1 - ctr) And -2 To ctr Step 2
                    If Abs(x) > typ Or Abs(y) > typ Then
                        .AddShape(msoShapeRectangle, ctr + x, ctr + y, 1, 1).Name = Application.Caller.Address ' down
                        If y And 15 Then
                            .AddShape(msoShapeRectangle, ctr + y, ctr + x, 1, 1).Name = Application.Caller.Address ' across
                        End If
                    End If
                Next y
            Next x
        End If
        md = (layers - 1) * (typ * 992 - 4896&) + el - 1 ' 2/5 + 6/11 mode bits
    End If
    el = typ - 3 ' process modes message compact/full
    For i = el - 1 To 0 Step -1
        enc(i) = md And 15 ' mode to 4 bit words
        md = md \ 16
    Next i
    addCheck typ \ 2 + 3, 15, 19 ' add 5/6 words error correction
    el = el + typ \ 2 + 3 ' init bit stream
    b = (typ * 3) \ 2  ' 7/10 bits per side
    eb = 0: j = IIf(layers, 0, 10) 'XOR Aztec rune data
    For i = 0 To b - 1
        push j Xor enc(i), 4 ' 8/16 words to 4 chunks
    Next i
    j = 1 ' layout mode data
    For i = 2 - typ To typ - 2
        If typ = 7 And i = 0 Then i = i + 1 ' skip reference grid
        If enc(b) And j Then .AddShape(msoShapeRectangle, ctr - i, ctr - typ, 1, 1).Name = Application.Caller.Address ' top
        If enc(b + 1) And j Then .AddShape(msoShapeRectangle, ctr + typ, ctr - i, 1, 1).Name = Application.Caller.Address ' right
        If enc(b + 2) And j Then .AddShape(msoShapeRectangle, ctr + i, ctr + typ, 1, 1).Name = Application.Caller.Address ' bottom
        If enc(b + 3) And j Then .AddShape(msoShapeRectangle, ctr - typ, ctr + i, 1, 1).Name = Application.Caller.Address     ' left
        j = j + j
    Next i
    m = .Count - m
    ReDim shps(m) As Integer ' group all shapes
    For i = .Count To 1 Step -1
        If .Range(i).Name = Application.Caller.Address Then
            shps(m) = i: m = m - 1
            If m < 0 Then Exit For
        End If
    Next i
    With .Range(shps).Group
        .Fill.ForeColor.RGB = fColor ' format barcode shape
        .line.ForeColor.RGB = bColor
        .line.Weight = line
        x = Application.Caller.MergeArea.Width
        y = Application.Caller.MergeArea.Height
        If x > y Then x = y
        .Width = x * (2 * ctr + 1) / (2 * ctr + 3) ' fit symbol in excel cell
        .Height = .Width
        .Left = Application.Caller.Left + (Application.Caller.MergeArea.Width - .Width) / 2
        .Top = Application.Caller.Top + (Application.Caller.MergeArea.Height - .Height) / 2
        .Name = Application.Caller.Address ' link shape to data
        .Title = text
        .AlternativeText = "Aztec " & IIf(typ = 5, "compact", "full") & " barcode, security " & security & "%, layers " & layers & ", " & (2 * ctr + 1) & "x" & (2 * ctr + 1) & " cells"
        .LockAspectRatio = True
        .Placement = xlMove
    End With
End With
failed:
If Err.Number Then Aztec = "ERROR Aztec: " & Err.Description
End Function

' get character encoding mode of ch
Private Function modeOf(ByVal ch As Integer) As Integer
Dim i As Integer, k As Variant
If ch = 32 Then modeOf = md * 32: Exit Function ' space
k = Array(0, 14, 65, 26, 32, 52, 32, 48, 69, 47, 58, 82, 57, 64, 59, 64, 91, -63, 96, 123, -63)
For i = 0 To UBound(k) Step 3 ' check range
    If ch > k(i) And ch < k(i + 1) Then Exit For
Next i
If i <= UBound(k) Then modeOf = ch + k(i + 2): Exit Function ' ch in range
i = InStr("@\^_'|~¦[]{}", Chr(ch))
modeOf = IIf(i = 0, -1, IIf(i < 9, 20 + 64, 27 + 96 - 8) + i - 1) ' binary/mixed/punct
End Function

' add value to data stream
Private Sub push(ByVal val As Long, Optional ByVal bits As Integer = 0)
val = val * 2 ^ b
If bits = 0 Then bits = IIf(md = 4, 4, 5)
eb = eb + bits
enc(el) = enc(el) + val \ 2 ^ eb ' add data
Do While eb >= b ' word full ?
    If typ = 0 And (enc(el) < 2 Or enc(el) + 3 > 2 ^ b) Then ' bit stuffing
        enc(el) = enc(el) Xor ((enc(el) + 3) \ 2 And 1) ' add complementary bit
        eb = eb + 1
    End If
    eb = eb - b: el = el + 1
    enc(el) = (val \ 2 ^ eb) And ((2 ^ b) - 1)
Loop
End Sub

' compute Reed Solomon error detection and correction
Private Sub addCheck(ByVal ec As Integer, ByVal s As Integer, ByVal p As Integer)
Dim i As Integer, j As Integer, x As Integer
ReDim rc(ec + 2) As Integer, lg(s + 1) As Integer, ex(s) As Integer
j = 1
For i = 0 To s - 1 ' compute log/exp table of Galois field
    ex(i) = j: lg(j) = i
    j = j + j: If (j > s) Then j = j Xor p ' GF polynomial
Next i
rc(ec + 1) = 0
For i = 0 To ec ' compute RS generator polynomial
    rc(ec - i) = 1
    For j = ec - i + 1 To ec
        rc(j) = rc(j + 1) Xor ex((lg(rc(j)) + i) Mod s)
    Next j
    enc(el + i) = 0
Next i
For i = 0 To el - 1 ' compute RS checkwords
    x = enc(el) Xor enc(i)
    For j = 1 To ec
        enc(el + j - 1) = enc(el + j) Xor IIf(x, ex((lg(rc(j)) + lg(x)) Mod s), 0)
    Next j
Next i
End Sub

' move one cell
Private Sub move(x As Long, y As Long, ByVal dx As Integer, ByVal dy As Integer)
Do
    x = x + dx
Loop While typ = 7 And (x And 15) = 0 ' skip reference grid
Do
    y = y + dy
Loop While typ = 7 And (y And 15) = 0
End Sub
