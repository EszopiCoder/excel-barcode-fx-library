Attribute VB_Name = "modDataMatrix"
Option Explicit

'   creates Data Matrix barcode symbol as shape in Excel cell.
'   param text to encode
'   param rectangle boolean, default autodetect on cell dimension
' Data Matrix symbol creation according ISO/IEC 16022:2006
Public Function DataMatrix(text As String, Optional rectangle As Integer = -2) As String
Attribute DataMatrix.VB_Description = "Draw DataMatrix barcode"
Attribute DataMatrix.VB_ProcData.VB_Invoke_Func = " \n18"
On Error GoTo failed
If Not TypeOf Application.Caller Is Range Then Err.Raise 513, "DataMatrix", "Call only from sheet"
Dim enc As String, en As String, el As Long, k As Variant, l As Long
Dim h As Long, w As Long, nc As Byte, nr As Byte, shp As Shape
Dim fw As Integer, fh As Integer, i As Long, j As Long, b As Double
Dim c As Long, r As Double, s As Long, x As Long, y As Long, txt As String
Dim fColor As Long, bColor As Long, line As Long
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
txt = IIf(text = "", " ", utf16to8(text)): l = Len(txt)
For i = 1 To l ' ASCII mode encoding
    c = Asc(Mid(txt, i, 1)): r = 0
    If i < l Then r = Asc(Mid(txt, i + 1, 1))
    If c > 47 And c < 58 And r > 47 And r < 58 Then
        enc = enc + Chr((c - 48) * 10 + r - 48 + 130) ' 2 digits
        i = i + 1
    ElseIf (c > 127) Then ' extended char
        enc = enc + Chr(235) + Chr(c - 127)
    Else
        enc = enc + Chr(c + 1)
    End If
Next i
For x = 0 To 2 ' C40, TEXT and X12 modes encoding
    k = Array(Array(230, 31, 0, 0, 32, 9, 32 - 3, 47, 1, 33, 57, 9, 48 - 4, 64, 1, 58 - 15, 90, 9, 65 - 14, 95, 1, 91 - 22, 127, 2, 96, 255, 1, 0), _
              Array(239, 31, 0, 0, 32, 9, 32 - 3, 47, 1, 33, 57, 9, 48 - 4, 64, 1, 58 - 15, 90, 2, 64, 95, 1, 91 - 22, 122, 9, 97 - 14, 127, 2, 123 - 27, 255, 1, 0), _
              Array(238, 12, 8, 0, 13, 9, 13, 31, 8, 0, 32, 9, 32 - 3, 41, 8, 0, 42, 9, 42 - 1, 47, 8, 0, 57, 9, 48 - 4, 64, 8, 0, 90, 9, 65 - 14, 255, 8, 0))(x)
    b = 0: h = 0
    en = Chr(k(0)) ' start switch
    For i = 1 To l
        If h = 0 And i = l Then Exit For
        c = Asc(Mid(txt, i, 1))
        If c > 127 And k(0) <> 238 Then
            b = b * 40 + 1: b = b * 40 + 30
            h = h + 2: c = c - 128 ' hi bit in C40 & TEXT
        End If
        For j = 1 To 90 Step 3 ' select char set
            If c <= k(j) Then Exit For
        Next j
        s = k(j + 1) ' set
        If s = 8 Or (s = 9 And h = 0 And i = l) Then
            en = txt + txt
            Exit For ' char not in set, next mode
        End If
        If s < 5 And h = 2 And i = l Then  'Exit For ' last char in ASCII
            b = b * 40: h = 3 ' add padding
            i = i - 1
        Else
            If s < 5 Then b = b * 40 + s: h = h + 1 ' set
            b = b * 40 + c - k(j + 2): h = h + 1 ' char offset
            If h Mod 3 = 2 And k(0) <> 238 And i = l Then b = b * 40: h = h + 1 ' add padding
        End If
        Do While h > 2 ' pack 3 chars in 2 bytes
            h = h - 3: r = 40& ^ h
            c = Int(b / r) + 1
            en = en + Chr((c \ 256) And 255) + Chr(c And 255)
            b = b - c * r + r
        Loop
    Next i
    en = en + Chr(254) ' return to ASCII
    For i = i - h To l ' add last chars
        c = Asc(Mid(txt, i, 1))
        If (c > 127) Then en = en + Chr(235)
        en = en + Chr((c And 127) + 1)
    Next i
    If Len(en) < Len(enc) Then enc = en ' take shorter
Next x

j = (l + 1) And -4: b = 0: en = Chr(240) ' switch to Edifact
For i = 1 To j
    If i < j Then  ' encode char
        c = Asc(Mid(txt, i, 1))
        If c < 32 Or c > 94 Then Exit For ' not in set
    Else
        c = 31 ' return to ASCII
    End If
    b = b * 64 + (c And 63)
    If (i And 3) = 0 Then ' 4 data in 3 bytes
        en = en + Chr(b \ 65536) + Chr((b \ 256) And 255) + Chr(b And 255)
        b = 0
    End If
Next i
If j > 0 And i > j Then
    For i = i - 1 To l ' add last chars
        c = Asc(Mid(txt, i, 1))
        If (c > 127) Then en = en + Chr(235)
        en = en + Chr((c And 127) + 1)
    Next i
    If Len(en) < Len(enc) Then enc = en ' take shorter
End If

en = Chr(231) ' Base256 encoding
If l > 250 Then en = en + Chr((l \ 250 + 37) And 255) ' len high byte
en = en + Chr((l Mod 250 + (149 * (Len(en) + 1)) Mod 255 + 1) And 255) ' low
For i = 1 To l ' data in 255 state algo
    en = en + Chr((Asc(Mid(txt, i, 1)) + (149 * (Len(en) + 1)) Mod 255 + 1) And 255)
Next i
If Len(en) < Len(enc) Then enc = en ' take shorter

' compute symbol size
nc = 1: nr = 1: j = -1: b = 1: el = Len(enc) ' symbol size, regions, region size
If (rectangle = -1 Or (rectangle = -2 And (Application.Caller.MergeArea.Width * 3 > Application.Caller.MergeArea.Height * 4))) And el < 50 Then ' rectangular pattern ?
    k = Array(16, 7, 28, 11, 24, 14, 32, 18, 32, 24, 44, 28) ' symbol width, checkwords
    Do
        j = j + 1: w = k(j) ' width w/o finder pattern
        h = 6 + (j And 12) ' height
        l = w * h / 8: j = j + 1 ' # of bytes in symbol
    Loop While l - k(j) < el ' data fit in symbol ?
    If w > 25 Then nc = 2 ' column regions
Else ' square symbol
    w = 6: h = w
    i = 2 ' size increment
    k = Array(5, 7, 10, 12, 14, 18, 20, 24, 28, 36, 42, 48, 56, 68, 84, _
            112, 144, 192, 224, 272, 336, 408, 496, 620) ' checkwords
    Do
        If j = UBound(k) Then Err.Raise 514, "DataMatrix", "Message too long"
        j = j + 1
        If w > 11 * i Then i = 4 + i And 12 ' advance increment
        w = w + i: h = w
        l = (w * h) \ 8
    Loop While l - k(j) < el
    If w > 27 Then nr = 2 * (w \ 54) + 2: nc = nr ' regions
    If l > 255 Then b = 2 * (l \ 512) + 2 ' blocks
End If
s = k(j) ' checkwords
fw = w / nc: fh = h / nr ' region size

If el < l - s Then enc = enc + Chr(129): el = el + 1 ' first padding
Do While el < l - s ' add more padding
    el = el + 1
    enc = enc + Chr((((149 * el) Mod 253) + 130) Mod 254)
Loop

enc = enc + Space(s) ' compute Reed Solomon error detection and correction
Dim rs(70) As Integer, rc(70) As Integer ' RS code
Dim lg(256) As Integer, ex(255) As Integer ' log/exp table
s = s / b: j = 1
For i = 0 To 254
    ex(i) = j: lg(j) = i ' compute log/exp table of Galois field
    j = j + j: If j > 255 Then j = j Xor 301 ' GF polynomial a^8+a^5+a^3+a^2+1 = 100101101b = 301
Next i
rs(s + 1) = 0 ' compute RS generator polynomial
For i = 0 To s
    rs(s - i) = 1
    For j = s - i + 1 To s
        rs(j) = rs(j + 1) Xor ex((lg(rs(j)) + i) Mod 255)
    Next j
Next i
For c = 1 To b ' compute RS correction data for each block
    For i = 0 To s: rc(i) = 0: Next i
    For i = c To el Step b
        x = rc(0) Xor Asc(Mid(enc, i, 1))
        For j = 1 To s
            rc(j - 1) = rc(j) Xor IIf(x, ex((lg(rs(j)) + lg(x)) Mod 255), 0)
        Next j
    Next i
    For i = 0 To s - 1 ' add interleaved correction data
        Mid(enc, el + c + i * b, 1) = Chr(rc(i))
    Next i
Next c

With Application.Caller.Parent.Shapes
    b = .Count + 1 ' layout DataMatrix barcode
    For i = 0 To h + 2 * nr - 1 Step fh + 2 ' finder horizontal
        For j = 0 To w + 2 * nc - 1
            .AddShape(msoShapeRectangle, j, i + fh + 1, 1, 1).Name = Application.Caller.Address
            If (j And 1) = 0 Then .AddShape(msoShapeRectangle, j, i, 1, 1).Name = Application.Caller.Address
        Next j
    Next i
    For i = 0 To w + 2 * nc - 1 Step fw + 2 ' finder vertical
        For j = 0 To h - 1
            .AddShape(msoShapeRectangle, i, j + (j \ fh) * 2 + 1, 1, 1).Name = Application.Caller.Address
            If (j And 1) = 1 Then .AddShape(msoShapeRectangle, i + fw + 1, j + (j \ fh) * 2, 1, 1).Name = Application.Caller.Address
        Next j
    Next i
    'layout data
    s = 2: c = 0: r = 4 ' step,column,row of data position
    For i = 1 To l
        If (r = h - 3 And c = -1) Then ' corner A
            k = Array(w, 6 - h, w, 5 - h, w, 4 - h, w, 3 - h, w - 1, 3 - h, 3, 2, 2, 2, 1, 2)
        ElseIf r = h + 1 And c = 1 And (w And 7) = 0 And (h And 7) = 6 Then ' corner D
            k = Array(w - 2, -h, w - 3, -h, w - 4, -h, w - 2, -1 - h, w - 3, -1 - h, w - 4, -1 - h, w - 2, -2, -1, -2)
        Else
            If r = 0 And c = w - 2 And (w And 3) Then i = i - 1: GoTo continue ' corner B
            If r < 0 Or c >= w Or r >= h Or c < 0 Then ' outside
                s = -s: r = r + 2 + s / 2: c = c + 2 - s / 2 ' turn around
                Do While r < 0 Or c >= w Or r >= h Or c < 0
                    r = r - s: c = c + s
                Loop
            End If
            If r = h - 2 And c = 0 And (w And 3) Then ' corner B
                k = Array(w - 1, 3 - h, w - 1, 2 - h, w - 2, 2 - h, w - 3, 2 - h, w - 4, 2 - h, 0, 1, 0, 0, 0, -1)
            ElseIf r = h - 2 And c = 0 And (w And 7) = 4 Then ' corner C
                k = Array(w - 1, 5 - h, w - 1, 4 - h, w - 1, 3 - h, w - 1, 2 - h, w - 2, 2 - h, 0, 1, 0, 0, 0, -1)
            ElseIf r = 1 And c = w - 1 And (w And 7) = 0 And (h And 7) = 6 Then ' omit corner D
                i = i - 1: GoTo continue
            Else
                k = Array(0, 0, -1, 0, -2, 0, 0, -1, -1, -1, -2, -1, -1, -2, -2, -2) ' nominal layout
            End If
        End If
        el = Asc(Mid(enc, i, 1))
        For j = 0 To 15 Step 2 ' layout each bit
            If el And 1 Then
                x = c + k(j): y = r + k(j + 1)
                If x < 0 Then x = x + w: y = y + 4 - ((w + 4) And 7) ' wrap around
                If y < 0 Then y = y + h: x = x + 4 - ((h + 4) And 7)
                .AddShape(msoShapeRectangle, x + 2 * (x \ fw) + 1, y + 2 * (y \ fh) + 1, 1, 1).Name = Application.Caller.Address
            End If
            el = el \ 2
        Next j
continue:
        r = r - s: c = c + s ' diagonal step
    Next i
    For i = (w And -4) + 1 To w ' unfilled corner
        .AddShape(msoShapeRectangle, i, i, 1, 1).Name = Application.Caller.Address
    Next i
    b = .Count - b
    ReDim shps(b) As Integer   ' group all shapes
    For i = .Count To 1 Step -1
        If .Range(i).Name = Application.Caller.Address Then
            shps(b) = i: b = b - 1
            If b < 0 Then Exit For
        End If
    Next i
    s = 2 ' padding around symbol
    x = Application.Caller.MergeArea.Width * w / (w + s)
    y = Application.Caller.MergeArea.Height * h / (h + s) * (w + s) / (h + s)
    If x > y Then x = y
    With .Range(shps).Group
        .Fill.ForeColor.RGB = fColor ' format barcode shape
        .line.ForeColor.RGB = bColor
        .line.Weight = line
        .Width = x ' fit symbol in excel cell
        .Height = .Width * (h + s) / (w + s)
        .Left = Application.Caller.Left + (Application.Caller.MergeArea.Width - .Width) / 2
        .Top = Application.Caller.Top + (Application.Caller.MergeArea.Height - .Height) / 2
        .Name = Application.Caller.Address ' link shape to data
        .Title = text
        .AlternativeText = "DataMatrix barcode, " & (w + 2 * nc) & "x" & (h + 2 * nr) & " cells"
        .LockAspectRatio = True
        .Placement = xlMove
    End With
End With
failed:
If Err.Number Then DataMatrix = "ERROR DataMatrix: " & Err.Description
End Function
