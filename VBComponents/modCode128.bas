Attribute VB_Name = "modCode128"
Option Explicit

' Code 128 symbol creation according ISO/IEC 15417:2007
Public Function Code128(text As String) As String
Attribute Code128.VB_Description = "Draw Code 128 barcode"
Attribute Code128.VB_ProcData.VB_Invoke_Func = " \n18"
On Error GoTo failed
If Not TypeOf Application.Caller Is Range Then Err.Raise 513, "Code 128", "Call only from sheet"
Dim m As Long, i As Long, j As Long, c As Long, l As Long, t As Long
Dim shp As Shape, color As Long, txt As String ' redraw barcode ?
color = vbBlack
For Each shp In Application.Caller.Parent.Shapes
    If shp.Name = Application.Caller.Address Then
        If shp.Title = text Then Exit Function ' same as prev ?
        color = shp.Fill.ForeColor.RGB ' redraw with same color
        shp.Delete
    End If
Next shp
txt = utf16to8(text): m = 3: l = 0: t = Len(txt)
ReDim enc(3 * t + 3) As Byte
For i = 1 To t
    If m <> 2 Then ' alpha mode
        For j = 0 To t - i ' count digits
            If Not IsNumeric(Mid(txt, i + j, 1)) Then Exit For
        Next j
        If (j > 1 And i = 1) Or (j > 3 And (i + j < t Or (j And 1) = 0)) Then
            enc(l) = IIf(i = 1, 105, 99) ' start / code C
            l = l + 1: m = 2 ' to digit
        End If
    End If
    If m = 2 Then ' digit mode
        If IsNumeric(Mid(txt, i, 1)) And IsNumeric(Mid(txt, i + 1, 1)) Then
            enc(l) = val(Mid(txt, i, 2)) ' two digits
            l = l + 1: i = i + 1
        Else
            m = 3 ' exit digit
        End If
    End If
    If m <> 2 Then ' alpha mode
        c = Asc(Mid(txt, i, 1))
        If m > 2 Or ((c And 127) < 32 And m) Or ((c And 127) > 95 And m = 0) Then  ' change ?
            For j = IIf(m > 2 Or i + 1 = t, i, i + 1) To t - 1 ' A or B needed?
                If Asc(Mid(txt, j, 1)) - 32 And 64 Then Exit For ' < 32 or > 95
            Next j
            j = IIf(Asc(Mid(txt, j, 1)) And 96, 1, 0) ' new set
            enc(l) = IIf(i = 1, 103 + j, IIf(j <> m, 101 - j, 98))
            l = l + 1: m = j ' change set: start,code,(shift)
        End If
        If c > 127 Then enc(l) = 101 - m: l = l + 1 ' FNC4: char > 127
        enc(l) = ((c And 127) + 64) Mod 96: l = l + 1
    End If
Next i
If i = 1 Then enc(0) = 103: l = 1 ' empty message
j = enc(0) ' check sum
For i = 1 To l
    j = j + i * enc(i)
Next i
enc(l) = j Mod 103: enc(l + 1) = 106 ' stop

With Application.Caller.Parent.Shapes
    For i = 0 To l + 1 ' code to pattern
        c = Array(277, 337, 341, 69, 73, 133, 84, 88, 148, 324, 328, 388, 22, 82, 86, 37, 97, _
            101, 356, 322, 326, 292, 352, 530, 517, 577, 581, 532, 592, 596, 273, 281, 401, 9, _
            129, 137, 24, 144, 152, 264, 384, 392, 18, 26, 146, 33, 41, 161, 545, 266, 386, 288, _
            296, 290, 513, 521, 641, 528, 536, 656, 560, 332, 896, 5, 13, 65, 77, 193, 197, 20, 28, _
            80, 92, 208, 212, 452, 320, 800, 448, 176, 7, 67, 71, 52, 112, 116, 772, 832, 836, 275, _
            305, 785, 3, 11, 131, 48, 56, 768, 776, 35, 50, 515, 770, 268, 260, 262, 416)(enc(i))
        m = c \ 256 + 1
        .AddShape(msoShapeRectangle, 11 * i, 0, m, 1).Name = Application.Caller.Address ' 1st bar
        j = 11 * i + m + ((c \ 64) And 3) + 1
        m = ((c \ 16) And 3) + 1
        .AddShape(msoShapeRectangle, j, 0, m, 1).Name = Application.Caller.Address ' 2nd bar
        j = j + m + ((c \ 4) And 3) + 1
        .AddShape(msoShapeRectangle, j, 0, (c And 3) + 1, 1).Name = Application.Caller.Address ' 3rd bar
    Next i
    .AddShape(msoShapeRectangle, 11 * i, 0, 2, 1).Name = Application.Caller.Address ' stop bar
    j = 3 * l + 6: m = j

    ReDim shps(j) As Integer ' group all shapes
    For i = .Count To 1 Step -1
        If .Range(i).Name = Application.Caller.Address Then
            shps(j) = i: j = j - 1
            If j < 0 Then Exit For
        End If
    Next i
    With .Range(shps).Group
        .Fill.ForeColor.RGB = color ' format barcode shape
        .line.Visible = False
        .Width = Application.Caller.MergeArea.Width * 2 * m / (2 * m + 1) ' fit symbol in excel cell
        .Height = Application.Caller.MergeArea.Height - .Width / (2 * m)
        .Left = Application.Caller.Left + (Application.Caller.MergeArea.Width - .Width) / 2
        .Top = Application.Caller.Top + (Application.Caller.MergeArea.Height - .Height) / 2
        .Name = Application.Caller.Address ' link shape to data
        .Title = text
        .AlternativeText = "Code128 barcode, " & (l + 2) & " characters"
    End With
End With
failed:
If Err.Number Then Code128 = "ERROR Code128: " & Err.Description
End Function
