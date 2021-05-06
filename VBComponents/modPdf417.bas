Attribute VB_Name = "modPdf417"
Option Explicit
'http://www.onbarcode.com/pdf417/pdf417_size_setting.html

Private Sub TestPDF417()
Dim strTest As String
strTest = PDF417String("test", -1, 1)
'Debug.Print strTest
Debug.Print PDF417ToBinary(strTest)(0)
Call PDF_417("test")
End Sub

Public Function PDF_417(text As String, Optional security = -1, Optional nbcol = 1, Optional CodeErr) As String
On Error GoTo failed
If Not TypeOf Application.Caller Is Range Then Err.Raise 513, "PDF417", "Call only from sheet"
Dim h As Long, w As Long, x As Long, y As Long
Dim K As Long, i As Long
Dim shp As Shape
Dim fColor As Long
Dim BinPDF417

fColor = vbBlack ' redraw graphic ?
For Each shp In Application.Caller.Parent.Shapes
    If shp.Name = Application.Caller.Address Then
        If shp.Title = text Then Exit Function ' same as prev ?
        fColor = shp.Fill.ForeColor.RGB  ' remember format
        shp.line.Visible = msoFalse
        shp.Delete
    End If
Next shp

' Get binary of PDF417
BinPDF417 = PDF417ToBinary(PDF417String(text, security, nbcol, CodeErr))
h = UBound(BinPDF417)
w = Len(BinPDF417(0))

With Application.Caller.Parent.Shapes
    K = .count + 1 ' layout PDF417
    For y = 0 To h
        For x = 1 To w
            If Mid(BinPDF417(y), x, 1) = 1 Then ' apply mask
                .AddShape(msoShapeRectangle, x, 3 * y, 1, 3).Name = Application.Caller.Address
            End If
        Next x
    Next y
    K = .count - K
    ReDim shps(K) As Integer   ' group all shapes
    For i = .count To 1 Step -1
        If .Range(i).Name = Application.Caller.Address Then
            shps(K) = i: K = K - 1
            If K < 0 Then Exit For
        End If
    Next i

    With .Range(shps).Group
        .Fill.ForeColor.RGB = fColor ' format barcode shape
        .line.Visible = msoFalse
        x = Application.Caller.MergeArea.Width
        y = Application.Caller.MergeArea.Height
        .Width = x * w / (w + 2) ' fit symbol in excel cell
        .Height = y * h / (h + 2) ' fit symbol in excel cell
        .Left = Application.Caller.Left + (Application.Caller.MergeArea.Width - .Width) / 2
        .Top = Application.Caller.Top + (Application.Caller.MergeArea.Height - .Height) / 2
        .Name = Application.Caller.Address ' link shape to data
        .Title = text
        .AlternativeText = "PDF417 barcode, security " & security & ", nbcol " & nbcol & ", " & h & "x" & w & " cells"
        .LockAspectRatio = True
        .Placement = xlMove
    End With
End With
failed:
If Err.Number Then PDF_417 = "ERROR PDF417: " & Err.Description
End Function
