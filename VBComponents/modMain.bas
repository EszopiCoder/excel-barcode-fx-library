Attribute VB_Name = "modMain"
' Barcode symbol creation by VBA
' Author: alois zingl
' Version: V1.1 jan 2016
' Copyright: Free and open-source software
' http://members.chello.at/~easyfilter/barcode.html
' Description: the indention of this library is a short and compact implementation to create barcodes
'  of Code 128, Data Matrix, (micro) QR, PDF417, or Aztec symbols so it could be easily adapted for individual requirements.
'  The Barcode is drawn as shape in the cell of the Excel sheet.
'  The smallest bar code symbol fitting the data is automatically selected,
'  but no size optimization for mixed data types in one code is done.
' Functions:
'   DataMatrix(text As String, Optional rectangle As Integer)
'   QuickResponse(text As String, Optional level As String = "L", Optional version As Integer = 1)
'   Aztec(text As String, Optional security As Integer, Optional layers As Integer = 1)
'   Code128(text As String)
'   PDFIVXVII(text As String, Optional security = -1, Optional nbcol = 1, Optional CodeErr)
'
Option Explicit

' convert UTF-16 (Windows) to UTF-8
Public Function utf16to8(text As String) As String
Dim i As Integer, c As Long
utf16to8 = text
For i = Len(text) To 1 Step -1
    c = AscW(Mid(text, i, 1)) And 65535
    If c > 127 Then
        If c > 4095 Then
            utf16to8 = Left(utf16to8, i - 1) + Chr(224 + c \ 4096) + Chr(128 + (c \ 64 And 63)) + Chr(128 + (c And 63)) & Mid(utf16to8, i + 1)
        Else
            utf16to8 = Left(utf16to8, i - 1) + Chr(192 + c \ 64) + Chr(128 + (c And 63)) & Mid(utf16to8, i + 1)
        End If
    End If
Next i
End Function

' read/write kanji conversion string from/to file
Public Sub Kanji()
Dim p As Variant, s As Worksheet, k1 As String, c As Long
Const K = "kanji" ' property name
For Each s In Application.ThisWorkbook.Worksheets
    For Each p In s.CustomProperties ' look for kanji conversion string
        If p.Name = K Then If Len(p.Value) > 10000 Then k1 = p.Value
    Next p
Next s
ChDir Application.ThisWorkbook.Path
If k1 = "" Then  ' not found, get from file
    p = Application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm", 1, "Read Kanji Conversion String for QRCodes from 'barcode.xlsm'")
    If p <> False Then
        Application.ScreenUpdating = False
        With Workbooks.Open(p, 0, True)
            For Each s In .Worksheets
                For Each p In s.CustomProperties ' look for kanji conversion string
                    If p.Name = K Then If Len(p.Value) > 10000 Then k1 = p.Value
                Next p
            Next s
            .Close
        End With
        Application.ScreenUpdating = True
        If Len(k1) < 10000 Or (Len(k1) And 1) Then MsgBox "No Kanji conversion string for QRCodes found in Excel file."
        For Each s In Application.ThisWorkbook.Worksheets
            c = 0
            For Each p In s.CustomProperties ' look for kanji conversion string
                If p.Name = K Then p.Value = k1: c = 1
            Next p
            If c = 0 Then s.CustomProperties.Add K, k1
        Next s
    End If
End If
End Sub

Public Function DRAWLINEAR(barcode As String) As String
    On Error GoTo failed
    If Not TypeOf Application.Caller Is Range Then Err.Raise 513, "Linear Barcode", "Call only from sheet"
    Dim i, j, K As Integer
    Dim x As Integer
    Dim shp As Shape, color As Long, txt As String ' redraw barcode ?
    color = vbBlack
    For Each shp In Application.Caller.Parent.Shapes
        If shp.Name = Application.Caller.Address Then
            If shp.Title = barcode Then Exit Function ' same as prev ?
            color = shp.Fill.ForeColor.RGB ' redraw with same color
            shp.line.Visible = msoFalse
            shp.Delete
        End If
    Next shp
    
    'Validate barcode
    If IsNumeric(barcode) = False Then
        DRAWLINEAR = barcode
        Exit Function
    End If
    
    With Application.Caller.Parent.Shapes
        K = .count + 1
        x = 0
        For i = 1 To Len(barcode) Step 1
            If i Mod 2 <> 0 Then
                .AddShape(msoShapeRectangle, x, 1, Mid(barcode, i, 1), 1).Name = Application.Caller.Address
            End If
            x = x + Mid(barcode, i, 1)
        Next i
        K = .count - K
        ReDim shps(K) As Integer   ' group all shapes
        For j = .count To 1 Step -1
            If .Range(j).Name = Application.Caller.Address Then
                shps(K) = j: K = K - 1
                If K < 0 Then Exit For
            End If
        Next j
        
        With .Range(shps).Group
            .Fill.ForeColor.RGB = color ' format barcode shape
            .line.Visible = msoFalse
            .Width = Application.Caller.MergeArea.Width * (calcModules(barcode) - 18) / calcModules(barcode) ' fit symbol in excel cell with padding
            .Height = Application.Caller.MergeArea.Height
            .Left = Application.Caller.Left + (Application.Caller.MergeArea.Width - .Width) / 2
            .Top = Application.Caller.Top
            .Name = Application.Caller.Address ' link shape to data
            .Title = barcode
            .AlternativeText = "Linear barcode, " & calcModules(barcode) & " modules"
            .LockAspectRatio = True
            .Placement = xlMove
        End With
        
    End With
    
failed:
    If Err.Number Then DRAWLINEAR = "ERROR DRAWLINEAR: " & Err.Description
End Function

Private Function calcModules(source As String) As Long
    Dim i As Long
    Dim count As Long
    
    count = 0
    For i = 1 To Len(source) Step 1
        count = count + Int(Mid(source, i, 1))
    Next i
    
    calcModules = count
    
End Function
