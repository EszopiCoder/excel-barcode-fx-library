Attribute VB_Name = "ModulBarcode"
' Barcode symbol creation by VBA
' Author: alois zingl
' Version: V1.1 jan 2016
' Copyright: Free and open-source software
' http://members.chello.at/~easyfilter/barcode.html
' Description: the indention of this library is a short and compact implementation to create barcodes
'  of Code 128, Data Matrix, (micro) QR or Aztec symbols so it could be easily adapted for individual requirements.
'  The Barcode is drawn as shape in the cell of the Excel sheet.
'  The smallest bar code symbol fitting the data is automatically selected,
'  but no size optimization for mixed data types in one code is done.
' Functions:
'   DataMatrix(text As String, Optional rectangle As Integer)
'   QuickResponse(text As String, Optional level As String = "L", Optional version As Integer = 1)
'   Aztec(text As String, Optional security As Integer, Optional layers As Integer = 1)
'   Code128(text As String)
'
Option Explicit

' add description to user defined barcode functions
Private Sub Workbook_Open()
ReDim arg(0) As String
arg(0) = "text to encode"
Application.MacroOptions macro:="Code128", Description:="Draw Code 128 barcode", Category:="Barcode", ArgumentDescriptions:=arg
Application.MacroOptions macro:="DataMatrix", Description:="Draw DataMatrix barcode", Category:="Barcode", ArgumentDescriptions:=arg
ReDim Preserve arg(2)
arg(1) = "percentage of checkwords (1..90)" + vbCrLf + "number, optional, default 23%"
arg(2) = "minimum number of layers (0-32)" + vbCrLf + "number, optional, default 1" + vbCrLf + "set to 0 for Aztec rune"
Application.MacroOptions macro:="Aztec", Description:="Draw Aztec barcode", Category:="Barcode", ArgumentDescriptions:=arg
arg(1) = "security level ""LMQH""" + vbCrLf + "low, medium, quartile, high" + vbCrLf + "letter, optional, default L"
arg(2) = "minimum version size(-3..40)" + vbCrLf + "number, optional, default 1" + vbCrLf + "MircoQR M1:-3, M2:-2, M3:-1, M4:0"
Application.MacroOptions macro:="QRCode", Description:="Draw QR code", Category:="Barcode", ArgumentDescriptions:=arg
End Sub

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

'update all barcodes in active sheet
Public Sub updateBarcodes()
Attribute updateBarcodes.VB_Description = "Updates all barcode shapes of the actual sheet."
Attribute updateBarcodes.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim shp As Shape, bc As Variant, str As String
On Error Resume Next
For Each shp In ActiveSheet.Shapes ' delete all lost barcode shapes
    If shp.Type = msoAutoShape Then
        str = LCase(shp.AlternativeText)
        For Each bc In Array("aztec", "code128", "datamatrix", "qrcode")
            If Left(str, Len(bc)) = bc Then
                shp.Title = "" ' force redraw
                If InStr(LCase(Range(shp.Name).Formula), bc) = 0 Then shp.Delete
                Exit For
            End If
        Next bc
    End If
Next shp
Application.CalculateFull ' refresh all barcodes
Kanji
End Sub

' read/write kanji conversion string from/to file
Public Sub Kanji()
Dim p As Variant, s As Worksheet, k1 As String, c As Long
Const k = "kanji" ' property name
For Each s In Application.ThisWorkbook.Worksheets
    For Each p In s.CustomProperties ' look for kanji conversion string
        If p.Name = k Then If Len(p.Value) > 10000 Then k1 = p.Value
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
                    If p.Name = k Then If Len(p.Value) > 10000 Then k1 = p.Value
                Next p
            Next s
            .Close
        End With
        Application.ScreenUpdating = True
        If Len(k1) < 10000 Or (Len(k1) And 1) Then MsgBox "No Kanji conversion string for QRCodes found in Excel file."
        For Each s In Application.ThisWorkbook.Worksheets
            c = 0
            For Each p In s.CustomProperties ' look for kanji conversion string
                If p.Name = k Then p.Value = k1: c = 1
            Next p
            If c = 0 Then s.CustomProperties.Add k, k1
        Next s
    End If
End If
End Sub
