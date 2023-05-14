Attribute VB_Name = "modAddInMenu"
Option Explicit
Dim BarcodeList As Variant

'*********************************XML CODE*********************************
'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
'   <ribbon>
'      <tabs>
'         <tab idMso="TabFormulas">
'            <group id="BarcodeLib" label="Barcode Function Library">
'               <gallery id="Barcode"
'                   label="Barcode Functions" columns="1"
'                   imageMso = "GroupFunctionLibrary"
'                   getItemCount = "Barcode_getItemCount"
'                   getItemLabel = "Barcode_getItemLabel"
'                   getItemScreentip = "Barcode_getItemScreentip"
'                   getItemSupertip = "Barcode_getItemSupertip"
'                   onAction = "Barcode_Click"
'                   showItemLabel = "true"
'                   size="large">
'                 <button id="insertFx"
'                    imageMso = "GroupFunctionLibrary"
'                    label = "Insert Function"
'                    screentip="Insert Function (Shift+F3)"
'                    supertip = "Work with the formula in the current cell. You can easily pick functions to use and get help on how to fill out the input values."
'                    onAction="insertFx_Click"/>
'               </gallery>
'               <button id="updateBarcodes"
'                   imageMso = "ConnectedToolSyncMenu"
'                   label = "Update Barcodes"
'                   screentip="Update Barcodes"
'                   supertip = "Update all barcodes in current workbook."
'                   onAction = "updateBarcodes_Click"
'                   size="large"/>
'               <button id="getHelp"
'                   imageMso = "Help"
'                   label = "Help"
'                   screentip="Help"
'                   supertip = "Open link to webpage."
'                   onAction = "getHelp_Click"
'                   size="large"/>
'            </group>
'         </tab>
'      </tabs>
'   </ribbon>
'</customUI>
'*********************************XML CODE*********************************

Private Sub AddInMenuProperties()
    ' Custom function for changing file properties (not used during run time)
    ActiveWorkbook.BuiltinDocumentProperties("Title").Value = "Barcode Fx 3.0"
    ActiveWorkbook.BuiltinDocumentProperties("Comments").Value = "Function library for barcodes"
End Sub

Sub Auto_Open()

    ' Populate BarcodeList
    BarcodeList = Array("Aztec()", "Code11()", "Code39()", "Code93()", "Code128()", _
                        "DataMatrix()", "EAN_2()", "EAN_5()", "EAN_13()", "ITF()", _
                        "ITF_14()", "PDF_417()", "QRCode()", "UPCA()", "UPCE()")

End Sub

'update all barcodes in active sheet
Sub updateBarcodes_Click(control As IRibbonControl)
    Dim shp As Shape, bc As Variant, str As String
    On Error Resume Next
    For Each shp In ActiveSheet.Shapes ' delete all lost barcode shapes
        If shp.Type = msoGroup Then
            str = LCase(shp.AlternativeText)
            For Each bc In Array("aztec", "code128", "datamatrix", "pdf417", "quickresponse barcode", "linear barcode")
                If Left(str, Len(bc)) = bc Then
                    shp.Title = "" ' force redraw
                    If InStr(LCase(Range(shp.Name).Formula), bc) = 0 Then shp.Delete
                    Exit For
                End If
            Next bc
        End If
    Next shp
    ' refresh all barcodes
    Application.CalculateFull
    'Call Kanji
End Sub

Sub getHelp_Click(control As IRibbonControl)

    Dim URL As String
    
    URL = "https://github.com/EszopiCoder/excel-barcode-fx-library"
    
    If MsgBox("You are leaving Microsoft Excel to the following website: " & URL & _
    vbNewLine & vbNewLine & "Would you like to proceed?", _
    vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    ActiveWorkbook.FollowHyperlink URL

End Sub

Sub Barcode_getItemCount(control As IRibbonControl, ByRef returnedVal)
    ' Return the number of functions in the array
    returnedVal = UBound(BarcodeList) - LBound(BarcodeList) + 1
End Sub

Sub Barcode_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error Resume Next
    ' Return the name of the function without arguments
    returnedVal = Left(BarcodeList(index), InStr(1, BarcodeList(index), "(") - 1)
    On Error GoTo 0
End Sub

Sub Barcode_getItemScreentip(control As IRibbonControl, index As Integer, ByRef returnedVal)
    On Error Resume Next
    ' Return the name of the function with arguments
    returnedVal = BarcodeList(index)
    On Error GoTo 0
End Sub

Sub Barcode_getItemSupertip(control As IRibbonControl, index As Integer, ByRef returnedVal)
    Dim Supertip As Variant
    Supertip = _
    Array("Draw Aztec barcode.", _
        "Draw Code 11 barcode.", _
        "Draw Code 39 barcode.", _
        "Draw Code 93 barcode.", _
        "Draw Code 128 barcode.", _
        "Draw DataMatrix barcode.", _
        "Draw EAN-2 barcode.", _
        "Draw EAN-5 barcode.", _
        "Draw EAN-13 barcode.", _
        "Draw ITF barcode.", _
        "Draw ITF-14 barcode.", _
        "Draw PDF417 barcode.", _
        "Draw QR code.", _
        "Draw UPC-A or EAN-8 barcode.", _
        "Draw UPC-E barcode.")

    On Error Resume Next
    returnedVal = Supertip(index)
    On Error GoTo 0
End Sub

Sub insertFx_Click(control As IRibbonControl)

    ActiveCell.FunctionWizard

End Sub

Sub Barcode_Click(control As IRibbonControl, id As String, index As Integer)
    On Error Resume Next
    ' Insert function into active cell (same as the other built-in functions)
    If InStr(1, ActiveCell.Formula, "=") > 0 Then
        ActiveCell.Formula = ActiveCell.Formula & "+" & BarcodeList(index)
    Else
        ActiveCell.Formula = "=" & BarcodeList(index)
    End If
    ' Open function wizard dialog. Clear cell if user hits cancel button.
    If Application.Dialogs(xlDialogFunctionWizard).Show = False Then
        ActiveCell.Formula = ""
    End If
    On Error GoTo 0
End Sub
