VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Add macro descriptions
    Application.MacroOptions "Aztec", "Draw Aztec barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.", _
        "OPTIONAL Percentage of checkwords used for security (1% to 90%) [Default=23%].", _
        "OPTIONAL Number of layers (0 to 32) [Default=1]")
    
    Application.MacroOptions "Code11", "Encode Code 11 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
    
    Application.MacroOptions "Code39", "Encode Code 39 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.", _
        "OPTIONAL Include check digit [Default=False]")
    
    Application.MacroOptions "Code93", "Encode Code 93 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
    
    Application.MacroOptions "Code128", "Draw Code 128 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
        
    Application.MacroOptions "EAN_2", "Encode EAN-2 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
    
    Application.MacroOptions "EAN_5", "Encode EAN-5 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
    
    Application.MacroOptions "EAN_13", "Encode EAN-13 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
    
    Application.MacroOptions "DataMatrix", "Draw DataMatrix barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")

    Application.MacroOptions "ITF", "Encode ITF barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
        
    Application.MacroOptions "ITF_14", "Encode ITF-14 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
    
    Application.MacroOptions "PDF_417", "Draw PDF417 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.", _
        "OPTIONAL Security level [Default=-1]", _
        "OPTIONAL Number of columns [Default=1]", _
        "OPTIONAL Code Error")
    
    Application.MacroOptions "QRCode", "Draw QR code.", , , , , "Barcode", , , , _
        Array("Text to be encoded.", _
        "OPTIONAL Security level ""LMQH""" + vbCrLf + "low, medium, quartile, high" + vbCrLf + "[Default=L]", _
        "OPTIONAL Minimum version size(-3 to 40) [Default=1]" + vbCrLf + "MicroQR M1:-3, M2:-2, M3:-1, M4:0")
        
    Application.MacroOptions "UPCA", "Encode UPC-A or EAN-8 barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
        
    Application.MacroOptions "UPCE", "Encode UPC-E barcode.", , , , , "Barcode", , , , _
        Array("Text to be encoded.")
        
    Application.MacroOptions "DRAWLINEAR", "Draw linear barcode.", , , , , "Barcode", , , , _
        Array("Barcode string to be drawn.")
End Sub

Private Sub Workbook_Open()
    Set App = Application
End Sub

Private Sub Workbook_AddinInstall()
    MsgBox "Thank you for installing the Barcode Fx Library." & vbNewLine & _
        "This is an open source project created by EszopiCoder." & vbNewLine & _
        "(https://github.com/EszopiCoder/excel-barcode-fx-library)", vbInformation
End Sub

Private Sub Workbook_AddinUninstall()
    MsgBox "Sorry to hear you will no longer be using the 'Barcode Fx Library'." & vbNewLine & _
        "Please send suggestions for improvement to pharm.coder@gmail.com", vbInformation
End Sub


