Attribute VB_Name = "modupcean"
Option Explicit

'GS1 EAN/UPC barcodes: https://www.gs1.org/standards/barcodes/ean-upc

'Wikipedia References
'UPC-A & UPC-E: https://en.wikipedia.org/wiki/Universal_Product_Code
'EAN-13: https://en.wikipedia.org/wiki/International_Article_Number
'EAN-8: https://en.wikipedia.org/wiki/EAN-8
'EAN-5: https://en.wikipedia.org/wiki/EAN-5
'EAN-2: https://en.wikipedia.org/wiki/EAN-2

Dim UPCParity0
Dim UPCParity1
Dim EAN2Parity
Dim EAN5Parity
Dim EAN13Parity
Dim EANsetA
Dim EANsetB

Private Sub upceanLibraries()
    UPCParity0 = Array("BBBAAA", "BBABAA", "BBAABA", "BBAAAB", "BABBAA", "BAABBA", "BAAABB", _
        "BABABA", "BABAAB", "BAABAB") 'Number set for UPC-E symbol (EN Table 4)
    UPCParity1 = Array("AAABBB", "AABABB", "AABBAB", "AABBBA", "ABAABB", "ABBAAB", "ABBBAA", _
        "ABABAB", "ABABBA", "ABBABA") 'Not covered by BS EN 797:1995
    EAN2Parity = Array("AA", "AB", "BA", "BB") 'Number sets for 2-digit add-on (EN Table 6)
    EAN5Parity = Array("BBAAA", "BABAA", "BAABA", "BAAAB", "ABBAA", "AABBA", "AAABB", "ABABA", _
        "ABAAB", "AABAB") 'Number set for 5-digit add-on (EN Table 7)
    EAN13Parity = Array("AAAAA", "ABABB", "ABBAB", "ABBBA", "BAABB", "BBAAB", "BBBAA", "BABAB", _
        "BABBA", "BBABA") 'Left hand of the EAN-13 symbol (EN Table 3)
    EANsetA = Array("3211", "2221", "2122", "1411", "1132", "1231", "1114", "1312", "1213", _
        "3112") 'Representation set A and C (EN Table 1)
    EANsetB = Array("1123", "1222", "2212", "1141", "2311", "1321", "4111", "2131", "3121", _
        "2113") 'Representation set B (EN Table 1)
End Sub

Public Function GS1_CHECK(source As String) As String
    'Calculate the correct check digit for a UPC barcode
    'Source: https://www.gs1.org/services/how-calculate-check-digit-manually
    Dim i As Integer
    Dim count As Integer
    Dim CHECK_DIGIT As Integer
    
    'Loop through each digit to get sum
    count = 0
    For i = 1 To Len(source) Step 1
        count = count + Mid(source, i, 1)
        'Length is even and even-numbered digit positions (2nd, 4th, 6th, etc.)
        'Length is odd and odd-numbered digit positions (1st, 3rd, 5th, etc.)
        If (Len(source) Mod 2 = 0 And i Mod 2 = 0) Or _
            (Len(source) Mod 2 <> 0 And i Mod 2 <> 0) Then
            count = count + (2 * Mid(source, i, 1))
        End If
    Next i
    
    'Calculate check digit
    CHECK_DIGIT = 10 - (count Mod 10)
    If CHECK_DIGIT = 10 Then CHECK_DIGIT = 0
    GS1_CHECK = CHECK_DIGIT
End Function

'UPC-A and EAN-8
Function UPCA(source As String) As String
Attribute UPCA.VB_Description = "Encode UPC-A or EAN-8 barcode."
Attribute UPCA.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim i As Integer
    Dim half_way As Integer
    Dim dest As String
    
    'Validate source
    If Len(source) <> 12 And Len(source) <> 11 And Len(source) <> 8 Then
        UPCA = "Improper EAN-8 or UPC-A barcode length (8 or 11-12 digits)"
        Exit Function
    ElseIf Len(source) = 12 And GS1_CHECK(Left(source, 11)) <> Right(source, 1) Then
        UPCA = "Invalid check digit (" & GS1_CHECK(Left(source, 11)) & ")"
        Exit Function
    End If
    
    'Calculate check digit (UPC-A only)
    If Len(source) = 11 Then source = source & GS1_CHECK(source)
    
    Call upceanLibraries

    half_way = Len(source) / 2
    'Start character
    dest = "111"
    'Left/Middle/Right characters
    For i = 1 To Len(source) Step 1
        If i = half_way + 1 Then
            dest = dest & "11111"
        End If
        dest = dest & EANsetA(Mid(source, i, 1))
    Next i
    'End character
    dest = dest & "111"

    UPCA = DRAWLINEAR(dest)
    
End Function

Function UPCE(source As String) As String
Attribute UPCE.VB_Description = "Encode UPC-E barcode."
Attribute UPCE.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim i As Integer
    Dim emode As Integer
    Dim equivalent As String
    Dim CHECK_DIGIT As Integer
    Dim parity As String
    Dim dest As String
    
    'Two number systems can be used - system 0 and system 1
    If Len(source) = 7 Then
        If Left(source, 1) > 1 Then _
            source = "0" & Right(source, 6)
    ElseIf Len(source) = 6 Then 'Default number system is 0
        source = "0" & source
    Else
        UPCE = "Improper UPC-E barcode length (6-7 digits)"
        Exit Function
    End If
    
    Call upceanLibraries
    
    'Expand the zero-compressed UPCE code to make a UPCA equivalent (EN Table 5)
    emode = Right(source, 1)
    equivalent = Left(source, 3)
    Select Case emode
        Case 0 To 2
            equivalent = equivalent & emode & "0000" & Mid(source, 4, 3)
        Case 3
            equivalent = equivalent & Mid(source, 4, 1) & "00000" & Mid(source, 5, 2)
        Case 4
            equivalent = equivalent & Mid(source, 4, 2) & "00000" & Mid(source, 6, 1)
        Case 5 To 9
            equivalent = equivalent & Mid(source, 4, 3) & "0000" & emode
    End Select
    'Calculate check digit
    CHECK_DIGIT = GS1_CHECK(equivalent)
    equivalent = equivalent & CHECK_DIGIT
    
    'Use number system and check digit to choose a parity scheme
    If Left(equivalent, 1) = 1 Then
        parity = UPCParity1(CHECK_DIGIT)
    Else
        parity = UPCParity0(CHECK_DIGIT)
    End If
    
    'Start character
    dest = "111"
    'Middle characters
    For i = 1 To Len(parity) Step 1
        Select Case Mid(parity, i, 1)
            Case "A"
                dest = dest & EANsetA(Mid(source, i + 1, 1))
            Case "B"
                dest = dest & EANsetB(Mid(source, i + 1, 1))
        End Select
    Next i
    'End character
    dest = dest & "111111"
    
    UPCE = DRAWLINEAR(dest)
    
End Function

Function EAN_13(source As String) As String
Attribute EAN_13.VB_Description = "Encode EAN-13 barcode."
Attribute EAN_13.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim parity As String
    Dim half_way As Integer
    Dim i As Integer
    Dim dest As String
    
    'Validate source
    If Len(source) > 13 Or Len(source) < 12 Then
        EAN_13 = "Improper UPC-A barcode length (12-13 digits)"
        Exit Function
    ElseIf Len(source) = 13 And GS1_CHECK(Left(source, 12)) <> Right(source, 1) Then
        EAN_13 = "Invalid check digit (" & GS1_CHECK(Left(source, 12)) & ")"
        Exit Function
    End If
    
    'Calculate check digit
    If Len(source) = 12 Then source = source & GS1_CHECK(source)
    
    Call upceanLibraries
    
    'Get parity for first half of symbol
    parity = EAN13Parity(Left(source, 1))
    
    half_way = 8
    
    'Start character
    dest = "111"
    'Middle characters
    For i = 2 To Len(source) Step 1
        If i = half_way Then dest = dest & "11111"
        If i > 2 And i < 8 Then
            If Mid(parity, i - 2, 1) = "B" Then
                dest = dest & EANsetB(Mid(source, i, 1))
            Else
                dest = dest & EANsetA(Mid(source, i, 1))
            End If
        Else
            dest = dest & EANsetA(Mid(source, i, 1))
        End If
    Next i
    'End character
    dest = dest & "111"
    
    EAN_13 = DRAWLINEAR(dest)
    
End Function

Function EAN_5(source As String) As String
Attribute EAN_5.VB_Description = "Encode EAN-5 barcode."
Attribute EAN_5.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim i As Integer
    Dim parity_sum As Integer
    Dim parity As String
    Dim dest As String
    
    'Validate source
    If Len(source) <> 5 Then
        EAN_5 = "Improper EAN-5 barcode length (5 digits)"
    End If
    
    Call upceanLibraries
    
    'Determine parity
    parity_sum = 3 * (Int(Left(source, 1)) + Int(Mid(source, 3, 1)) + Int(Right(source, 1)))
    parity_sum = parity_sum + (9 * (Int(Mid(source, 2, 1)) + Int(Mid(source, 4, 1))))
    parity_sum = parity_sum Mod 10
    parity = EAN5Parity(parity_sum)
    
    'Start character
    dest = "112"
    'Middle characters
    For i = 1 To Len(parity) Step 1
        Select Case Mid(parity, i, 1)
            Case "A"
                dest = dest & EANsetA(Mid(source, i, 1))
            Case "B"
                dest = dest & EANsetB(Mid(source, i, 1))
        End Select
        If i <> Len(parity) Then
            dest = dest & "11"
        End If
    Next i
    
    EAN_5 = DRAWLINEAR(dest)
    
End Function

Function EAN_2(source As String) As String
Attribute EAN_2.VB_Description = "Encode EAN-2 barcode."
Attribute EAN_2.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim i As Integer
    Dim parity_sum As Integer
    Dim parity As String
    Dim dest As String
    
    'Validate source
    If Len(source) <> 2 Then
        EAN_2 = "Improper EAN-2 barcode length (2 digits)"
    End If
    
    Call upceanLibraries
    
    'Determine parity
    parity_sum = Left(source, 1) * 10 + Right(source, 1)
    parity_sum = parity_sum Mod 4
    parity = EAN2Parity(parity_sum)
    
    'Start character
    dest = "112"
    'Middle characters
    For i = 1 To Len(parity) Step 1
        Select Case Mid(parity, i, 1)
            Case "A"
                dest = dest & EANsetA(Mid(source, i, 1))
            Case "B"
                dest = dest & EANsetB(Mid(source, i, 1))
        End Select
        If i <> Len(parity) Then
            dest = dest & "11"
        End If
    Next i
    
    EAN_2 = DRAWLINEAR(dest)
    
End Function
