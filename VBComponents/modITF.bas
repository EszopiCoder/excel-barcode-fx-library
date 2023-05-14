Attribute VB_Name = "modITF"
Option Explicit

'GS1 ITF-14 barcodes: https://www.gs1standards.info/itf-14-barcodes/

'Wikipedia References
'Interleaved 2 of 5 (ITF): https://en.wikipedia.org/wiki/Interleaved_2_of_5

Function ITF(source As String) As String
    Dim ITFTable
    Dim i As Integer
    Dim j As Integer
    Dim dest As String
    
    ITFTable = Array("11221", "21112", "12112", "22111", "11212", _
        "21211", "12211", "11122", "21121", "12121")
    
    'ITF requires an even number of digits. If odd, add a zero to the beginning
    If Len(source) Mod 2 <> 0 Then
        source = "0" & source
    End If
    
    'Start character
    dest = "1111"
    'Middle characters
    For i = 1 To Len(source) Step 2
        'Interleave 2 digits at a time (1st digit is bars, 2nd digit is spaces)
        For j = 1 To 5 Step 1
            dest = dest & Mid(ITFTable(Mid(source, i, 1)), j, 1) & Mid(ITFTable(Mid(source, i + 1, 1)), j, 1)
        Next j
    Next i
    'End character
    dest = dest & "211"
    
    ITF = DRAWLINEAR(dest)
End Function

Function ITF_14(source As String) As String
    Dim ITFTable
    Dim i As Integer
    Dim j As Integer
    Dim dest As String
    
    ITFTable = Array("11221", "21112", "12112", "22111", "11212", _
        "21211", "12211", "11122", "21121", "12121")

    'Validate source
    If Len(source) <> 14 And Len(source) <> 13 Then
        ITF_14 = "Improper ITF-14 barcode length (13-14 digits)"
        Exit Function
    ElseIf Len(source) = 14 And GS1_CHECK(Left(source, 13)) <> Right(source, 1) Then
        ITF_14 = "Invalid check digit (" & GS1_CHECK(Left(source, 13)) & ")"
        Exit Function
    End If
    
    'Calculate check digit
    If Len(source) = 13 Then source = source & GS1_CHECK(Left(source, 13))
    
    'Start character
    dest = "1111"
    'Middle characters
    For i = 1 To Len(source) Step 2
        'Interleave 2 digits at a time (1st digit is bars, 2nd digit is spaces)
        For j = 1 To 5 Step 1
            dest = dest & Mid(ITFTable(Mid(source, i, 1)), j, 1) & Mid(ITFTable(Mid(source, i + 1, 1)), j, 1)
        Next j
    Next i
    'End character
    dest = dest & "211"
    
    ITF_14 = DRAWLINEAR(dest)
End Function
