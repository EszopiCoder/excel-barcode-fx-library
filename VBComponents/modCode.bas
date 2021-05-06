Attribute VB_Name = "modCode"
Option Explicit

'Wikipedia References
'Code 39: https://en.wikipedia.org/wiki/Code_39
'Code 11: https://en.wikipedia.org/wiki/Code_11
'Code 93: https://en.wikipedia.org/wiki/Code_93

Function Code11(source As String) As String
Attribute Code11.VB_Description = "Encode Code 11 barcode."
Attribute Code11.VB_ProcData.VB_Invoke_Func = " \n20"
    Const strCode11 As String = "0123456789-"
    Dim Coode11Table
    Dim i As Integer
    Dim count As Integer
    Dim dest As String
    
    Coode11Table = Array("111121", "211121", "121121", "221111", "112121", _
        "212111", "122111", "111221", "211211", "211111", "112111")
        
    'Validate source
    For i = 1 To Len(source) Step 1
        If InStr(1, strCode11, Mid(source, i, 1)) = 0 Then
            Code11 = "Invalid character found: " & Mid(source, i, 1)
            Exit Function
        End If
    Next i
    
    count = 0
    'Start character (asterisk)
    dest = "112211"
    'Middle characters
    For i = 1 To Len(source) Step 1
        dest = dest & Coode11Table(InStr(1, strCode11, Mid(source, i, 1)) - 1)
        count = count + InStr(1, strCode11, Mid(source, i, 1)) - 1
    Next i
    'End character (asterisk)
    dest = dest & "11221"
    
    Code11 = DRAWLINEAR(dest)
End Function

Function Code39(source As String, Optional CHECK_DIGIT As Boolean = False) As String
Attribute Code39.VB_Description = "Encode Code 39 barcode."
Attribute Code39.VB_ProcData.VB_Invoke_Func = " \n20"
    Const strCode39 As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
    Dim Code39Table
    Dim i As Integer
    Dim count As Integer
    Dim dest As String
    
    Code39Table = Array("1112212111", "2112111121", "1122111121", "2122111111", "1112211121", _
        "2112211111", "1122211111", "1112112121", "2112112111", "1122112111", "2111121121", _
        "1121121121", "2121121111", "1111221121", "2111221111", "1121221111", "1111122121", _
        "2111122111", "1121122111", "1111222111", "2111111221", "1121111221", "2121111211", _
        "1111211221", "2111211211", "1121211211", "1111112221", "2111112211", "1121112211", _
        "1111212211", "2211111121", "1221111121", "2221111111", "1211211121", "2211211111", _
        "1221211111", "1211112121", "2211112111", "1221112111", "1212121111", "1212111211", _
        "1211121211", "1112121211")
    source = UCase(source)
    
    'Validate source
    For i = 1 To Len(source) Step 1
        If InStr(1, strCode39, Mid(source, i, 1)) = 0 Then
            Code39 = "Invalid character found: " & Mid(source, i, 1)
            Exit Function
        End If
    Next i
    
    count = 0
    'Start character (asterisk)
    dest = "1211212111"
    'Middle characters
    For i = 1 To Len(source) Step 1
        dest = dest & Code39Table(InStr(1, strCode39, Mid(source, i, 1)) - 1)
        count = count + InStr(1, strCode39, Mid(source, i, 1)) - 1
    Next i
    'Check digit (Code 39 mod 43)
    If CHECK_DIGIT = True Then
        count = count Mod 43
        dest = dest & Code39Table(count)
    End If
    'End character (asterisk)
    dest = dest & "1211212111"
    
    Code39 = DRAWLINEAR(dest)
End Function

Function Code93(source As String) As String
Attribute Code93.VB_Description = "Encode Code 93 barcode."
Attribute Code93.VB_ProcData.VB_Invoke_Func = " \n20"
    Const strCode93 As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
    Dim Code93Table
    Dim i As Integer
    Dim c, K As Integer
    Dim weight As Integer
    Dim dest As String
    
    Code93Table = Array("131112", "111213", "111312", "111411", "121113", "121212", "121311", _
        "111114", "131211", "141111", "211113", "211212", "211311", "221112", "221211", "231111", _
        "112113", "112212", "112311", "122112", "132111", "111123", "111222", "111321", "121122", _
        "131121", "212112", "212211", "211122", "211221", "221121", "222111", "112122", "112221", _
        "122121", "123111", "121131", "311112", "311211", "321111", "112131", "113121", "211131", _
        "121221", "312111", "311121", "122211")
    source = UCase(source)
    
    'Validate source
    For i = 1 To Len(source) Step 1
        If InStr(1, strCode93, Mid(source, i, 1)) = 0 Then
            Code93 = "Invalid character found: " & Mid(source, i, 1)
            Exit Function
        End If
    Next i
    
    'Start character
    dest = "111141"
    'Middle character
    For i = 1 To Len(source) Step 1
        dest = dest & Code93Table(InStr(1, strCode93, Mid(source, i, 1)) - 1)
    Next i
    'Check digit C
    c = 0
    weight = 1
    For i = Len(source) To 1 Step -1
        c = c + (InStr(1, strCode93, Mid(source, i, 1)) - 1) * weight
        weight = weight + 1
        If weight = 21 Then weight = 1
    Next i
    c = c Mod 47
    source = source & Mid(strCode93, c + 1, 1)
    dest = dest & Code93Table(InStr(1, strCode93, Mid(strCode93, c + 1, 1)) - 1)
    'Check digit K
    K = 0
    weight = 1
    For i = Len(source) To 1 Step -1
        K = K + (InStr(1, strCode93, Mid(source, i, 1)) - 1) * weight
        weight = weight + 1
        If weight = 16 Then weight = 1
    Next i
    K = K Mod 47
    source = source & Mid(strCode93, K + 1, 1)
    dest = dest & Code93Table(InStr(1, strCode93, Mid(strCode93, K + 1, 1)) - 1)
    'End character
    dest = dest & "1111411"
    
    Code93 = DRAWLINEAR(dest)
End Function
