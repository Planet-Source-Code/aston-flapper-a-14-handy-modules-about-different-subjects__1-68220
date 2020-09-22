Attribute VB_Name = "modStrings"

Public Function StrToArrayStr(ParamArray Array_() As Variant) As String
    Dim str1 As String
    Dim zt As Long
    str1 = nts4T(UBound(Array_) + 1)
    zt = 1
    For a = 0 To UBound(Array_)
        str1 = str1 & nts4T(zt)
        zt = zt + Len(Array_(a))
    Next
    str1 = str1 & nts4T(zt)
    For a = 0 To UBound(Array_)
        str1 = str1 & Array_(a)
    Next
    StrToArrayStr = str1
End Function


Public Function StrFromArrayStr(Arr As String, ArrCnt As Long) As String
    Dim strs As Long
    Dim data As String
    strs = nts4F(Mid$(Arr, 1, 4))
    If ArrCnt > strs Then Exit Function
    data = Mid$(Arr, strs * 4 + 9)
    bgn = nts4F(Mid$(Arr, ArrCnt * 4 + 1, 4))
    ent = nts4F(Mid$(Arr, ArrCnt * 4 + 5, 4))
    ln = ent - bgn
    StrFromArrayStr = Mid$(data, bgn, ln)
End Function

Public Function ArrayStrLen(ArrStr As String) As Long
    ArrayStrLen = nts4F(Mid$(ArrStr, 1, 4))
End Function

Public Function StrByte(txt, a) As Variant
    StrByte = Mid$(txt, a, 1)
End Function

Public Function StrCompare(str1, str2) As Long
    Dim t As Long
    tol = (Len(str1) + Len(str2)) / 2
    If Len(str2) > Len(str1) Then
        t = t + Len(str2) - Len(str1)
        str2 = Mid$(str2, 1, Len(str1))
    End If
    If Len(str1) > Len(str2) Then
        t = t + Len(str1) - Len(str2)
        str1 = Mid$(str1, 1, Len(str2))
    End If
    For a = 1 To Len(str1)
        If Mid$(str1, a, 1) <> Mid$(str2, a, 1) Then
            If Asc(Mid$(str1, a, 1)) > Asc(Mid$(str2, a, 1)) Then
                t = t + Asc(Mid$(str1, a, 1)) - Asc(Mid$(str2, a, 1))
            Else
                t = t + Asc(Mid$(str2, a, 1)) - Asc(Mid$(str1, a, 1))
            End If
        End If
        DoEvents
    Next
    StrCompare = t / tol
End Function

Public Function StrToMix(txt1, key)
    txt2 = key
    If Len(txt2) > Len(txt1) Then txt2 = Mid$(txt2, 1, Len(txt1))
    If Len(txt1) > Len(txt2) Then txt1 = Mid$(txt1, 1, Len(txt2))
    For a = 1 To Len(txt1)
        dat = dat & Chr$(IIf(Asc(Mid$(txt1, a, 1)) + Asc(Mid$(txt2, a, 1)) >= 256, Asc(Mid$(txt1, a, 1)) + Asc(Mid$(txt2, a, 1)) - 255, Asc(Mid$(txt1, a, 1)) + Asc(Mid$(txt2, a, 1))))
        DoEvents
    Next
    StrToMix = dat
End Function

Public Function StrFromMix(txt1, key)
    txt2 = key
    If Len(txt2) > Len(txt1) Then txt2 = Mid$(txt2, 1, Len(txt1))
    If Len(txt1) > Len(txt2) Then txt2 = Mid$(txt1, 1, Len(txt2))
    For a = 1 To Len(txt1)
        dat = dat & Chr$(IIf(Asc(Mid$(txt1, a, 1)) - Asc(Mid$(txt2, a, 1)) < 0, Asc(Mid$(txt1, a, 1)) - Asc(Mid$(txt2, a, 1)) + 255, Asc(Mid$(txt1, a, 1)) - Asc(Mid$(txt2, a, 1))))
        DoEvents
    Next
    StrFromMix = dat
End Function

Public Function StrToURL(StringToEncode As String) As String
    Dim TempAns As String
    Dim CurChr As Integer
    CurChr = 1
    Do Until CurChr - 1 = Len(StringToEncode)
        Select Case Asc(Mid(StringToEncode, CurChr, 1))
        Case 48 To 57, 65 To 90, 97 To 122
            TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
        Case 32
            TempAns = TempAns & "%" & Hex(32)
        Case Else
            TempAns = TempAns & "%" & _
            Format(Hex(Asc(Mid(StringToEncode, _
            CurChr, 1))), "00")
        End Select
        CurChr = CurChr + 1
    Loop
    StrToURL = TempAns
End Function


Public Function StrFromURL(StringToDecode As String) As String
    Dim TempAns As String
    Dim CurChr As Integer
    CurChr = 1
    Do Until CurChr - 1 = Len(StringToDecode)
        Select Case Mid(StringToDecode, CurChr, 1)
        Case "+"
            TempAns = TempAns & " "
        Case "%"
            TempAns = TempAns & Chr(Val("&h" & _
            Mid(StringToDecode, CurChr + 1, 2)))
            CurChr = CurChr + 2
        Case Else
            TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
        End Select
        CurChr = CurChr + 1
    Loop
    StrFromURL = TempAns
End Function



Public Function StrSetLength(txt, length, Optional fill = " ", Optional side = 2) As String
    Dim dat As String
    dat = txt
    Do
        If Len(txt) > length Then
            If side = 1 Then
                dat = StrLeft(dat, length)
            ElseIf side = 2 Then
                dat = StrRight(dat, length)
            End If
        ElseIf Len(txt) < length Then
            If side = 1 Then
                dat = fill & dat
            ElseIf side = 2 Then
                dat = dat & fill
            End If
        End If
        DoEvents
    Loop Until Len(dat) = length
    StrSetLength = dat
End Function

Public Function StrIn(txt, b) As Boolean
    If InStr(txt, b) <> 0 Then StrIn = True
End Function

Public Function StrReplace(txt, ParamArray b() As Variant)
    Dim dat As String
    dat = txt
    For a = 0 To UBound(b) Step 2
        dat = Replace$(dat, CStr(b(a)), CStr(b(a + 1)))
    Next
    StrReplace = dat
End Function

Public Function StrLeft(txt, Optional a = -1, Optional b = -1)
    If a <> -1 Then
        StrLeft = left$(txt, a)
    ElseIf b <> -1 Then
        StrLeft = left$(txt, Len(txt) - Len(b))
    End If
End Function

Public Function StrRight(txt, Optional a = -1, Optional b = -1)
    If a <> -1 Then
        StrRight = right$(txt, a)
    ElseIf b <> -1 Then
        StrRight = right$(txt, Len(txt) - Len(b))
    End If
End Function

Public Function StrTo(txt, b) As String
    For a = 1 To Len(txt)
        If Mid$(txt, a, Len(b)) = b Then Exit For
        c = c & Mid$(txt, a, 1)
    Next
    StrTo = c
End Function

Public Function StrFrom(txt, b) As String
    StrFrom = Mid$(txt, InStr(txt, b) + Len(b))
End Function

Public Function StrBetween(txt, a, b)
    dat = StrFrom(txt, a)
    StrBetween = StrTo(dat, b)
End Function

Public Function StrToHex(txt, Optional st = " ")
    For a = 1 To Len(txt)
        b = b & StrSetLength(Hex(Asc(Mid$(txt, a, 1))), 2, "0", 1) & st
        DoEvents
    Next
    StrToHex = Mid$(b, 1, Len(b) - Len(st))
End Function

Public Function StrFromHex(txt, Optional st = " ")
    g = Split(txt, st)
    For a = 0 To UBound(g)
        b = b & Chr$(XHexToDecimall(g(a)))
    Next
    StrFromHex = b
End Function

Private Function XHexToDecimall(num As Variant)
    For a = 1 To Len(num)
        If Mid$(num, a, 1) <> "0" Then
            Exit For
        Else
            zh = True
        End If
    Next
    If zh = True Then num = Mid$(num, a)
    num = UCase$(num)
    Dim nums(13) As Currency
    nums(1) = 1
    nums(2) = 16
    For a = 3 To 13
        nums(a) = nums(a - 1) * 16
    Next
    For a = Len(num) To 1 Step -1
        g = g + Mid$(num, a, 1)
    Next
    num = g
    For a = 1 To Len(num)
        gh = Mid$(num, a, 1)
        If gh = "0" Then numm = 0
        If gh = "1" Then numm = 1
        If gh = "2" Then numm = 2
        If gh = "3" Then numm = 3
        If gh = "4" Then numm = 4
        If gh = "5" Then numm = 5
        If gh = "6" Then numm = 6
        If gh = "7" Then numm = 7
        If gh = "8" Then numm = 8
        If gh = "9" Then numm = 9
        If gh = "A" Then numm = 10
        If gh = "B" Then numm = 11
        If gh = "C" Then numm = 12
        If gh = "D" Then numm = 13
        If gh = "E" Then numm = 14
        If gh = "F" Then numm = 15
        numm = numm * nums(a)
        gg = gg + numm
    Next
    XHexToDecimall = gg
End Function


Private Function nts4T(num As Long) As String
    Dim hlen As String
    hlen = Format(Hex(num), "00000000")
    If Len(hlen) < 8 Then hlen = String(8 - Len(hlen), "0") & hlen
    nts4T = Chr$(XHexToDecimall(Mid$(hlen, 1, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 3, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 5, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 7, 2)))
End Function

Private Function nts4F(num As String) As Long
    If Len(num) <> 4 Then Exit Function
    num = StrSetLength(Hex(Asc(Mid$(num, 1, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 2, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 3, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 4, 1))), 2, "0", 1)
    nts4F = XHexToDecimall(CStr(num))
End Function




