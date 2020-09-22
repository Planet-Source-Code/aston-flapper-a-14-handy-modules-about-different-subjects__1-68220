Attribute VB_Name = "modNumtoStr"
Public Function nts4T(num As Long) As String
    Dim hlen As String
    hlen = Hex(num)
    If Len(hlen) < 8 Then hlen = String(8 - Len(hlen), "0") & hlen
    nts4T = Chr$(XHexToDecimall(Mid$(hlen, 1, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 3, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 5, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 7, 2)))
End Function

Public Function nts4F(num As String) As Long
    If Len(num) <> 4 Then Exit Function
    num = StrSetLength(Hex(Asc(Mid$(num, 1, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 2, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 3, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 4, 1))), 2, "0", 1)
    nts4F = XHexToDecimall(CStr(num))
End Function

Private Function StrSetLength(txt, length, Optional fill = " ", Optional side = 2) As String
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

Private Function XHexToDecimall(num As String) As Long
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


