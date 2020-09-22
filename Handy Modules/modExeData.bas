Attribute VB_Name = "modExeData"
Public Property Get edPath() As String
    t = App.path
    If right$(t, 1) = "\" Then
        t = t & App.EXEName & ".exe"
    Else
        t = t & "\" & App.EXEName & ".exe"
    End If
    edPath = t
End Property

Public Function edSetData(Optional fil As String = "", Optional data As String = "")
    Dim ff As Long
    Dim lens As Long
    Dim dat As String
    If fil = "" Then fil = edPath
    ff = FreeFile
    Open fil For Binary As ff
        dat = Space$(LOF(ff))
        Get ff, , dat
    Close ff
    dat = dat & data
    pt1 = nts4T(Len(data))
    dat = dat & pt1
    Kill fil
    ff = FreeFile
    Open fil For Binary As ff
        Put ff, , dat
    Close ff
End Function

Public Function edGetData(Optional fil As String = "") As String
    Dim ff As Long
    Dim lens As Long
    Dim dat As String
    If fil = "" Then fil = edPath
    ff = FreeFile
    Open fil For Binary As ff
        dat = Space$(LOF(ff))
        Get ff, , dat
    Close ff
    If Len(dat) <= 4 Then Exit Function
    lens = nts4F(right$(dat, 4))
    dat = right$(dat, lens + 4)
    dat = left$(dat, lens)
    edGetData = dat
End Function

Public Function edRemoveData(Optional fil As String = "")
    Dim ff As Long
    Dim lens As Long
    Dim dat As String
    If fil = "" Then fil = edPath
    ff = FreeFile
    Open fil For Binary As ff
        dat = Space$(LOF(ff))
        Get ff, , dat
    Close ff
    If Len(dat) <= 4 Then Exit Function
    lens = nts4F(right$(dat, 4))
    If lens = 0 Then Exit Function
    Kill fil
    dat = left$(dat, Len(dat) - lens - 4)
    ff = FreeFile
    Open fil For Binary As ff
        Put ff, , dat
    Close ff
End Function

Private Function nts4T(num As Long) As String
    Dim hlen As String
    hlen = Format(Hex(num), "00000000")
    If Len(hlen) <> 8 Then hlen = String(8 - Len(hlen), "0") & hlen
    nts4T = Chr$(XHexToDecimall(Mid$(hlen, 1, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 3, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 5, 2))) & Chr$(XHexToDecimall(Mid$(hlen, 7, 2)))
End Function

Private Function nts4F(num As String) As Long
    If Len(num) <> 4 Then Exit Function
    num = StrSetLength(Hex(Asc(Mid$(num, 1, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 2, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 3, 1))), 2, "0", 1) & StrSetLength(Hex(Asc(Mid$(num, 4, 1))), 2, "0", 1)
    nts4F = XHexToDecimall(CStr(num))
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


