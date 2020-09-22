Attribute VB_Name = "modEncryption"
Option Explicit
Option Base 0

Private m_lOnBits(30)           As Long
Private m_l2Power(30)           As Long
Private lngTrack                As Long
Private arrLongConversion(4)    As Long
Private arrSplit64(63)          As Byte
Private aDecTab(255)            As Integer
Private s(0 To 255)             As Integer
Private kep(0 To 255)           As Integer
Private i As Integer, j         As Integer
Private path                    As String

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21
Private Const sEncTab As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4     As Long
    Dim lY4     As Long
    Dim lX8     As Long
    Dim lY8     As Long
    Dim lResult As Long
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    AddUnsigned = lResult
End Function

Private Function LRot(ByVal x As Long, ByVal n As Long) As Long
    LRot = LShift(x, n) Or RShift(x, (32 - n))
End Function

Private Function ConvertToWordArray(sMessage As String) As Long()
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long
    Const MODULUS_BITS      As Long = 512
    Const CONGRUENT_BITS    As Long = 448
    lMessageLength = Len(sMessage)
    lNumberOfWords = (((lMessageLength + _
        ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ _
        (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * _
        (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
        
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or _
    LShift(&H80, lBytePosition)
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    ConvertToWordArray = lWordArray
End Function

Private Function EncodeQuantum(B() As Byte) As String
    Dim sOutput As String
    Dim c As Integer
    
    sOutput = ""
    c = SHR2(B(0)) And &H3F
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    c = SHL4(B(0) And &H3) Or (SHR4(B(1)) And &HF)
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    c = SHL2(B(1) And &HF) Or (SHR6(B(2)) And &H3)
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    c = B(2) And &H3F
    sOutput = sOutput & Mid(sEncTab, c + 1, 1)
    
    EncodeQuantum = sOutput
    
End Function

Private Function DecodeQuantum(d() As Byte) As String
    Dim sOutput As String
    Dim c As Long
    
    sOutput = ""
    c = SHL2(d(0)) Or (SHR4(d(1)) And &H3)
    sOutput = sOutput & Chr$(c)
    c = SHL4(d(1) And &HF) Or (SHR2(d(2)) And &HF)
    sOutput = sOutput & Chr$(c)
    c = SHL6(d(2) And &H3) Or d(3)
    sOutput = sOutput & Chr$(c)
    
    DecodeQuantum = sOutput
    
End Function

Private Function MakeDecTab()
' Set up Radix 64 decoding table
    Dim t As Integer
    Dim c As Integer

    For c = 0 To 255
        aDecTab(c) = -1
    Next
  
    t = 0
    For c = Asc("A") To Asc("Z")
        aDecTab(c) = t
        t = t + 1
    Next
  
    For c = Asc("a") To Asc("z")
        aDecTab(c) = t
        t = t + 1
    Next
    
    For c = Asc("0") To Asc("9")
        aDecTab(c) = t
        t = t + 1
    Next
    
    c = Asc("+")
    aDecTab(c) = t
    t = t + 1
    
    c = Asc("/")
    aDecTab(c) = t
    t = t + 1
    
    c = Asc("=")    ' flag for the byte-deleting char
    aDecTab(c) = t  ' should be 64

End Function

' Version 3: ShiftLeft and ShiftRight functions improved.
Private Function SHL2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 2 bits
' i.e. VB equivalent of "bytValue << 2" in C
    SHL2 = (bytValue * &H4) And &HFF
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 4 bits
' i.e. VB equivalent of "bytValue << 4" in C
    SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 6 bits
' i.e. VB equivalent of "bytValue << 6" in C
    SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 2 bits
' i.e. VB equivalent of "bytValue >> 2" in C
    SHR2 = bytValue \ &H4
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 4 bits
' i.e. VB equivalent of "bytValue >> 4" in C
    SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 6 bits
' i.e. VB equivalent of "bytValue >> 6" in C
    SHR6 = bytValue \ &H40
End Function



Private Function MD5Round(strRound As String, A As Long, B As Long, c As Long, d As Long, x As Long, s As Long, ac As Long) As Long
    Select Case strRound
        Case Is = "FF"
            A = MD5LongAdd4(A, (B And c) Or (Not (B) And d), x, ac)
            A = MD5Rotate(A, s)
            A = MD5LongAdd(A, B)
        Case Is = "GG"
            A = MD5LongAdd4(A, (B And d) Or (c And Not (d)), x, ac)
            A = MD5Rotate(A, s)
            A = MD5LongAdd(A, B)
        Case Is = "HH"
            A = MD5LongAdd4(A, B Xor c Xor d, x, ac)
            A = MD5Rotate(A, s)
            A = MD5LongAdd(A, B)
        Case Is = "II"
            A = MD5LongAdd4(A, c Xor (B Or Not (d)), x, ac)
            A = MD5Rotate(A, s)
            A = MD5LongAdd(A, B)
    End Select
End Function

Private Function MD5Rotate(lngValue As Long, lngBits As Long) As Long
Dim lngSign As Long
Dim lngI As Long
    lngBits = (lngBits Mod 32)
    
    If lngBits = 0 Then MD5Rotate = lngValue: Exit Function
    
    For lngI = 1 To lngBits
        lngSign = lngValue And &HC0000000
        lngValue = (lngValue And &H3FFFFFFF) * 2
        lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next
    
    MD5Rotate = lngValue
End Function

Private Function TRID() As String

    Dim sngNum As Single, lngnum As Long
    Dim strResult As String
   
    sngNum = Rnd(2147483648#)
    strResult = CStr(sngNum)
    
    strResult = Replace(strResult, "0.", "")
    strResult = Replace(strResult, ".", "")
    strResult = Replace(strResult, "E-", "")
    
    TRID = strResult

End Function

Private Function MD564Split(lngLength As Long, bytBuffer() As Byte) As String

    Dim lngBytesTotal As Long, lngBytesToAdd As Long
    Dim intLoop As Integer, intLoop2 As Integer, lngTrace As Long
    Dim intInnerLoop As Integer, intLoop3 As Integer
    
    lngBytesTotal = lngTrack Mod 64
    lngBytesToAdd = 64 - lngBytesTotal
    lngTrack = (lngTrack + lngLength)
    
    If lngLength >= lngBytesToAdd Then
        For intLoop = 0 To lngBytesToAdd - 1
            arrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
        Next intLoop
        
        MD5Conversion arrSplit64
        
        lngTrace = (lngLength) Mod 64

        For intLoop2 = lngBytesToAdd To lngLength - intLoop - lngTrace Step 64
            For intInnerLoop = 0 To 63
                arrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
            Next intInnerLoop
            
            MD5Conversion arrSplit64
        
        Next intLoop2
        
        lngBytesTotal = 0
    Else
    
      intLoop2 = 0
    
    End If
    
    For intLoop3 = 0 To lngLength - intLoop2 - 1
        
        arrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)
    
    Next intLoop3
     
End Function

Private Function MD5StringArray(strInput As String) As Byte()
    
    Dim intLoop As Integer
    Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(strInput))
    
    For intLoop = 0 To Len(strInput) - 1
        bytBuffer(intLoop) = Asc(Mid(strInput, intLoop + 1, 1))
    Next intLoop
    
    MD5StringArray = bytBuffer
    
End Function

Private Sub MD5Conversion(bytBuffer() As Byte)

    Dim x(16) As Long, A As Long
    Dim B As Long, c As Long
    Dim d As Long
    
    A = arrLongConversion(1)
    B = arrLongConversion(2)
    c = arrLongConversion(3)
    d = arrLongConversion(4)
    
    MD5Decode 64, x, bytBuffer
    
    MD5Round "FF", A, B, c, d, x(0), S11, -680876936
    MD5Round "FF", d, A, B, c, x(1), S12, -389564586
    MD5Round "FF", c, d, A, B, x(2), S13, 606105819
    MD5Round "FF", B, c, d, A, x(3), S14, -1044525330
    MD5Round "FF", A, B, c, d, x(4), S11, -176418897
    MD5Round "FF", d, A, B, c, x(5), S12, 1200080426
    MD5Round "FF", c, d, A, B, x(6), S13, -1473231341
    MD5Round "FF", B, c, d, A, x(7), S14, -45705983
    MD5Round "FF", A, B, c, d, x(8), S11, 1770035416
    MD5Round "FF", d, A, B, c, x(9), S12, -1958414417
    MD5Round "FF", c, d, A, B, x(10), S13, -42063
    MD5Round "FF", B, c, d, A, x(11), S14, -1990404162
    MD5Round "FF", A, B, c, d, x(12), S11, 1804603682
    MD5Round "FF", d, A, B, c, x(13), S12, -40341101
    MD5Round "FF", c, d, A, B, x(14), S13, -1502002290
    MD5Round "FF", B, c, d, A, x(15), S14, 1236535329

    MD5Round "GG", A, B, c, d, x(1), S21, -165796510
    MD5Round "GG", d, A, B, c, x(6), S22, -1069501632
    MD5Round "GG", c, d, A, B, x(11), S23, 643717713
    MD5Round "GG", B, c, d, A, x(0), S24, -373897302
    MD5Round "GG", A, B, c, d, x(5), S21, -701558691
    MD5Round "GG", d, A, B, c, x(10), S22, 38016083
    MD5Round "GG", c, d, A, B, x(15), S23, -660478335
    MD5Round "GG", B, c, d, A, x(4), S24, -405537848
    MD5Round "GG", A, B, c, d, x(9), S21, 568446438
    MD5Round "GG", d, A, B, c, x(14), S22, -1019803690
    MD5Round "GG", c, d, A, B, x(3), S23, -187363961
    MD5Round "GG", B, c, d, A, x(8), S24, 1163531501
    MD5Round "GG", A, B, c, d, x(13), S21, -1444681467
    MD5Round "GG", d, A, B, c, x(2), S22, -51403784
    MD5Round "GG", c, d, A, B, x(7), S23, 1735328473
    MD5Round "GG", B, c, d, A, x(12), S24, -1926607734
  
    MD5Round "HH", A, B, c, d, x(5), S31, -378558
    MD5Round "HH", d, A, B, c, x(8), S32, -2022574463
    MD5Round "HH", c, d, A, B, x(11), S33, 1839030562
    MD5Round "HH", B, c, d, A, x(14), S34, -35309556
    MD5Round "HH", A, B, c, d, x(1), S31, -1530992060
    MD5Round "HH", d, A, B, c, x(4), S32, 1272893353
    MD5Round "HH", c, d, A, B, x(7), S33, -155497632
    MD5Round "HH", B, c, d, A, x(10), S34, -1094730640
    MD5Round "HH", A, B, c, d, x(13), S31, 681279174
    MD5Round "HH", d, A, B, c, x(0), S32, -358537222
    MD5Round "HH", c, d, A, B, x(3), S33, -722521979
    MD5Round "HH", B, c, d, A, x(6), S34, 76029189
    MD5Round "HH", A, B, c, d, x(9), S31, -640364487
    MD5Round "HH", d, A, B, c, x(12), S32, -421815835
    MD5Round "HH", c, d, A, B, x(15), S33, 530742520
    MD5Round "HH", B, c, d, A, x(2), S34, -995338651
 
    MD5Round "II", A, B, c, d, x(0), S41, -198630844
    MD5Round "II", d, A, B, c, x(7), S42, 1126891415
    MD5Round "II", c, d, A, B, x(14), S43, -1416354905
    MD5Round "II", B, c, d, A, x(5), S44, -57434055
    MD5Round "II", A, B, c, d, x(12), S41, 1700485571
    MD5Round "II", d, A, B, c, x(3), S42, -1894986606
    MD5Round "II", c, d, A, B, x(10), S43, -1051523
    MD5Round "II", B, c, d, A, x(1), S44, -2054922799
    MD5Round "II", A, B, c, d, x(8), S41, 1873313359
    MD5Round "II", d, A, B, c, x(15), S42, -30611744
    MD5Round "II", c, d, A, B, x(6), S43, -1560198380
    MD5Round "II", B, c, d, A, x(13), S44, 1309151649
    MD5Round "II", A, B, c, d, x(4), S41, -145523070
    MD5Round "II", d, A, B, c, x(11), S42, -1120210379
    MD5Round "II", c, d, A, B, x(2), S43, 718787259
    MD5Round "II", B, c, d, A, x(9), S44, -343485551
    
    arrLongConversion(1) = MD5LongAdd(arrLongConversion(1), A)
    arrLongConversion(2) = MD5LongAdd(arrLongConversion(2), B)
    arrLongConversion(3) = MD5LongAdd(arrLongConversion(3), c)
    arrLongConversion(4) = MD5LongAdd(arrLongConversion(4), d)
    
End Sub

Private Function MD5LongAdd(lngVal1 As Long, lngVal2 As Long) As Long
    
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    
    MD5LongAdd = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Function MD5LongAdd4(lngVal1 As Long, lngVal2 As Long, lngVal3 As Long, lngVal4 As Long) As Long
    
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    MD5LongAdd4 = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Sub MD5Decode(intLength As Integer, lngOutBuffer() As Long, bytInBuffer() As Byte)
    
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double
    
    intDblIndex = 0
    
    For intByteIndex = 0 To intLength - 1 Step 4
        
        dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256# + bytInBuffer(intByteIndex + 2) * 65536# + bytInBuffer(intByteIndex + 3) * 16777216#
        lngOutBuffer(intDblIndex) = MD5LongConversion(dblSum)
        intDblIndex = (intDblIndex + 1)
    
    Next intByteIndex

End Sub

Private Function MD5LongConversion(dblValue As Double) As Long
    
    If dblValue < 0 Or dblValue >= OFFSET_4 Then Error 6
        
    If dblValue <= MAXINT_4 Then
        MD5LongConversion = dblValue
    Else
        MD5LongConversion = dblValue - OFFSET_4
    End If
        
End Function

Private Sub MD5Finish()
Dim dblBits As Double
Dim arrPadding(72) As Byte
Dim lngBytesBuffered As Long
    
    arrPadding(0) = &H80
    dblBits = lngTrack * 8
    
    lngBytesBuffered = lngTrack Mod 64
    
    If lngBytesBuffered <= 56 Then
        MD564Split (56 - lngBytesBuffered), arrPadding
    Else
        MD564Split (120 - lngTrack), arrPadding
    End If
    
    
    arrPadding(0) = MD5LongConversion(dblBits) And &HFF&
    arrPadding(1) = MD5LongConversion(dblBits) \ 256 And &HFF&
    arrPadding(2) = MD5LongConversion(dblBits) \ 65536 And &HFF&
    arrPadding(3) = MD5LongConversion(dblBits) \ 16777216 And &HFF&
    arrPadding(4) = 0
    arrPadding(5) = 0
    arrPadding(6) = 0
    arrPadding(7) = 0
    
    MD564Split 8, arrPadding
End Sub

Private Function MD5StringChange(lngnum As Long) As String
Dim bytA As Byte
Dim bytB As Byte
Dim bytC As Byte
Dim bytD As Byte
     bytA = lngnum And &HFF&
     If bytA < 16 Then
         MD5StringChange = "0" & Hex(bytA)
     Else
         MD5StringChange = Hex(bytA)
     End If
            
     bytB = (lngnum And &HFF00&) \ 256
     If bytB < 16 Then
         MD5StringChange = MD5StringChange & "0" & Hex(bytB)
     Else
         MD5StringChange = MD5StringChange & Hex(bytB)
     End If
     
     bytC = (lngnum And &HFF0000) \ 65536
     If bytC < 16 Then
         MD5StringChange = MD5StringChange & "0" & Hex(bytC)
     Else
         MD5StringChange = MD5StringChange & Hex(bytC)
     End If
    
     If lngnum < 0 Then
         bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
     Else
         bytD = (lngnum And &HFF000000) \ 16777216
     End If
     
     If bytD < 16 Then
         MD5StringChange = MD5StringChange & "0" & Hex(bytD)
     Else
         MD5StringChange = MD5StringChange & Hex(bytD)
     End If
End Function

Private Function MD5Value() As String
    MD5Value = LCase(MD5StringChange(arrLongConversion(1)) & MD5StringChange(arrLongConversion(2)) & MD5StringChange(arrLongConversion(3)) & MD5StringChange(arrLongConversion(4)))
End Function

Private Sub MD5Start()
    lngTrack = 0
    arrLongConversion(1) = MD5LongConversion(1732584193#)
    arrLongConversion(2) = MD5LongConversion(4023233417#)
    arrLongConversion(3) = MD5LongConversion(2562383102#)
    arrLongConversion(4) = MD5LongConversion(271733878#)
End Sub
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################
'#############################################################################


Public Function MD5(strMessage As String) As String
    Dim bytBuffer() As Byte
    bytBuffer = MD5StringArray(strMessage)
    MD5Start
    MD564Split Len(strMessage), bytBuffer
    MD5Finish
    MD5 = MD5Value
End Function

Public Function SHA1(sMessage As String) As String
    Dim HASH(4)         As Long
    Dim M()             As Long
    Dim W(79)           As Long
    Dim A, B, c, d, e   As Long
    Dim G, h, i, j      As Long
    Dim T1, T2          As Long
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
    ' Initial hash container values
    HASH(0) = &H67452301
    HASH(1) = &HEFCDAB89
    HASH(2) = &H98BADCFE
    HASH(3) = &H10325476
    HASH(4) = &HC3D2E1F0
    
    ' Preprocessing. Append padding bits and length and convert to words.
    M = ConvertToWordArray(sMessage)
    
    ' We process sixteen 32-bit words at a time, which is a 512-bit data block
    For i = 0 To UBound(M) Step 16
        ' Set inital values for the hash operators,
        ' this includes previous hash values.
        A = HASH(0)
        B = HASH(1)
        c = HASH(2)
        d = HASH(3)
        e = HASH(4)
        
        ' We grab the sixteen 32-bit words from our Message word array,
        ' this is the data we'll be working on.
        For G = 0 To 15
            W(G) = M(i + G)
        Next G
        
        ' These sixteen 32-bit words must now be extended through the
        ' initial hashing phase to eighty 32-bit words.
        For G = 16 To 79
            W(G) = LRot(W(G - 3) Xor W(G - 8) Xor W(G - 14) Xor W(G - 16), 1)
        Next G
        
        ' We now begin processing these eighty 32-bit words.
        For j = 0 To 79
            
            ' The processing below is as per SHA1's specification.
            If j <= 19 Then
                T1 = (B And c) Or ((Not B) And d)
                T2 = &H5A827999
            ElseIf j <= 39 Then
                T1 = B Xor c Xor d
                T2 = &H6ED9EBA1
            ElseIf j <= 59 Then
                T1 = (B And c) Or (B And d) Or (c And d)
                T2 = &H8F1BBCDC
            ElseIf j <= 79 Then
                T1 = B Xor c Xor d
                T2 = &HCA62C1D6
            End If
            
            ' For each word we process we run the below hashing function and
            ' set it equal to a, shifting the previous a's value down to b,
            ' so a becomes b, b becomes c, after a 30 Left Rotate of 30,
            ' c becomes d, d becomes e.
            h = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(LRot(A, 5), T1), e), T2), W(j))
            e = d
            d = c
            c = LRot(B, 30)
            B = A
            A = h
        Next j
            
        ' We now add the newley hashed values to the hash container.
        HASH(0) = AddUnsigned(A, HASH(0))
        HASH(1) = AddUnsigned(B, HASH(1))
        HASH(2) = AddUnsigned(c, HASH(2))
        HASH(3) = AddUnsigned(d, HASH(3))
        HASH(4) = AddUnsigned(e, HASH(4))
        
    Next i
    
    ' Output the 160-bit digest
    SHA1 = LCase(right("00000000" & Hex(HASH(0)), 8) & _
        right("00000000" & Hex(HASH(1)), 8) & _
        right("00000000" & Hex(HASH(2)), 8) & _
        right("00000000" & Hex(HASH(3)), 8) & _
        right("00000000" & Hex(HASH(4)), 8))
End Function


Public Function toBase64(sInput As String) As String
    Dim sOutput As String, sLast As String
    Dim B(2) As Byte
    Dim j As Integer
    Dim i As Long, nLen As Long, nQuants As Long
    Dim iIndex As Long
    
    nLen = Len(sInput)
    nQuants = nLen \ 3
    sOutput = String(nQuants * 4, " ")
    iIndex = 0
    ' Now start reading in 3 bytes at a time
    For i = 0 To nQuants - 1
        For j = 0 To 2
           B(j) = Asc(Mid(sInput, (i * 3) + j + 1, 1))
        Next
        Mid$(sOutput, iIndex + 1, 4) = EncodeQuantum(B)
        iIndex = iIndex + 4
    Next
    
    ' Cope with odd bytes
    Select Case nLen Mod 3
    Case 0
        sLast = ""
    Case 1
        B(0) = Asc(Mid(sInput, nLen, 1))
        B(1) = 0
        B(2) = 0
        sLast = EncodeQuantum(B)
        ' Replace last 2 with =
        sLast = left(sLast, 2) & "=="
    Case 2
        B(0) = Asc(Mid(sInput, nLen - 1, 1))
        B(1) = Asc(Mid(sInput, nLen, 1))
        B(2) = 0
        sLast = EncodeQuantum(B)
        ' Replace last with =
        sLast = left(sLast, 3) & "="
    End Select
    
    toBase64 = sOutput & sLast
End Function

Public Function fromBase64(sEncoded As String) As String
    Dim sDecoded As String
    Dim d(3) As Byte
    Dim c As Byte
    Dim di As Integer
    Dim i As Long
    Dim nLen As Long
    Dim iIndex As Long
    
    nLen = Len(sEncoded)
    sDecoded = String((nLen \ 4) * 3, " ")
    iIndex = 0
    di = 0
    Call MakeDecTab
    ' Read in each char in trun
    For i = 1 To Len(sEncoded)
        c = CByte(Asc(Mid(sEncoded, i, 1)))
        c = aDecTab(c)
        If c >= 0 Then
            d(di) = c
            di = di + 1
            If di = 4 Then
                Mid$(sDecoded, iIndex + 1, 3) = DecodeQuantum(d)
                iIndex = iIndex + 3
                If d(3) = 64 Then
                    sDecoded = left(sDecoded, Len(sDecoded) - 1)
                    iIndex = iIndex - 1
                End If
                If d(2) = 64 Then
                    sDecoded = left(sDecoded, Len(sDecoded) - 1)
                    iIndex = iIndex - 1
                End If
                di = 0
            End If
        End If
    Next i
    
    fromBase64 = sDecoded
End Function
