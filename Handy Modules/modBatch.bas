Attribute VB_Name = "modBATCH"
Const PrefPath = "C:\"
Dim Batchpath As String
Dim Batchfile As String
Dim Batchstr As String

Public Property Get batString() As String
    batString = Batchstr
End Property

Public Property Let batString(newval As String)
    Batchstr = newval
End Property

Public Property Get batPath() As String
    batPath = Batchpath
End Property

Public Property Let batPath(cnewval As String)
    Batchpath = cnewval
End Property

Public Sub batPrint(pval As String)
    Batchstr = Batchstr & pval & vbCrLf
End Sub

Public Sub batClear()
    Batchstr = ""
End Sub

Public Sub batCls()
    Batchstr = ""
End Sub

Public Sub batKill()
    On Error Resume Next
    Kill Batchpath & Batchfile & ".bat"
    Kill Batchpath & Batchfile & ".ret"
End Sub

Public Function batRun(visible As Boolean, waitforretval As Boolean, autokill As Boolean) As String
    On Error Resume Next
    Dim ff As Long
    Batchfile = "batchfile" & XRndNum(0, 5000000)
    If Batchpath = "" Then Batchpath = PrefPath
    If waitforretval Then Batchstr = Batchstr & vbCrLf & "cd " & Batchpath & vbCrLf & "ECHO ;*-HeRE ENDS0983572-*;>>" & Batchfile & ".ret"
    Batchstr = Replace$(Batchstr, "%FILE%", Batchfile & ".ret")
    Batchstr = "@ECHO OFF" & vbCrLf & "cd " & Batchpath & vbCrLf & "@ECHO ON" & vbCrLf & Batchstr
    ff = FreeFile
    Kill Batchpath & Batchfile & ".bat"
    Open Batchpath & Batchfile & ".bat" For Binary As ff
    Put ff, , Batchstr
    Close ff
    If Not visible Then
        Shell Batchpath & Batchfile & ".bat", vbHide
    Else
        Shell Batchpath & Batchfile & ".bat", vbNormalFocus
    End If
    If waitforretval Then
        Dim dat As String
        Do
            DoEvents
            ff = FreeFile
            Open Batchpath & Batchfile & ".ret" For Binary As ff
                dat = Space$(LOF(ff))
                Get ff, , dat
            Close ff
            pt = Timer
            Do
                DoEvents
            Loop Until Timer >= pt + 0.5
        Loop Until Right$(dat, Len(";*-HeRE ENDS0983572-*;") + 2) = ";*-HeRE ENDS0983572-*;" & vbCrLf
        If autokill Then Kill Batchpath & Batchfile & ".bat"
        If autokill Then Kill Batchpath & Batchfile & ".ret"
        batRun = Mid$(dat, 1, Len(dat) - Len(";*-HeRE ENDS0983572-*;") - 2)
    End If
End Function


Private Function XRndNum(Lowest, Highest)
Lowest = Lowest - 1
Highest = Highest + 1
Randomize Timer
XRndNum = Int(Rnd * (Highest - Lowest - 1)) + Lowest + 1
End Function

