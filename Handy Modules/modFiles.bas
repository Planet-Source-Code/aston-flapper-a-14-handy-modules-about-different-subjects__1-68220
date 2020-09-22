Attribute VB_Name = "modFiles"
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long


Public Function FileAppend(File, text)
    On Error Resume Next
    ff = FreeFile
    Open File For Append As ff
        Print #ff, CStr(text)
    Close ff
End Function

Public Function FileWrite(File, text)
    On Error Resume Next
    ff = FreeFile
    Kill File
    Open File For Binary As ff
        Put #ff, , CStr(text)
    Close ff
End Function

Public Function FileDelete(File)
    On Error Resume Next
    Kill File
End Function

Public Function FileData(File) As String
    Dim dat As String
    ff = FreeFile
    Open File For Binary As ff
        dat = Space$(LOF(ff))
        Get #ff, , dat
    Close ff
    FileData = dat
End Function

Public Function FileCopy(File1, File2)
    FileCopy = CopyFile(CStr(File1), CStr(File2), False)
End Function

Public Function FileMove(File1, File2)
    FileMove = MoveFile(CStr(File1), CStr(File2))
End Function

Public Function FileLen(File)
    ff = FreeFile
    Open File For Binary As ff
        FileLen = LOF(ff)
    Close ff
End Function

Public Function FileAttributes(File)
    FileAttributes = GetFileAttributes(CStr(File))
End Function

