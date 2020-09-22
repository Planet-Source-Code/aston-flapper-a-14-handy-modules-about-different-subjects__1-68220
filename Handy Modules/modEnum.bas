Attribute VB_Name = "modEnum"
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Function enumProcesses(Seperator As String)
    Dim snap As Long
    Dim proc As PROCESSENTRY32
    Dim p As Long
    Dim str1 As String
    snap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    proc.dwSize = Len(proc)
    p = Process32First(snap, proc)
    p = Process32Next(snap, proc)
    Do While p
        str1 = str1 & xto(proc.szExeFile, Chr$(0)) & Seperator
        p = Process32Next(snap, proc)
    Loop
    enumProcesses = Mid$(str1, 1, Len(str1) - 1)
End Function

Private Function xto(txt As String, B As String) As String
    For a = 1 To Len(txt)
        If Mid$(txt, a, Len(B)) = B Then Exit For
        c = c & Mid$(txt, a, 1)
    Next
    xto = c
End Function

