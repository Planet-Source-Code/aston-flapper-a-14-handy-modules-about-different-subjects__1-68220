Attribute VB_Name = "modInStart"
Option Explicit
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Sub PutInStart(Var, Value)
    RegCreateKey &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", 1
    RegSetValueEx 1, Var, 0, 1, Value, Len(Value)
End Sub

