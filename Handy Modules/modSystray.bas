Attribute VB_Name = "modSystray"
Option Explicit
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Type NOTIFYICONDATA
  cbSize As Long
  Hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4
Private Const STI_CALLBACKEVENT = &H201
Private Const REG_SZ = 1
Private Const LOCALMACHINE = &H80000002

Public Sub TrayAdd(parentForm As Form, Tip As String)
    On Error Resume Next
    Dim notIcon As NOTIFYICONDATA
    
    With notIcon
        .cbSize = Len(notIcon)
        .Hwnd = parentForm.Hwnd
        .uID = vbNull
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = STI_CALLBACKEVENT
        .hIcon = parentForm.Icon
        .szTip = Tip & vbNullChar
    End With
    
    Shell_NotifyIconA NIM_ADD, notIcon
End Sub

Public Sub TrayModify(parentForm As Form, Tip As String)
    On Error Resume Next
    Dim notIcon As NOTIFYICONDATA
    
    With notIcon
        .cbSize = Len(notIcon)
        .Hwnd = parentForm.Hwnd
        .uID = vbNull
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = STI_CALLBACKEVENT
        .hIcon = parentForm.Icon
        .szTip = Tip & vbNullChar
    End With
    
    Shell_NotifyIconA NIM_MODIFY, notIcon
End Sub

Public Sub TrayDelete(parentForm As Form)
    On Error Resume Next
    Dim notIcon As NOTIFYICONDATA
    
    With notIcon
      .cbSize = Len(notIcon)
      .Hwnd = parentForm.Hwnd
      .uID = vbNull
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .uCallbackMessage = vbNull
      .hIcon = vbNull
      .szTip = "" & vbNullChar
    End With
    
    Shell_NotifyIconA NIM_DELETE, notIcon
End Sub
