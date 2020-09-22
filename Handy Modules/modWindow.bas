Attribute VB_Name = "modWindow"
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public Function WndHide(wnd)
    ShowWindow CLng(wnd), 0
End Function

Public Function WndShow(wnd)
    ShowWindow CLng(wnd), 5
End Function

Public Function WndFlash(wnd)
    FlashWindow CLng(wnd), True
End Function
