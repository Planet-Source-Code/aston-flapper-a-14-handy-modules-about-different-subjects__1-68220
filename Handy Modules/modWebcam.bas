Attribute VB_Name = "modWebcam"
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054
Private Const WM_CAP_SET_VIDEOFORMAT = &H400 + 45


Public Sub wbcStart(hWnd As Long, Optional wX = 0, Optional wY = 0, Optional wWidth = 100, Optional wHeight = 100)
    mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, wX, wY, wWidth, wHeight, hWnd, 0)
    DoEvents
    SendMessage mCapHwnd, CONNECT, 0, 0
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
    DoEvents
End Sub

Public Sub wbcStop()
    DoEvents
    SendMessage mCapHwnd, DISCONNECT, 0, 0
    SendMessage mCapHwnd, DISCONNECT, 0, 0
    SendMessage mCapHwnd, DISCONNECT, 0, 0
    DoEvents
End Sub

Public Sub wbcTopPicture(Picbox As PictureBox)
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    Picbox.Picture = Clipboard.GetData
End Sub
