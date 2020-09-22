Attribute VB_Name = "modMouse"
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Property Get MouseX() As Long
    Dim t As POINTAPI
    GetCursorPos t
    MouseX = t.X
End Property

Public Property Get MouseY() As Long
    Dim t As POINTAPI
    GetCursorPos t
    MouseY = t.Y
End Property

Public Property Let MouseX(newval As Long)
    SetCursorPos newval, MouseY
End Property

Public Property Let MouseY(newval As Long)
    SetCursorPos MouseX, newval
End Property

Public Property Get MouseDoubleClickTime() As Long
    MouseDoubleClickTime = GetDoubleClickTime
End Property

Public Property Let MouseDoubleClickTime(newval As Long)
    SetDoubleClickTime newval
End Property

Public Sub MouseClip(mX, mY, mWidth, mHeight)
    Dim rct As RECT
    rct.left = mX
    rct.top = mY
    rct.right = mWidth + mX
    rct.bottom = mHeight + mY
    ClipCursor rct
End Sub
Public Sub MouseSwapButtons(swap As Boolean)
    SwapMouseButton CLng(swap)
End Sub

