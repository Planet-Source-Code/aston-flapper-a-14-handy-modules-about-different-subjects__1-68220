Attribute VB_Name = "SpecialApi"
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function ReleaseDC& Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetComputerName Lib "kernel32.dll" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetLocalTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SetDoubleClickTime Lib "user32.dll" (ByVal wCount As Long) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SetWindowPos Lib "USER32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long ' Get the cursor position
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long  ' Get the handle of the window that is foremost on a particular X, Y position. Used here To get the window under the cursor
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Private Declare Function GetKeyBoardState Lib "user32" Alias "GetKeyboardState" (kbArray As KeyboardBytes) As Long
Private Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PaintDesktop Lib "user32.dll" (ByVal hdc As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Dim kbArray As KeyboardBytes, CapsLock As Boolean

Public Const VK_CAPITAL = &H14
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_USED = VK_SCROLL
Public Const WndKladblok = "Notepad"
Public Const WndInternet_Explorer = "Ieframe"
Public Const WndStartButton = "Button"
Public Const WndDesktop = "Progman"
Public Const WndTask = "Shell_traywnd"
Public Const WndWinCalculator = "SciCalc"
Public Const WndWinPaint = "MsPaintApp"
Public Const WndMicrosoftWord = "OpusApp"
Public Const WndWordPad = "WordPadClass"
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const ERROR_SUCCESS = 0&
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const TokenPrivileges = 3
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const ANYSIZE_ARRAY = 1
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const MOUSEEVENTF_LEFTDOWN          As Long = &H2
Private Const MOUSEEVENTF_LEFTUP            As Long = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN        As Long = &H20
Private Const MOUSEEVENTF_MIDDLEUP          As Long = &H40
Private Const MOUSEEVENTF_RIGHTDOWN         As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP           As Long = &H10
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Type POINTAPI
x As Long
y As Long
End Type

Private Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type

Public Type FileInfo
SectorsPerCluster As Long
BytesPerSector As Long
FreeClusters As Long
TotalClusters As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Type Luid
    lowpart As Long
    highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    'pLuid As Luid
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RGB
Red As Integer
Green As Integer
Blue As Integer
End Type

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
    szexeFile As String * 260
End Type

Public Enum ButtonsSwapped
Normal = &H0
Swapped = &H1
End Enum

Public Function XGetMouseHwnd() As Long
Dim Cursor As POINTAPI ' Cursor position
Dim RetVal As Long ' Dummy returnvalue
Dim hdc As Long ' hDC that we're going To be using
GetCursorPos Cursor
RetVal = WindowFromPoint(Cursor.x, Cursor.y)
XGetMouseHwnd = RetVal
End Function

Public Function XGetMouseX() As Long
Dim Cursor As POINTAPI ' Cursor position
GetCursorPos Cursor
XGetMouseX = Cursor.x
End Function

Public Function XGetMouseY() As Long
Dim Cursor As POINTAPI ' Cursor position
GetCursorPos Cursor
XGetMouseY = Cursor.y
End Function

Public Sub XKeyboardStateTurnOn(vkKey As Long)
GetKeyBoardState kbArray
kbArray.kbByte(vkKey) = 1
SetKeyboardState kbArray
End Sub
Public Sub XKeyboardStateTurnOff(vkKey As Long)
GetKeyBoardState kbArray
kbArray.kbByte(vkKey) = 0
SetKeyboardState kbArray
End Sub

Public Function XGetDiskInfo() As FileInfo
Dim Sectors As Long, Bytes As Long, FreeC As Long, TotalC As Long, Total As Long, Freeb As Long
GetDiskFreeSpace "C:\", Sectors, Bytes, FreeC, TotalC
XGetDiskInfo.BytesPerSector = Bytes
XGetDiskInfo.SectorsPerCluster = Sectors
XGetDiskInfo.FreeClusters = FreeC
XGetDiskInfo.TotalClusters = TotalC
End Function

Public Sub XSwapMouseButtons(Swapped As ButtonsSwapped)
Dim Swapped2 As Integer
Swapped2 = Swapped
SwapMouseButton Swapped
End Sub


Public Function XGetComName() As String
Dim dwLen As Long
Dim strString As String
dwLen = 32
strString = String(dwLen, "X")
GetComputerName strString, dwLen
strString = Left(strString, dwLen)
XGetComName = strString
End Function

Public Sub XChangeColors(hwnd, ForeGroundColor As OLE_COLOR, BackGroundColor As OLE_COLOR)
SendMessage hwnd, PBM_SETBARCOLOR, 0, ByVal ForeGroundColor
SendMessage hwnd, PBM_SETBKCOLOR, 0, ByVal BackGroundColor
End Sub

Public Sub XHideStartButton()
Dim ssave As String * 200
Dim tWnd As Integer
Dim bwnd As Integer
tWnd = FindWindow("Shell_traywnd", vbNullString)
bwnd = GetWindow(tWnd, GW_CHILD)
Do
GetClassName bwnd, ssave, 250
If LCase(Left$(ssave, 6)) = "button" Then Exit Do
bwnd = GetWindow(bwnd, GW_HWNDNEXT)
Loop
SetWindowPos bwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
End Sub

Private Function XSendMCIString(cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 200
rc = mciSendString(cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
mciGetErrorString rc, errStr, Len(errStr)
MsgBox errStr
End If
XSendMCIString = (rc = 0)
End Function

Public Sub XCDOpen()
XSendMCIString "close all", False
If (App.PrevInstance = True) Then
End
End If
fCDLoaded = False
If (XSendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
End
End If
XSendMCIString "set cd time format tmsf wait", True
XSendMCIString "set cd door open", False
End Sub
Public Sub XCDClose()
XSendMCIString "close all", False
If (App.PrevInstance = True) Then
End
End If
fCDLoaded = False
If (XSendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
End
End If
XSendMCIString "set cd time format tmsf wait", True
XSendMCIString "set cd door closed", False
End Sub

Public Sub XShutdownComputer()
ExitWindowsEx 1, 0&
End Sub

Public Sub XRestartComputer()
ExitWindowsEx 2, 0&
End Sub

Public Sub XHideTask()
Dim tReturn
tReturn = FindWindow("Shell_traywnd", "")
SetWindowPos tReturn, 0, 0, 0, 0, 0, &H80
End Sub

Public Sub XShowTask()
Dim tReturn As Long
tReturn = FindWindow("Shell_traywnd", "")
SetWindowPos tReturn, 0, 0, 0, 0, 0, &H40
End Sub

Public Sub XClearBin()
SHEmptyRecycleBin hwnd, "", &H2
End Sub

Public Sub XShowDesktop()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 5
End Sub

Public Sub XHideDesktop()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 0
End Sub

Public Sub XShowWindow(WindowName As String)
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, WindowName, vbNullString)
ShowWindow hwnd, 5
End Sub

Public Sub XHideWindow(WindowName As String)
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, WindowName, vbNullString)
ShowWindow hwnd, 0
End Sub

Public Sub XShowHwnd(hwnd As Long)
ShowWindow hwnd, 5
End Sub

Public Sub XHideHwnd(hwnd As Long)
ShowWindow hwnd, 0
End Sub

Public Sub XSleep(Milliseconds As Long)
Sleep Milliseconds
End Sub

Public Function XGetWindowHwnd(WindowName As String) As Long
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, WindowName, vbNullString)
XGetWindowHwnd = hwnd
End Function


Public Function XGetRGB(dsColor As OLE_COLOR) As RGB
Dim Hexcolor As String
Dim HexRed As String
Dim HexGreen As String
Dim HexBlue As String
Hexcolor = Hex(dsColor)
If Len(Hexcolor) = 1 Then Hexcolor = "00000" & Hexcolor
If Len(Hexcolor) = 2 Then Hexcolor = "0000" & Hexcolor
If Len(Hexcolor) = 3 Then Hexcolor = "000" & Hexcolor
If Len(Hexcolor) = 4 Then Hexcolor = "00" & Hexcolor
If Len(Hexcolor) = 5 Then Hexcolor = "0" & Hexcolor
HexRed = Mid$(Hexcolor, 1, 2)
HexGreen = Mid$(Hexcolor, 3, 2)
HexBlue = Mid$(Hexcolor, 5, 2)
XGetRGB.Red = Decimall(HexBlue)
XGetRGB.Green = Decimall(HexGreen)
XGetRGB.Blue = Decimall(HexRed)
End Function
Private Function Decimall(Getal As String) As Long
Dim t(2) As Long
Dim g(2) As String
Dim d(2) As Long
Dim s(2) As Long
t(2) = 1
t(1) = 16
g(2) = Mid$(Getal, 2, 1)
g(1) = Mid$(Getal, 1, 1)
Select Case g(2)
Case "1"
d(2) = 1
Case "2"
d(2) = 2
Case "3"
d(2) = 3
Case "4"
d(2) = 4
Case "5"
d(2) = 5
Case "6"
d(2) = 6
Case "7"
d(2) = 7
Case "8"
d(2) = 8
Case "9"
d(2) = 9
Case "A"
d(2) = 10
Case "B"
d(2) = 11
Case "C"
d(2) = 12
Case "D"
d(2) = 13
Case "E"
d(2) = 14
Case "F"
d(2) = 15
End Select
Select Case g(1)
Case "1"
d(1) = 1
Case "2"
d(1) = 2
Case "3"
d(1) = 3
Case "4"
d(1) = 4
Case "5"
d(1) = 5
Case "6"
d(1) = 6
Case "7"
d(1) = 7
Case "8"
d(1) = 8
Case "9"
d(1) = 9
Case "A"
d(1) = 10
Case "B"
d(1) = 11
Case "C"
d(1) = 12
Case "D"
d(1) = 13
Case "E"
d(1) = 14
Case "F"
d(1) = 15
End Select
s(1) = d(1) * t(1)
s(2) = d(2) * t(2)
Decimall = s(1) + s(2)
End Function

Public Sub XSetFocus(hwnd As Long)
SetFocus hwnd
End Sub
Public Sub XSetCursorPos(x As Long, y As Long)
SetCursorPos x, y
End Sub
Public Sub XSetWindowText(hwnd As Long, Text As String)
SetWindowText hwnd, Text
End Sub

Public Sub XSetDoubleClickTime(Milliseconds As Long)
SetDoubleClickTime Milliseconds
End Sub

Public Sub XTypeText(Text As String, Optional Delay As Long = -1)
If Delay = -1 Then
SendKeys Text
Else
SendKeys Text, Delay
End If
End Sub

Public Sub XSetWindowPos(hwnd As Long, x As Long, y As Long, X2 As Long, Y2 As Long)
SetWindowPos hwnd, 0, x, y, X2, Y2, 0
End Sub

Public Function XGetDc(hwnd As Long)
XGetDc = GetDC(hwnd)
End Function

Public Sub XStuckWindow(hwnd As Long)
SetCapture hwnd
End Sub

Public Sub XSetComName(NewName As String)
SetComputerName NewName
End Sub

Public Sub XSetPixel(hwnd As Long, x As Long, y As Long, Color As OLE_COLOR)
SetPixel GetDC(hwnd), x, y, Color
End Sub

Public Sub XPaintDesktop(hwnd As Long)
PaintDesktop XGetDc(hwnd)
End Sub

Public Sub XClickMouseButton(Optional MouseButton As MouseButtonConstants = vbLeftButton)
    If (MouseButton = vbLeftButton) Then
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
     ElseIf (MouseButton = vbMiddleButton) Then 'NOT (MOUSEBUTTON...
        Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0&, 0&, 0&, 0&)
     ElseIf (MouseButton = vbRightButton) Then 'NOT (MOUSEBUTTON...
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0&, 0&, 0&, 0&)
    End If
End Sub

Public Function XGetScreenColor(x As Long, y As Long) As OLE_COLOR
rDC& = GetDC(0&)
rPixel& = GetPixel(rDC&, x, y)
ReleaseDC 0&, rDC&
XGetScreenColor = rPixel
End Function

Public Function XHexToDecimall(Num As String)
For a = 1 To Len(Num)
If Mid$(Num, a, 1) <> "0" Then
Exit For
Else
zh = True
End If
Next
If zh = True Then Num = Mid$(Num, a)
Num = UCase$(Num)
Dim nums(13) As Currency
nums(1) = 1
nums(2) = 16
For a = 3 To 13
nums(a) = nums(a - 1) * 16
Next
For a = Len(Num) To 1 Step -1
g = g + Mid$(Num, a, 1)
Next
Num = g
For a = 1 To Len(Num)
gh = Mid$(Num, a, 1)
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

Public Function XDecimallToBinair(Num As Long)
Dim nums(50) As Currency
nums(1) = 1
For a = 2 To 50
nums(a) = nums(a - 1) * 2
Next
For a = 50 To 1 Step -1
If nums(a) < Num Then Exit For
Next
For b = a To 1 Step -1
If Num >= nums(b) Then
Num = Num - nums(b)
g = g & "1"
Else
g = g & "0"
End If
Next
DecimallToBinair = g
End Function
Public Function XBinairToDecimall(Num As String)
Dim nums(50) As Currency
nums(1) = 1
For a = 2 To 50
nums(a) = nums(a - 1) * 2
Next
For a = Len(Num) To 1 Step -1
g = g & Mid$(Num, a, 1)
Next
Num = g
g = 0
For a = 1 To Len(Num)
gg = Mid$(Num, a, 1)
If gg = 1 Then
g = g + nums(a)
End If
Next

BinairToDecimall = g
End Function

Public Sub XRegisterFileType(IconPath As String, ProgPath As String, FileType As String)
    
SaveKey HKEY_CLASSES_ROOT, FileType ' your new file type
SaveKey HKEY_CLASSES_ROOT, FileType & "\DefaultIcon"  ' your new file types icon root
SaveKey HKEY_CLASSES_ROOT, FileType & "\shell"
SaveKey HKEY_CLASSES_ROOT, FileType & "\shell\open"
SaveKey HKEY_CLASSES_ROOT, FileType & "\shell\open\command"
SaveString HKEY_CLASSES_ROOT, FileType & "\DefaultIcon", "", IconPath ' your new filetype icon to use
SaveString HKEY_CLASSES_ROOT, FileType & "\shell\open\command", "", Chr(34) & ProgPath & Chr(34) & " %1"
End Sub

Private Sub SaveKey(hKey As Long, strPath As String)
Dim Keyhand&
    r = RegCreateKey(hKey, strPath, Keyhand&)
    r = RegCloseKey(Keyhand&)
End Sub

Private Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim Keyhand As Long
Dim r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    r = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(Keyhand)
End Sub

Public Function XThreeDtoTwoD(x As Long, y As Long, z As Long) As POINTAPI
XThreeDtoTwoD.x = x + z
XThreeDtoTwoD.y = y + z
End Function

Public Sub XPlayWave(File As String, InLoop As Boolean)
Dim SoundName As String
SoundName$ = File
If InLoop = False Then
wFlags% = SND_ASYNC Or SND_NODEFAULT
Else
wFlags% = SND_ASYNC Or SND_LOOP
End If
x = sndPlaySound(SoundName$, wFlags%)
End Sub

Public Sub XStopWave()
Dim SoundName As String
SoundName$ = " "
wFlags% = SND_ASYNC Or SND_NODEFAULT
x = sndPlaySound(SoundName$, wFlags%)
End Sub

Public Sub XStopWaveRepeat(File As String)
Dim SoundName As String
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
x = sndPlaySound(SoundName$, wFlags%)
End Sub

Public Sub XDestroyFile(sFileName As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1
    For iLoop = 1 To Blocks
        offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, offset, Block2
    Next iLoop
    Close hFileHandle
    Kill sFileName
End Sub

Public Function XFadeColors(Percentage, Max, Color1 As OLE_COLOR, Color2 As OLE_COLOR) As OLE_COLOR
'On Error Resume Next
color1r = XGetRGB(Color1).Red
color1g = XGetRGB(Color1).Green
color1b = XGetRGB(Color1).Blue

color2r = XGetRGB(Color2).Red
color2g = XGetRGB(Color2).Green
color2b = XGetRGB(Color2).Blue
If color1r = color2r Then color3r = color1r
If color1g = color2g Then color3g = color1g
If color1b = color2b Then color3b = color1b
'color3r = XLowestNum(color1r, color2r) + ((XHighestNum(color1r, color2r) - XLowestNum(color1r, color2r)) * (Percentage / Max))
'color3g = XLowestNum(color1g, color2g) + ((XHighestNum(color1g, color2g) - XLowestNum(color1g, color2g)) * (Percentage / Max))
'color3b = XLowestNum(color1b, color2b) + ((XHighestNum(color1b, color2b) - XLowestNum(color1b, color2b)) * (Percentage / Max))
color3r = color1r + (color2r - color1r) * (Percentage / Max)
color3g = color1g + (color2g - color1g) * (Percentage / Max)
color3b = color1b + (color2b - color1b) * (Percentage / Max)

XFadeColors = RGB(color3r, color3g, color3b)
End Function

Public Function XLowestNum(Num1, Num2)
If Num1 > Num2 Then
XLowestNum = Num2
ElseIf Num1 < Num2 Then
XLowestNum = Num1
Else
XLowestNum = Num1
End If
End Function

Public Function XHighestNum(Num1, Num2)
If Num1 > Num2 Then
XHighestNum = Num1
ElseIf Num1 < Num2 Then
XHighestNum = Num2
Else
XHighestNum = Num1
End If
End Function

Public Sub XWindowFlash(hwnd As Long, XStart As Boolean)
On Error Resume Next
Dim returnval As Integer
returnval = FlashWindow(hwnd, XStart)
End Sub

Public Sub XFormOnTop(hWindow As Long, bTopMost As Boolean)
wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
Select Case bTopMost
Case True
Placement = HWND_TOPMOST
Case False
Placement = HWND_NOTOPMOST
End Select
SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub

Public Function XMsgBox(hwnd As Long, Text As String, Optional Title As String = "x92746x k815k n0Title0n", Optional xType As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxStyle
If Title = "x92746x k815k n0Title0n" Then Title = App.Title
XMsgBox = MessageBox(hwnd, Text, Title, xType)
End Function

Public Function XRndNum(Lowest, Highest)
Lowest = Lowest - 1
Highest = Highest + 1
Randomize Timer
XRndNum = Int(Rnd * (Highest - Lowest - 1)) + Lowest + 1
End Function

Public Function XRndStr(xStr As String)
XRndStr = Mid$(xStr, XRndNum(1, Len(xStr)), 1)
End Function

Public Function XRndString(xName As String, xStr1 As String, Optional xStr2 As String = "", Optional xStr3 As String = "", Optional xStr4 As String = "")
For a = 1 To Len(xName)
b = Mid$(xName, a, 1)
If InStr(xStr1, b) <> 0 Then g = g & XRndStr(xStr1)
If InStr(xStr2, b) <> 0 Then g = g & XRndStr(xStr2)
If InStr(xStr3, b) <> 0 Then g = g & XRndStr(xStr3)
If InStr(xStr4, b) <> 0 Then g = g & XRndStr(xStr4)
Next
XRndString = g
End Function

Public Function XRndName(xName As String)
XRndName = XRndString(xName, "aeiou", "bcdfghjklmnpqrstvwxyz", "AEIOU", "BCDFGHJKLMNPQRSTVWXYZ")
End Function

Public Function XTextBetween(Text As String, Str1 As String, Str2 As String)
XTextBetween = Left$(Mid$(Text, InStr(Text, Str1) + Len(Str1)), InStr(Mid$(Text, InStr(Text, Str1) + Len(Str1)), Str2) - 1)
End Function


Public Sub XPutProcess(TheList As ListBox)
TheList.Clear
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapshot = 0 Then Exit Sub
uProcess.dwSize = Len(uProcess)
r = ProcessFirst(hSnapshot, uProcess)
Do While r
TheList.AddItem uProcess.szexeFile
r = ProcessNext(hSnapshot, uProcess)
Loop
Call CloseHandle(hSnapshot)
TheList.ListIndex = 0
End Sub


