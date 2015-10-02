Attribute VB_Name = "taskmngr11"
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, wParam As Any, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Public Declare Function SendMessageTimeout Lib "user32" _
    Alias "SendMessageTimeoutA" (ByVal hwnd As Long, _
    ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
    ByVal fuFlags As Long, ByVal uTimeout As Long, _
    pdwResult As Long) As Long

Public Const SMTO_BLOCK = &H1
Public Const SMTO_ABORTIFHUNG = &H2
Public Const WM_NULL = &H0

Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
           ByVal hwnd As Long, _
           ByVal wMsg As Long, _
           ByVal wParam As Long, _
           ByVal lParam As Long _
) As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public tmplog As String
Public tmpdeskhwnd As Single
Public tmptaskbarhwnd As Single



Public Function getalltopwindows(ByVal hwnd As Long, ByVal lParam As Long) As Long

Dim foregroundwindow As Long
Dim textlen As Long
Dim windowtext As String
Dim svar As Long
Static lastwindowtext As String



foregroundwindow = hwnd


textlen = GetWindowTextLength(foregroundwindow) + 1

windowtext = Space(textlen)
svar = GetWindowText(foregroundwindow, windowtext, textlen)
windowtext = Left(windowtext, Len(windowtext) - 1)

If windowtext = "" Then GoTo slask

If Form1.Check2.Value = 1 Then
If IsWindowVisible(foregroundwindow) > 0 Then

If windowtext = Form1.Caption Then GoTo slask
Form1.List1.AddItem windowtext
Form1.List1.ItemData(Form1.List1.NewIndex) = foregroundwindow
lastwindowtext = windowtext

End If

Else
If windowtext = Form1.Caption Then GoTo slask
Form1.List1.AddItem windowtext
Form1.List1.ItemData(Form1.List1.NewIndex) = foregroundwindow
lastwindowtext = windowtext


End If



slask:



getalltopwindows = 1
End Function


