Attribute VB_Name = "Modusle1"
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Const GWL_EXSTYLE = (-20)
    Public Const WS_EX_APPWINDOW = &H40000

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Public Const SW_NORMAL = 1
Public nowuser As String
Rem œ«·… ·· Õﬁﬁ „‰ ÊÃÊœ „·› „⁄Ì¯‰
Public Function FileExists(strPath As String) As Boolean
    strPath = Trim(strPath)
    If strPath = "" Then
        FileExists = False
        Exit Function
    End If
  FileExists = Len(Dir(strPath)) <> 0
End Function


'œ«·…œÊ„« ›Ì «·„ﬁœ„…
Public Sub StayOnTop(frm As Form)
  SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

'› Õ „Êﬁ⁄ «‰ —‰ 

Public Sub OpenWebsite(strWebsite As String)
  If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
    ' Insert Error handling code here
  End If
End Sub

'«” Œ—«Ã «”„ „·›
Public Function GetFileNameFromPath(strPath As String) As String
  Dim intX As Integer
  Dim intPlace As Integer
  Dim intLastPlace As Integer
    
  intLastPlace = 0

  For intX = 1 To Len(strPath)
    intPlace = InStr(intLastPlace + 1, strPath, "\")
    
    If intPlace = 0 Then
      GetFileNameFromPath = Right(strPath, Len(strPath) - intLastPlace)
      Exit Function
    Else
      intLastPlace = intPlace
    End If
  Next intX

End Function


