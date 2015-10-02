Attribute VB_Name = "Modsule2"
#If Win32 Then
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile _
As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
#Else
Declare Function ShellExecute Lib "SHELL" (ByVal hwnd%, _
ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, _
ByVal lpszDir$, ByVal fsShowCmd%) As Integer
Declare Function GetDesktopWindow Lib "USER" () As Integer
#End If
Public Const SW_SHOWNORMAL = 1

'Insert this code to your form:

Function StartDoc(DocName As String) As Long
Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function


