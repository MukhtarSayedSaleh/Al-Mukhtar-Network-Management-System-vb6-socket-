VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{E5CEE37F-8CF8-489E-BFA0-8201CBD6AEE8}#1.0#0"; "PicFormat32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰Ÿ«„ «·„Œ «— ·≈œ«—… „ﬁ«ÂÌ «·≈‰ —‰ "
   ClientHeight    =   10680
   ClientLeft      =   10635
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   4680
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "frm_main.frx":29C12
      Top             =   5160
   End
   Begin ACTIVESKINLibCtl.SkinLabel Label2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frm_main.frx":29E46
      TabIndex        =   11
      Top             =   8760
      Width           =   3495
   End
   Begin ACTIVESKINLibCtl.SkinLabel Label1 
      Height          =   375
      Left            =   3600
      OleObjectBlob   =   "frm_main.frx":29EBC
      TabIndex        =   10
      Top             =   8760
      Width           =   975
   End
   Begin VB.TextBox timelen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   3120
   End
   Begin VB.TextBox usedtime 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin PicFormat32a.PicFormat32 PicFormat321 
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command5 
      Caption         =   " ”ÃÌ· «·Œ—ÊÃ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "≈⁄œ«œ«  «·‰Ÿ«„"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   9240
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "≈€·«ﬁ"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈Œ›«¡ «·‘«‘…"
      Height          =   375
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«·ﬂ«› —Ì«"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   10200
      Width           =   4455
   End
   Begin VB.TextBox ipa 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   15055
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long

'’Ê—… ·”ÿÕ «·„ﬂ »
Private Const VK_SNAPSHOT = &H2C

Const SE_PRIVILEGE_ENABLED = &H2
Const TokenPrivileges = 3
Const TOKEN_ASSIGN_PRIMARY = &H1
Const TOKEN_DUPLICATE = &H2
Const TOKEN_IMPERSONATE = &H4
Const TOKEN_QUERY = &H8
Const TOKEN_QUERY_SOURCE = &H10
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_ADJUST_GROUPS = &H40
Const TOKEN_ADJUST_DEFAULT = &H80
Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Const ANYSIZE_ARRAY = 1
Private Type LARGE_INTEGER
   lowpart As Long
   highpart As Long
End Type
Private Type LUID
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
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long



'**************************************************************
'to Shutdown WIndows 98
   Const EWX_LOGOFF = 0
   Const EWX_SHUTDOWN = 1
   Const EWX_REBOOT = 2
   Const EWX_FORCE = 4
   Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'************************************************************
'to Get windows versiton
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Dim strOs As String   'Operating System
Dim strWinV As String 'Windows Version

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
  Dim l_ExitWindows As New cls_ExitWindows



'**************************************************************
Public Function InitiateShutdownMachine(ByVal Machine As String, Optional Force As Variant, Optional Restart As Variant, Optional AllowLocalShutdown As Variant, Optional Delay As Variant, Optional Message As Variant) As Boolean
   Dim hProc As Long
   Dim OldTokenStuff As TOKEN_PRIVILEGES
   Dim OldTokenStuffLen As Long
   Dim NewTokenStuff As TOKEN_PRIVILEGES
   Dim NewTokenStuffLen As Long
   Dim pSize As Long
   If IsMissing(Force) Then Force = False
   If IsMissing(Restart) Then Restart = True
   If IsMissing(AllowLocalShutdown) Then AllowLocalShutdown = False
   If IsMissing(Delay) Then Delay = 0
   If IsMissing(Message) Then Message = ""
   'Make sure the Machine-name doesn't start with '\\'
   If InStr(Machine, "\\") = 1 Then
       Machine = Right(Machine, Len(Machine) - 2)
   End If
   'check if it's the local machine that's going to be shutdown
   If (LCase(GetMyMachineName) = LCase(Machine)) Then
       'may we shut this computer down?
       If AllowLocalShutdown = False Then Exit Function
       'open access token
       If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc) = 0 Then
           MsgBox "OpenProcessToken Error: " & GetLastError()
           Exit Function
       End If
       'retrieve the locally unique identifier to represent the Shutdown-privilege name
       If LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, OldTokenStuff.Privileges(0).pLuid) = 0 Then
           MsgBox "LookupPrivilegeValue Error: " & GetLastError()
           Exit Function
       End If
       NewTokenStuff = OldTokenStuff
       NewTokenStuff.PrivilegeCount = 1
       NewTokenStuff.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
       NewTokenStuffLen = Len(NewTokenStuff)
       pSize = Len(NewTokenStuff)
       'Enable shutdown-privilege
       If AdjustTokenPrivileges(hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen) = 0 Then
           MsgBox "AdjustTokenPrivileges Error: " & GetLastError()
           Exit Function
       End If
       'initiate the system shutdown
       If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
           Exit Function
       End If
       NewTokenStuff.Privileges(0).Attributes = 0
       'Disable shutdown-privilege
       If AdjustTokenPrivileges(hProc, False, NewTokenStuff, Len(NewTokenStuff), OldTokenStuff, Len(OldTokenStuff)) = 0 Then
           Exit Function
       End If
   Else
       'initiate the system shutdown
       If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
           Exit Function
       End If
   End If
   InitiateShutdownMachine = True
End Function

Function GetMyMachineName() As String
   Dim sLen As Long
   'create a buffer
   GetMyMachineName = Space(100)
   sLen = 100
   'retrieve the computer name
   If GetComputerName(GetMyMachineName, sLen) Then
       GetMyMachineName = Left(GetMyMachineName, sLen)
   End If
End Function


Private Sub Command1_Click()
If Winsock1.State = 7 Then
Winsock1.SendData "caftlst"
End If

End Sub

Private Sub Command2_Click()
frm_main.Hide
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
FrmSetting.Show
StayOnTop FrmSetting
End Sub

Private Sub Command5_Click()
If Winsock1.State = 7 Then
Winsock1.SendData "logoutt" & nowuser & "%%" & usedtime.Text
frm_closed.Show
frm_main.Hide
Timer1.Enabled = False
End If
End Sub

Private Sub Form_Load()
'«·”ﬂ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin Me.hwnd

Rem  Õ„Ì· ’›Õ… «·œ⁄«Ì…
WebBrowser1.Navigate App.Path & ("\Adv.MPage")
StayOnTop Me
End Sub


Private Sub Timer1_Timer()
usedtime.Text = usedtime.Text + 1
If IsNumeric(timelen.Text) = True Then
If Int(timelen.Text) < Int(usedtime.Text / 60) Then
Call Command5_Click
End If
End If
End Sub

Private Sub Winsock1_Connect()
If Winsock1.State = 7 Then
Winsock1.SendData "ready"
End If
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim s As String
Dim ssxx As String
Dim web As String
Winsock1.GetData s
ssxx = Mid(s, 1, 7)
web = Mid(s, 8, Len(s))
'› Õ «·ÃÂ«“
If s = "[mokhatropen]" Then
 frm_closed.Hide
 frm_main.Show
 frm_main.Timer1.Enabled = True
 frm_main.label1.Visible = False
 frm_main.label2.Visible = False
 usedtime.Text = "0"
 
 If Winsock1.State = 7 Then
  Winsock1.SendData "opened"
 End If
End If
'≈ë€·«ﬁ «·ÃÂ«“
If s = "[mokhatrclose]" Then
 frm_main.Hide
 frm_closed.Show
End If
'’Ê—… ·”ÿÕ «·„ﬂ »
If s = "[mokdesktopsnap]" Then
 '«· ﬁ«ÿ «·’Ê—… √Ê·«
 keybd_event VK_SNAPSHOT, 0, 0, 0
 DoEvents
 SavePicture Clipboard.GetData(), "c:\snap1.snp"
 ' ÕÊÌ· «·’Ê—… «·Ï jpg «Ê ≈·Ï gif
 PicFormat321.SaveBmpToGif "c:\snap1.snp", "c:\snap1.gif"
 '≈—”«· «·’Ê—… »⁄œ «·≈· ﬁ«ÿ
 Dim cmd As String
 Dim ss As String
 Dim msg As String
 Dim i As Integer
 Dim f As Integer
 Dim p As Long
 Const Size = 2048
 f = FreeFile
 Open "c:\snap1.gif" For Binary As f
 For p = 1 To LOF(f) \ Size
 ss = Space(Size)
 Get #f, , ss
 If Winsock1.State = 7 Then
 Winsock1.SendData "newpart" & ss
 Else
 GoTo bye
 End If
 DoEvents
 Next
 If LOF(f) Mod Size > 0 Then
 ss = Space(LOF(f) Mod Size)
 If Winsock1.State = 7 Then
 Get #f, , ss
 Winsock1.SendData "newpart" & ss
 DoEvents
 Else
 GoTo bye
 End If
 End If
 If Winsock1.State = 7 Then
 Winsock1.SendData "endfile"
 DoEvents
 End If
bye:
 Close f
End If
 '› Õ ’›Õ… ≈‰ —‰ 
If ssxx = "[wbrse]" Then
 OpenWebsite (web)
End If
'—”«·… Œ«’…
If ssxx = "[pmmsg]" Then
 pmmsg.Text1.Text = web
 pmmsg.Show
End If
' ·ﬁÌ √«„·›
'›Ì «·»œ«Ì…  ·ﬁ¯Ì «”„ Ê ‰Ê⁄ «·„·›
If ssxx = "filesou" Then
If FileExists("c:\mokhtar.tmp") = True Then
Kill "c:\mokhtar.tmp"
End If
End If
'⁄‰œ  ·ﬁ¯Ì ﬁ”„ ÃœÌœ „‰ «·„·›
If ssxx = "newfart" Then
Dim filee As Integer
filee = FreeFile
Open "c:\mokhtar.tmp" For Binary As filee
Put #filee, (LOF(filee) + 1), web
Close filee
End If
'⁄‰œ Ê’Ê· «·ﬁ”„ «·√ŒÌ—
If ssxx = "enffile" Then
 Dim lahe As String
 '„·«ÕŸ… : «·ŒÊ«—“„Ì¯… ’ÕÌÕ… ·ﬂ‰ ⁄‰œ «· ‰›Ì– ›Ì «·„—¯… «·√ŒÌ—…
 '·‰ ÌﬂÊ‰ Â‰«ﬂ «”„ „·› ﬂÌ Ì»ﬁÏ ›Ì «·–«ﬂ—…
 '›„« «·Õ· ø
 lahe = "c:\" & GetFileNameFromPath(web)
 Close filee
 CopyFile "c:\mokhtar.tmp", lahe, False
 If FileExists("c:\mokhtar.tmp") = True Then
  Kill "c:\mokhtar.tmp"
 End If
End If
'⁄‰œ «—”«· √„— Êﬁ› «·≈—”«·
If ssxx = "stpsend" Then
 Close filee
 If FileExists("c:\mokhtar.tmp") = True Then
 Kill "c:\mokhtar.tmp"
 End If
End If
'«„—  ‘€Ì· «·„·› »⁄œ Ê’Ê·Â
If ssxx = "runfile" Then
 Dim lahe2 As String
 '„·«ÕŸ… : «·ŒÊ«—“„Ì¯… ’ÕÌÕ… ·ﬂ‰ ⁄‰œ «· ‰›Ì– ›Ì «·„—¯… «·√ŒÌ—…
 '·‰ ÌﬂÊ‰ Â‰«ﬂ «”„ „·› ﬂÌ Ì»ﬁÏ ›Ì «·–«ﬂ—…
 '›„« «·Õ· ø
 lahe2 = "c:\" & GetFileNameFromPath(web)
 If FileExists(lahe2) = True Then
 Dim r As Long
 'Replace the c:\mp3\song.mp3 with the file you want to launch
 r = StartDoc(lahe2)
 End If
End If
 
 ' ÕœÌÀ ”ÿÕ «·„ﬂ »
If ssxx = "refdesk" Then
 InvalidateRect 0&, 0&, False
End If

'≈€·«ﬁ «·⁄„Ì·
If ssxx = "clsclnt" Then
 Winsock1.Close
 End
End If
'≈⁄«œ… «·≈ﬁ·«⁄
If ssxx = "[rebot]" Then
    l_ExitWindows.ExitWindows WE_REBOOT     '\\ Reboot was selected
End If
'≈Ìﬁ«› «· ‘€Ì·
If ssxx = "[shtdn]" Then
    l_ExitWindows.ExitWindows WE_SHUTDOWN   '\\ Shutdown was selected
End If

'› Õ ··√œ„‰
If ssxx = "adlogok" Then
 frm_closed.Hide
 frm_main.Show
 frm_main.Command3.Enabled = True
 frm_main.Command4.Enabled = True
 frm_main.label1.Visible = True
frm_main.label2.Visible = True
frm_main.label2.Caption = "Admin"
 Unload signin
End If

'› Õ ··ÌÊ“—“ «·⁄«œÌÌ‰
If ssxx = "usrlogn" Then
timelen.Text = web
frm_closed.Hide
frm_main.Show
frm_main.Command3.Enabled = False
frm_main.Command4.Enabled = False
frm_main.Timer1.Enabled = True
frm_main.label1.Visible = True
frm_main.label2.Visible = True
frm_main.label2.Caption = nowuser
usedtime.Text = 0
Unload signin
End If

'›‘· ›Ì  ”ÃÌ· «·œŒÊ·
If ssxx = "loginff" Then
MsgBox "«”„ «·„” Œœ„ √Ê ﬂ·„… «·„—Ê— Œ«ÿ∆…", 16, "Œÿ√"
End If

'≈‰Â«¡ Ã·”… «·⁄„·
If ssxx = "[endjb]" Then
If Winsock1.State = 7 Then
Winsock1.SendData "endjobs" + usedtime.Text
frm_closed.Show
frm_main.Hide
Timer1.Enabled = False
End If
End If

'ﬁ«∆„… «· ÿ»Ìﬁ« 
If ssxx = "taskmgr" Then
 Dim tasks As String
 Dim n As Integer
 form1.Show
 tasks = form1.List1.List(0)
 For n = 1 To form1.List1.ListCount - 1
 If form1.List1.List(n) <> App.ProductName Then
  tasks = tasks & "%" & form1.List1.List(n)
 End If
 Next
 Unload form1
 If Winsock1.State = 7 Then
  Winsock1.SendData "taskmgr" & tasks & "%"
 End If
End If

' ≈‰Â«¡  ÿ»Ìﬁ „⁄Ì¯‰
If ssxx = "endtask" Then
 Dim appe As String
 appe = form1.List1.ItemData(web)
 On Error Resume Next
 svar = GetWindowThreadProcessId(appe, nyprocessid)
 procname = OpenProcess(PROCESS_ALL_ACCESS, 0&, nyprocessid)
 svar2 = TerminateProcess(procname, 0&)
 DoEvents
 Unload form1
 '»⁄œ ≈‰Â«¡ «· ÿ»Ìﬁ ⁄„·  ÕœÌÀ
 Dim tasks2 As String
 Dim n2 As Integer
 form1.Show
 tasks2 = form1.List1.List(0)
 For n2 = 0 To form1.List1.ListCount - 1
 If form1.List1.List(n2) <> App.ProductName Then
  tasks2 = tasks2 & "%" & form1.List1.List(n2)
 End If
 Next
 Unload form1
 If Winsock1.State = 7 Then
  Winsock1.SendData "taskmgr" & tasks2 & "%"
 End If
End If

'≈‰Â«¡ Ã„Ì⁄ «· ÿ»Ìﬁ« 

If s = "endalltasks" Then
 Dim z As Integer
 Dim aap As String
 For z = 0 To form1.List1.ListCount - 1
  aap = form1.List1.ItemData(z)
  If aap <> App.ProductName Then
    On Error Resume Next
    svar = GetWindowThreadProcessId(aap, nyprocessid)
    procname = OpenProcess(PROCESS_ALL_ACCESS, 0&, nyprocessid)
    svar2 = TerminateProcess(procname, 0&)
    DoEvents
  End If
 Next
'»⁄œ ≈‰Â«¡ «· ÿ»Ìﬁ ⁄„·  ÕœÌÀ
 Dim tasks3 As String
 Dim n3 As Integer
 tasks3 = form1.List1.List(0)
 For n3 = 0 To form1.List1.ListCount - 1
 If form1.List1.List(n3) <> App.ProductName Then
 tasks3 = tasks3 & "%" & form1.List1.List(n3)
 End If
 Next
 Unload form1
 If Winsock1.State = 7 Then
 Winsock1.SendData "taskmgr" & tasks3 & "%"
 End If
End If

'›Ì Õ«· ⁄œ„ ÊÃÊœ √’‰«› ›Ì «·ﬂ«› —Ì«
If ssxx = "nocaftr" Then
MsgBox "·« ÌÊÃœ √Ì √’‰«› ›Ì «·ﬂ«› —Ì«", 64, "‰Ÿ«„ «·„Œ «—"
End If

' Õ„Ì· «·√’‰«› «·„ÊÃÊœ… ›Ì «·ﬂ«› —Ì«
If ssxx = "caflist" Then
 Dim ac As Integer
 Dim firc As String
 Dim snowc As String
 For ac = 8 To Len(s)
 firc = Mid(s, ac, 1)
 If firc = "%" Then
   If snowc <> "" Then
    FrmCafterea.List1.AddItem snowc
    snowc = ""
   End If
 Else
   snowc = snowc & firc
 End If
 Next
 FrmCafterea.Show
 StayOnTop FrmCafterea
End If

'«·ﬂ„Ì… €Ì— ﬂ«›Ì…
If ssxx = "cafnota" Then
MsgBox "«·ﬂ„Ì¯… «·„ÿ·Ê»… €Ì— „ Ê›¯—… Õ«·Ì¯«", 16, "‰Ÿ«„ «·„Œ «—"
End If
'«⁄ –«— ⁄‰  ·»Ì… «·ÿ·»
If ssxx = "cafreqn" Then
MsgBox "«·„”ƒÊ· ⁄‰ «·ﬂ«› —Ì« Ì⁄ –— ⁄‰  ·»Ì… «·ÿ·»", 64, "‰Ÿ«„ «·„Œ «—"
End If
' „  ·»Ì… «·ÿ·»
If ssxx = "cafreok" Then
MsgBox " „ ≈—”«· «·ÿ·»Ì… Ê «·„Ê«›ﬁ… ⁄·ÌÂ« ”Ì „ «· ‰›Ì– ›Ì √ﬁ—» Êﬁ ", 64, "‰Ÿ«„ «·„Œ «—"
End If
End Sub
