VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Jii´s Window Manager"
   ClientHeight    =   5370
   ClientLeft      =   3540
   ClientTop       =   3675
   ClientWidth     =   7545
   ForeColor       =   &H00000000&
   Icon            =   "taskmngr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7545
   Begin VB.CommandButton Command2 
      Caption         =   "Registry Tweaks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Disable shutdown options when locked"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Disable Sytem-keys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lock Computer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   25
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Only show visible windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   3240
      TabIndex        =   23
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   22
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdDontstayontop 
      Caption         =   "Don´t stay on top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdStayontop 
      Caption         =   "Stay on top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdDestroy 
      Caption         =   "Destroy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdMaximize 
      Caption         =   "Maximize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Type in the new Titlebar text for the selected window and hit ""Set Text"""
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton cmdSettext 
      Caption         =   "Set Text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   14
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   8
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   7335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7200
      Top             =   2520
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stay on top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      ToolTipText     =   "Settings for this program"
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6840
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      Height          =   135
      Left            =   0
      Top             =   5160
      Width           =   7695
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7440
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7440
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "<------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Handle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'* Al-Mokhtar For Networks Client       *
'*  By : Mokhtar saied saleh            *
'*      Syria - Abokamal                *
'*      WWW.ABOKAMAL.COM                *
'*  MOKHTAR_SS@HOTMAIL.COM              *
'*       0096394467547                  *
'****************************************




Private Sub Check2_Click()
cmdRefresh_Click
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, 0, 0)
Else
Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, 0, 0)
End If

End Sub

Private Sub cmdRefresh_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
List1.Clear

svar = EnumWindows(AddressOf getalltopwindows, 0)

End Sub

Private Sub cmdMaximize_Click()
On Error Resume Next
svar = ShowWindow(Text2.Text, 3)

Me.SetFocus
End Sub

Private Sub cmdHide_Click()
On Error Resume Next

svar = ShowWindow(Text2.Text, 0)
End Sub

Private Sub cmdShow_Click()
On Error Resume Next
svar = ShowWindow(Text2.Text, 5)

Me.SetFocus
End Sub

Private Sub cmdSettext_Click()
On Error Resume Next
If Text1.Text = "" Then svar = MsgBox("Choose a task", vbInformation, "Error"): Text3.SetFocus: Exit Sub
svar = SetWindowText(Text2.Text, Text3.Text)
cmdRefresh_Click
End Sub

Private Sub cmdMinimize_Click()
On Error Resume Next
svar = ShowWindow(Text2.Text, 6)

Me.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End
End Sub

Private Sub cmdRestore_Click()
On Error Resume Next
svar = ShowWindow(Text2.Text, 9)

Me.SetFocus
End Sub

Private Sub cmdDestroy_Click()
On Error Resume Next

svar = GetWindowThreadProcessId(Text2.Text, nyprocessid)

procname = OpenProcess(PROCESS_ALL_ACCESS, 0&, nyprocessid)

svar2 = TerminateProcess(procname, 0&)

DoEvents
cmdRefresh_Click
End Sub






Private Sub cmdEnable_Click()
On Error Resume Next
svar = EnableWindow(Text2.Text, 1)

End Sub

Private Sub cmdDisable_Click()
On Error Resume Next
svar = EnableWindow(Text2.Text, 0)
End Sub



Private Sub Command1_Click()

Dim newpass As String

' hide taskbar
For x = 0 To List2.ListCount - 1
If List2.List(x) = "Taskbar" Then
tmptaskbarhwnd = List2.ItemData(x)
svar = ShowWindow(List2.ItemData(x), 0)
End If
Next x

'disable systemkeys
svar = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, 0, 0)

form1.Hide
Form4.BackColor = &H80000001
Form4.Show

If form1.Check4.Value = 1 Then
Form4.Command4.Enabled = False
Form4.Command5.Enabled = False
End If

On Error GoTo errhandler
Open App.Path & "\pw.fil" For Input As #1
Close #1

Exit Sub

errhandler:

Do While newpass = ""
newpass = InputBox("Type in your admin password:", "Set password")
Loop

tmppass2 = ""

For x = 1 To Len(newpass)
tmppass2 = tmppass2 & Asc(Mid(newpass, x, 1)) & " "
Next x

Open App.Path & "\pw.fil" For Output As #1
Print #1, tmppass2
Close #1

svar = MsgBox("Your password has now been changed!", vbOKOnly, "Jii´s Window Manager")

End Sub

Private Sub Command2_Click()
form1.Hide
frmreg.Show
End Sub

Private Sub Form_Load()
cmdRefresh_Click

taskhwnd = FindWindow("Shell_TrayWnd", vbNullString)
form1.List2.AddItem "Taskbar"
form1.List2.ItemData(form1.List2.NewIndex) = taskhwnd
starthwnd = FindWindowEx(taskhwnd, 0&, "Button", vbNullString)
form1.List2.AddItem "Start Button"
form1.List2.ItemData(form1.List2.NewIndex) = starthwnd
systrayhwnd = FindWindowEx(taskhwnd, 0&, "TrayNotifywnd", vbNullString)
form1.List2.AddItem "System Tray"
form1.List2.ItemData(form1.List2.NewIndex) = systrayhwnd
clockhwnd = FindWindowEx(systrayhwnd, 0&, "TrayClockWClass", vbNullString)
form1.List2.AddItem "System Clock"
form1.List2.ItemData(form1.List2.NewIndex) = clockhwnd

tasklisthwnd = FindWindowEx(FindWindowEx(taskhwnd, 0&, "ReBarWindow32", vbNullString), 0&, "MSTaskSwWClass", vbNullString)

form1.List2.AddItem "Task List"
form1.List2.ItemData(form1.List2.NewIndex) = tasklisthwnd

quicklaunchhwnd = FindWindowEx(FindWindowEx(taskhwnd, 0&, "ReBarWindow32", vbNullString), 0&, "ToolBarWindow32", vbNullString)

form1.List2.AddItem "Quick Launch"
form1.List2.ItemData(form1.List2.NewIndex) = quicklaunchhwnd



End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Image1_Click()
On Error Resume Next

svar = MsgBox("Made by: " & Chr(74) & Chr(111) & Chr(110) & Chr(97) & Chr(116) & Chr(97) & Chr(110) & vbCrLf & "jonatan.dahl@telia.com" & vbCrLf & "All Rights Reserved, Jonatan Dahl, 2001", vbInformation, "About")
SHELL ("start readme.txt"), vbHide
svar = ShowWindow(form1.hwnd, 6)
End Sub

Private Sub Label4_Click()
On Error Resume Next

svar = MsgBox("Made by: Jii" & vbCrLf & "jonatan.dahl@telia.com", vbInformation, "About")
SHELL ("start readme.txt"), vbHide
svar = ShowWindow(form1.hwnd, 6)
End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub List1_Click()
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
Text1.Text = List1.List(x)
Text2.Text = List1.ItemData(x)
GoTo seeifresponse
End If
Next x

seeifresponse:
tempsvar = SendMessageTimeout(Text2.Text, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, 1000, WM_NULL)
If tempsvar > 0 Then
lblStatus.Caption = "Status: Running"
Else
lblStatus.Caption = "Status: Doesn´t respond"
End If

End Sub

Private Sub List1_LostFocus()
lblStatus.Caption = ""
End Sub

Private Sub List2_Click()

For x = 0 To List2.ListCount - 1
If List2.Selected(x) = True Then
Text1.Text = List2.List(x)
Text2.Text = List2.ItemData(x)
Exit Sub
End If
Next x

End Sub

Private Sub Text1_Change()
Text3.Text = Text1.Text
End Sub

Private Sub Timer1_Timer()
Dim foregroundwindow As Long
Dim textlen As Long
Dim windowtext As String
Dim svar As Long
Static lastwindowtext As String



foregroundwindow = GetForegroundWindow()

textlen = GetWindowTextLength(foregroundwindow) + 1

windowtext = Space(textlen)
svar = GetWindowText(foregroundwindow, windowtext, textlen)
windowtext = Left(windowtext, Len(windowtext) - 1)

If windowtext = "" Then Exit Sub

If Not windowtext = lastwindowtext Then


List1.AddItem windowtext
List1.ItemData(List1.NewIndex) = foregroundwindow
End If

lastwindowtext = windowtext
End Sub
