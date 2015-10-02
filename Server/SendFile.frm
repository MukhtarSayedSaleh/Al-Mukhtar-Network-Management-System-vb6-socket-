VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmSendFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈—”«· „·›"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   Icon            =   "SendFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "SendFile.frx":29C12
      TabIndex        =   6
      Top             =   120
      Width           =   6855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3240
      OleObjectBlob   =   "SendFile.frx":29C88
      Top             =   600
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "≈€·«ﬁ"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ‘€Ì· «·„·› ›Ì «·ÃÂ«“ «·Âœ›"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "»œ¡ «·≈—”«·"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«” ⁄—«÷"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "C:\abokamal.txt"
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sorcefile As String
Dim issending As Boolean


Private Sub Command1_Click()
cd1.Filter = "All Files |*.*|"
cd1.ShowOpen
sorcefile = cd1.FileName
Text1.Text = sorcefile
End Sub

Private Sub Command2_Click()
'≈—”«· «·„·›
Dim cmd As String
Dim ss As String
Dim msg As String
Dim f As Integer
Dim p As Long
Const Size = 2048
f = FreeFile
Open sorcefile For Binary As f
bar1.Max = LOF(f)
bar1.Min = 0
bar1.Value = 0
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
If frm_main.wsk(nowcompnum).State = 7 Then
issending = True
canbutton
frm_main.wsk(nowcompnum).SendData "filesou"
End If
For p = 1 To LOF(f) \ Size
ss = Space(Size)
Get #f, , ss
If issending = True Then
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "newfart" & ss
canbutton
bar1.Value = bar1.Value + Len(ss)
Else
GoTo bye
End If
DoEvents
End If
Next
If LOF(f) Mod Size > 0 Then
ss = Space(LOF(f) Mod Size)
If issending = True Then
If frm_main.wsk(nowcompnum).State = 7 Then
Get #f, , ss
frm_main.wsk(nowcompnum).SendData "newfart" & ss
bar1.Value = bar1.Value + Len(ss)
canbutton
DoEvents
Else
GoTo bye
End If
End If
End If
If issending = True Then
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "enffile" & Text1.Text
bar1.Value = LOF(f)
issending = False
canbutton
DoEvents
MsgBox " „ ≈—”«· «·„·› »‰Ã«Õ", 64, "·„⁄·Ê„« ﬂ"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
End If
End If
bye:
Close f
End Sub

Private Sub Command3_Click()
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "runfile" & Text1.Text
End If
End Sub

Private Sub Command4_Click()
If issending = True Then
If frm_main.wsk(nowcompnum).State = 7 Then
  frm_main.wsk(nowcompnum).SendData "stpsend"
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = False
MsgBox "›‘· ≈—”«· «·„·› »”»» ≈Ìﬁ«›  ‰›Ì– «·⁄„·Ì¯…", 16, "Œÿ« √À‰«¡ ‰ﬁ· «·„·›"
issending = False
bar1.Value = 0
canbutton
End If
Else
Unload Me
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

issending = False
bar1.Value = 0
End Sub

Public Sub canbutton()
If issending = True Then
Command4.Caption = "≈·€«¡ «·≈—”«·"
Else
Command4.Caption = "«€·«ﬁ"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If issending = True Then
If frm_main.wsk(nowcompnum).State = 7 Then
  frm_main.wsk(nowcompnum).SendData "stpsend"
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = False
MsgBox "›‘· ≈—”«· «·„·› »”»» ≈Ìﬁ«›  ‰›Ì– «·⁄„·Ì¯…", 16, "Œÿ« √À‰«¡ ‰ﬁ· «·„·›"
issending = False
bar1.Value = 0
canbutton
End If
Else
Unload Me
End If

End Sub
