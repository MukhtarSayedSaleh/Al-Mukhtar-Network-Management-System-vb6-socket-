VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_closed 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   -135
   ClientTop       =   -135
   ClientWidth     =   15360
   Icon            =   "frm_closed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7440
      OleObjectBlob   =   "frm_closed.frx":29C12
      Top             =   5520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hide For Programmer Only"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   10920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   2160
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   " ”ÃÌ· «·œŒÊ· - ··„‘ —ﬂÌ‰"
      Height          =   375
      Left            =   11520
      TabIndex        =   5
      Top             =   10560
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÿ·» › Õ «·ÃÂ«“ „‰ «·≈œ«—…"
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   11040
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   9600
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   11040
      Visible         =   0   'False
      Width           =   975
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.TextBox ipa 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash sw1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _cx             =   4215285
      _cy             =   4200680
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
      Stacking        =   "below"
   End
End
Attribute VB_Name = "frm_closed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If frm_main.Winsock1.State = 7 Then
frm_main.Winsock1.SendData "[openrequest]"
End If
End Sub

Private Sub Command2_Click()
signin.Show
End Sub

Private Sub Command3_Click()
frm_closed.Hide
End Sub

Private Sub Form_Load()
'«·„·› «·”ﬂÌ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin Command1.hwnd
Skin1.ApplySkin Command2.hwnd
' Õ„Ì· —ﬁ„ «·¬Ì »Ì „‰ «·„·›
Dim s As String, m As String
s = App.Path & ("\mokip.dll")
If FileExists(s) = True Then
Open s For Input As #1
m = Input(LOF(1), 1)
ipa.Text = Mid(m, 4, Len(m) - 6)
Close #1
End If
Rem «·≈ ’«· ⁄·Ï «·„‰›– Ê «·¬Ì »Ì «·„Õœœ
frm_main.Winsock1.Close
frm_main.Winsock1.Connect ipa.Text, 13770
' Õ„Ì· „·› «·›·«‘
sw1.Width = frm_closed.Width
sw1.Height = frm_closed.Width
sw1.Top = 0
sw1.Left = 0
sw1.Movie = App.Path & ("\Closedfl.dll")
'„‰⁄ «·“— «·Ì„Ì‰

'œÊ„« ›Ì «·„ﬁœ„…
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
frmlogin.Show
Else
Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
Label1.Caption = frm_main.Winsock1.State
End Sub
