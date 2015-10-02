VERSION 5.00
Begin VB.Form signin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «·œŒÊ·"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "signin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   2400
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   600
      Width           =   1200
   End
   Begin VB.PictureBox SkinLabel1 
      Height          =   255
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Label1 
      Height          =   255
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "signin"
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
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "«·—Ã«¡ «· Õﬁﬁ „‰ „‰  ⁄»∆… ﬂ«›… «·ÕﬁÊ· «·„ÿ·Ê»…", 16, "Œÿ√"
Exit Sub
End If
If frm_main.Winsock1.State = 7 Then
frm_main.Winsock1.SendData "opncomp" & Trim(Text1.Text) & "%%" & Trim(Text2.Text)
nowuser = Trim(Text1.Text)
End If

If frm_main.Winsock1.State = 9 Then
If Text1.Text = "„Œ «—" And Text2.Text = "”Ìœ ’«·Õ" Then
Unload frm_closed
frm_main.Show
frm_main.Label2.Caption = "„” Œœ„ ’Ì«‰… «·‰Ÿ«„"
frm_main.Command1.Enabled = True
frm_main.Command2.Enabled = True
frm_main.Command3.Enabled = True
frm_main.Command4.Enabled = True
frm_main.Command5.Enabled = True
Unload Me

End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'«·”ﬂ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin Me.hwnd

StayOnTop Me
End Sub

Private Sub label1_Click()

End Sub
