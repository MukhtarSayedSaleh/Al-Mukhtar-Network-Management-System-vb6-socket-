VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form pmmsg 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "—”«·… Œ«’… „‰ „œÌ— «·‘»ﬂ…"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   ControlBox      =   0   'False
   Icon            =   "pmmsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "pmmsg.frx":29C12
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "„Ê«›ﬁ"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00004040&
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "pmmsg.frx":29E46
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "pmmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
'«·”ﬂ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin Me.hwnd

StayOnTop Me
End Sub
