VERSION 5.00
Begin VB.Form FrmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈⁄œ«œ«  «·‰Ÿ«„"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "‚.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "≈€·«ﬁ"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   5535
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "  ‘€Ì· «·‰Ÿ«„  ·ﬁ«∆Ì¯« ⁄‰œ ﬂ· ≈ﬁ·«⁄"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Õ›Ÿ «· €ÌÌ—« "
         Height          =   375
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   3255
      End
      Begin VB.PictureBox SkinLabel1 
         Height          =   255
         Left            =   4080
         ScaleHeight     =   195
         ScaleWidth      =   1275
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   2760
      Width           =   1200
   End
End
Attribute VB_Name = "FrmSetting"
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
Dim file As String
file = App.Path & ("\mokip.dll")
If FileExists(file) = True Then
On Error Resume Next
Kill file
End If

Open file For Output As #1
Write #1, (LOF(1)), Trim(Text1.Text)
Close #1
MsgBox " „  €ÌÌ— «·—ﬁ„ »‰Ã«Õ ·‰  ’»Õ «·≈⁄œ«œ«  «·ÃœÌœ… ‰«›–… «·„›⁄Ê· ≈·« ›Ì «· ‘€Ì· «·ﬁ«œ„", 64, "‰Ÿ«„ «·„Œ «—"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'«·”ﬂ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin Me.hwnd


End Sub
