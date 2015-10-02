VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÕÊ· ‰Ÿ«„ «·„Œ «—"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   Icon            =   "Ï.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Height          =   495
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8100
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   45
         Picture         =   "Ï.frx":29C12
         RightToLeft     =   -1  'True
         ScaleHeight     =   2775
         ScaleWidth      =   7935
         TabIndex        =   1
         Top             =   120
         Width           =   7935
      End
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   3840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   3
      Top             =   1560
      Width           =   1200
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************
'* Al-Mokhtar For Networks Server       *
'*  By : Mokhtar saied saleh            *
'*      Syria - Abokamal                *
'*      WWW.ABOKAMAL.COM                *
'*  MOKHTAR_SS@HOTMAIL.COM              *
'*       0096394467547                  *
'****************************************
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub
