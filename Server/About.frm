VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÕÊ· ‰Ÿ«„ «·„Œ «—"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   Icon            =   "About.frx":0000
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
         Picture         =   "About.frx":29C12
         RightToLeft     =   -1  'True
         ScaleHeight     =   2775
         ScaleWidth      =   7935
         TabIndex        =   1
         Top             =   120
         Width           =   7935
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3840
      OleObjectBlob   =   "About.frx":49BE6
      Top             =   1560
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub
