VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form reqmsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÿ·» › Õ «·ÃÂ«“"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "reqmsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "reqmsg.frx":29C12
         TabIndex        =   2
         Top             =   240
         Width           =   5895
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "reqmsg.frx":29C69
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„‹‹‹‹Ê«›‹‹‹ﬁ"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   6135
   End
End
Attribute VB_Name = "reqmsg"
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

StayOnTop reqmsg
End Sub

