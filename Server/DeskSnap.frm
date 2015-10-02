VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DeskSnap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’Ê—… ”ÿÕ «·„ﬂ »"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "DeskSnap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "≈€·«ﬁ"
      Height          =   855
      Left            =   0
      Picture         =   "DeskSnap.frx":29C12
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Õ›Ÿ «·’Ê—…"
      Height          =   855
      Left            =   840
      Picture         =   "DeskSnap.frx":2A054
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "DeskSnap.frx":2A496
      Top             =   1320
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox image1 
      Height          =   495
      Left            =   1800
      RightToLeft     =   -1  'True
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "DeskSnap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
cd1.DialogTitle = DeskSnap.Caption
cd1.Filter = "BitMap Pictures" & "|*.bmp|"
cd1.ShowSave
On Error Resume Next
SavePicture DeskSnap.image1.Picture, cd1.FileName
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub
