VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„Ê«÷Ì⁄ «· ⁄·Ì„« "
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   Icon            =   "J.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "≈€·«ﬁ"
      Height          =   375
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   10320
      Width           =   1575
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   1560
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   2400
      Width           =   1200
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10680
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   15240
      ExtentX         =   26882
      ExtentY         =   18838
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "help"
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
Me.WindowState = 2
With WebBrowser1
.Left = 10
.Top = 10
.Height = Me.Height - 770
.Width = Me.Width - 10
.Navigate App.Path & ("\help\help.mok")

End With
End Sub

Private Sub Form_Resize()
With WebBrowser1
.Left = 10
.Top = 10
.Height = Me.Height - 770
.Width = Me.Width - 10
.Navigate App.Path & ("\help\help.mok")

End With

End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

End Sub
