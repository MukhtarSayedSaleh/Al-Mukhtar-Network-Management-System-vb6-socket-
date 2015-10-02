VERSION 5.00
Begin VB.Form Taskman 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁ«∆„… «· ÿ»Ìﬁ« "
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   ".frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   3240
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   2640
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈‰Â«¡ Ã„Ì⁄ «·„Â«„"
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈‰Â«¡ «·„Â„…"
      Height          =   375
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
End
Attribute VB_Name = "Taskman"
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
Dim X As Integer
For X = 0 To List1.ListCount - 1
If List1.Selected(X) = True Then
Exit For
End If
Next
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "endtask" & X
Unload Me

End If
End Sub

Private Sub Command2_Click()
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "endalltasks"
End If
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub
