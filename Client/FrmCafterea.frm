VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCafterea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·ﬂ«› —Ì«"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   ControlBox      =   0   'False
   Icon            =   "FrmCafterea.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   3000
      OleObjectBlob   =   "FrmCafterea.frx":29C12
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   5160
      OleObjectBlob   =   "FrmCafterea.frx":29C92
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   4200
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3120
      OleObjectBlob   =   "FrmCafterea.frx":29D14
      Top             =   2280
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4680
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈—”«· «·ÿ·»"
      Height          =   375
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4680
      Width           =   3135
   End
End
Attribute VB_Name = "FrmCafterea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.Text = "" Or Text1.Text = "" Then
MsgBox "«·—Ã«¡ ≈œŒ«· ﬂ«›… «·»Ì«‰«  «·„ÿ·Ê»… ·≈ „«„ «·⁄„·Ì…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

If Text1.Text = "0" Or IsNumeric(Text1.Text) = False Then
MsgBox "«·—Ã«¡ ≈œŒ«· —ﬁ„ ’ÕÌÕ", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

Dim msg As String
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
 msg = List1.List(i)
End If
Next

If frm_main.Winsock1.State = 7 Then
frm_main.Winsock1.SendData "reqcaft" & msg & "%%" & Int(Trim(Text1.Text))
Unload Me
End If


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'«·”ﬂ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin FrmCafterea.hwnd
End Sub
