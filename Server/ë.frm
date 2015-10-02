VERSION 5.00
Begin VB.Form cafadd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«÷«›… ’‰› ÃœÌœ"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   Icon            =   "¯.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   2280
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   9
      Top             =   960
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   4815
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "≈·€«¡ «·√„—"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   160
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "„Ê«›ﬁ"
         Default         =   -1  'True
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   160
         Width           =   2415
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   3495
   End
   Begin VB.PictureBox SkinLabel1 
      Height          =   255
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.PictureBox SkinLabel2 
      Height          =   255
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox SkinLabel3 
      Height          =   255
      Left            =   3720
      ScaleHeight     =   195
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "cafadd"
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "«·—Ã«¡ «· Õﬁﬁ „‰ ≈œŒ«· Ã„Ì⁄ «·»Ì«‰«  «·„ÿ·Ê»…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If
If IsNumeric(Text2.Text) = False Or IsNumeric(Text3.Text) = False Then
MsgBox "«·—Ã«¡ «· Õﬁﬁ „‰ «·„œŒ·«  «·—ﬁ„Ì…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If
If usrcommand2 = "Add" Then
With FrmCaftrea
On Error Resume Next
.Data1.Recordset.AddNew
.Text1.Text = Text1.Text
.Text2.Text = CInt(Text2.Text)
.Text3.Text = CInt(Text3.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
On Error Resume Next
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Text1.Text = ""
Text2.Text = "0"
Text3.Text = "0"
Unload Me
End If

' ⁄œÌ· «·”Ã·
If usrcommand2 = "Edit" Then
With FrmCaftrea
On Error Resume Next
.Data1.Recordset.Edit
.Text1.Text = Text1.Text
.Text2.Text = CInt(Text2.Text)
.Text3.Text = CInt(Text3.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
On Error Resume Next
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Text1.Text = ""
Text2.Text = "0"
Text3.Text = "0"
Unload Me
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

If usrcommand2 = "Edit" Then
Text1.Text = FrmCaftrea.Text1.Text
Text2.Text = FrmCaftrea.Text2.Text
Text3.Text = FrmCaftrea.Text3.Text
Else
Text1.Text = ""
Text2.Text = "0"
Text3.Text = "0"
End If

End Sub

