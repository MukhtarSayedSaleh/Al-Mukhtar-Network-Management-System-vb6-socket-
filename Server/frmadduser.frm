VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmadduser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈÷«›… «‘ —«ﬂ ÃœÌœ"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   Icon            =   "frmadduser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   3840
      OleObjectBlob   =   "frmadduser.frx":29C12
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "frmadduser.frx":29C86
      Top             =   1080
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   3840
      OleObjectBlob   =   "frmadduser.frx":29EBA
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3840
      OleObjectBlob   =   "frmadduser.frx":29F2E
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   3720
      OleObjectBlob   =   "frmadduser.frx":29FB2
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "frmadduser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "«·—Ã«¡ «· Õﬁﬁ „‰ ≈œŒ«· Ã„Ì⁄ «·»Ì«‰«  «·„ÿ·Ê»…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If
If IsNumeric(Text3.Text) = False Or IsNumeric(Text4.Text) = False Then
MsgBox "ÌÃ» √‰  Õ ÊÌ ⁄œœ «·”«⁄«  «·ﬂ·Ì¯… ⁄·Ï √—ﬁ«„ ›ﬁÿ", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If
'≈÷«›… ”Ã· ÃœÌœ
If usrcommand = "Add" Then
With frmUsers
On Error Resume Next
.Data1.Recordset.AddNew
.Text1.Text = Text1.Text
.Text2.Text = Text2.Text
.Text3.Text = CInt(Text3.Text)
.Text4.Text = CInt(Text4.Text)
.Text5.Text = CInt(Text3.Text)
.Text6.Text = CInt(Text4.Text)
On Error Resume Next
.Data1.Recordset.MoveNext
On Error Resume Next
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
    FrmSetting.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
    FrmSetting.Data1.RecordSource = "select * from settings"
    FrmSetting.Data1.Refresh
    Dim xxwwq As Double
    xxwwq = Round(CDbl(Text4.Text) * CDbl(FrmSetting.Text2.Text), 2)
    Unload FrmSetting
    With Frmmoney
   .Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
   .Data1.RecordSource = "select * from in_out_price"
   .Data1.Refresh
   .Data1.Recordset.AddNew
   .Text3.Text = CBool(True)
   .Text4.Text = CDbl(xxwwq)
   .Text5.Text = " ﬁÌ„… ≈‘ —«ﬂ " & Text1.Text
   .Text6.Text = Date
    On Error Resume Next
    .Data1.Recordset.MoveNext
    .Data1.Recordset.MovePrevious
    .Data1.Refresh
    End With
    Unload Frmmoney
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = "0"
Unload Me
End If

' ⁄œÌ· «·”Ã·
If usrcommand = "Edit" Then
With frmUsers
On Error Resume Next
.Data1.Recordset.Edit
.Text1.Text = Text1.Text
.Text2.Text = Text2.Text
.Text3.Text = CInt(Text3.Text)
.Text4.Text = CInt(Text4.Text)
.Text5.Text = .Text5.Text
.Text6.Text = .Text6.Text
On Error Resume Next
.Data1.Recordset.MoveNext
On Error Resume Next
.Data1.Recordset.MovePrevious
.DBGrid1.Refresh
.DBGrid1.ReBind
End With
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = "0"
Unload Me

End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

If usrcommand = "Edit" Then
Text1.Text = frmUsers.Text1.Text
Text2.Text = frmUsers.Text2.Text
Text3.Text = frmUsers.Text3.Text
Text4.Text = frmUsers.Text4.Text
Else
Text1.Text = ""
Text2.Text = ""
Text3.Text = "0"
Text4.Text = "0"
End If
End Sub

Private Sub Text3_Change()
If IsNumeric(Text3.Text) = True Then
Text4.Text = Int(Text3.Text) * 60
Else
MsgBox "«·—Ã«¡ ≈œŒ«· √—ﬁ«„ ›ﬁÿ", 16, "‰Ÿ«„ «·„Œ «—"
Text3.Text = "0"
End If

End Sub
