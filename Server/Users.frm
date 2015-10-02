VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„‘ —ﬂÌ‰"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4080
      OleObjectBlob   =   "Users.frx":29C12
      Top             =   1320
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Õ–› «‘ —«ﬂ"
      Height          =   375
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ⁄œÌ· «‘ —«ﬂ"
      Height          =   375
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈÷«›… «‘ —«ﬂ"
      Height          =   375
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "≈€·«ﬁ"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "minutesp"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   12522
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "hoursp"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   12522
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "allminutes"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   12522
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "allhours"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   12522
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      DataField       =   "passwordd"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   12522
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "uname"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   12522
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\almokhtar network\Server\mokdatabase.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      RightToLeft     =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Users.frx":29E46
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "Users.frx":29E5A
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
usrcommand = "Add"
frmadduser.Show
frmadduser.Caption = "«÷«›… «‘ —«ﬂ ÃœÌœ"
StayOnTop frmadduser
End Sub

Private Sub Command3_Click()
usrcommand = "Edit"
frmadduser.Show
frmadduser.Caption = " ⁄œÌ· »Ì«‰«  «·„‘ —ﬂ"
StayOnTop frmadduser
End Sub

Private Sub Command4_Click()
Dim mok
mok = MsgBox(" Â· √‰  „ √ﬂœ √‰ﬂ  —Ìœ Õ–› «‘ —«ﬂ " & Text1.Text, 64 + vbYesNo, "‰Ÿ«„ «·„Œ «—")
If mok = vbYes Then
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
If Text1.Text = "" Then
Command4.Enabled = False
Command3.Enabled = False
Else
Command4.Enabled = True
Command3.Enabled = True
End If
Else
Exit Sub
End If
End Sub

Private Sub DBGrid1_Click()
If Text1.Text = "" Then
Command4.Enabled = False
Command3.Enabled = False
Else
Command4.Enabled = True
Command3.Enabled = True
End If
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
Data1.RecordSource = "select * from users"
Data1.Refresh
If Text1.Text = "" Then
Command4.Enabled = False
Command3.Enabled = False
Else
Command4.Enabled = True
Command3.Enabled = True
End If

End Sub
