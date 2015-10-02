VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6FF9A514-A943-11D2-8D43-F90F0D71B6F6}#1.0#0"; "ChangeRes.ocx"
Begin VB.Form signin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «·œŒÊ·"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "signin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1635
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin ChangeResProject.ChangeRes ChangeRes1 
      Left            =   360
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1085
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   3840
      OleObjectBlob   =   "signin.frx":29C12
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   3840
      OleObjectBlob   =   "signin.frx":29C86
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "signin.frx":29CFC
      Top             =   600
   End
   Begin VB.TextBox txtapass 
      Alignment       =   1  'Right Justify
      DataField       =   "apass"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   12255
      Width           =   375
   End
   Begin VB.TextBox txtaname 
      Alignment       =   1  'Right Justify
      DataField       =   "aname"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   12255
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      RightToLeft     =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   375
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "signin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "«·—Ã«¡ «· √ﬂœ „‰  ⁄»∆… ﬂ«›… «·ÕﬁÊ· «·„ÿ·Ê»…", 16, "Œÿ√ ≈œŒ«·"
Exit Sub
End If

Dim vadmin As String
Dim vpass As String
vadmin = Trim(Text1.Text)
vpass = Trim(Text2.Text)
Data1.RecordSource = "select * from admin where aname='" & vadmin & "'"
Data1.Refresh
If txtaname.Text = vadmin Then
  If txtapass.Text = vpass Then
   frm_main.Show
   frm_main.SkinLabel6.Caption = Text1.Text
   Unload Me
  Else
   MsgBox "ﬂ·„… «·„—Ê— Œ«ÿ∆…", 16, "›‘·  ”ÃÌ· «·œŒÊ·"
  End If
Else
 MsgBox "«”„ «·„” Œœ„ Œ«ÿ∆", 16, "›‘·  ”ÃÌ· «·œŒÊ·"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With ChangeRes1
.Xpixels = 1024
.Ypixels = 768
.ChangeResolution = True
End With
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

'—»ÿ ﬁ«⁄œ… «·»Ì«‰«  »«·œ« « ”Ì 
Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
Data1.RecordSource = "select * from admin"
Data1.Refresh

End Sub
