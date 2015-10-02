VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Frmmoney 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·√—»«Õ Ê «·„’—Ê›« "
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "Frmmoney.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "date"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "why"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "amount"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "type"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\almokhtar network\Server\mokdatabase.dll"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "in_out_price"
      RightToLeft     =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "Frmmoney.frx":29C12
      Top             =   2400
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   4395
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Frmmoney.frx":29E46
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "Frmmoney.frx":29E5A
      TabIndex        =   8
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   2295
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "„—»Õ"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   200
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "„’—Ê›"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   200
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "Frmmoney.frx":2ABA1
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   5415
      Begin VB.CommandButton Command4 
         Cancel          =   -1  'True
         Caption         =   "&Œ‹—ÊÃ"
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   170
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "≈÷«›…"
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   170
         Width           =   3375
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   2000
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "Frmmoney.frx":2AC1B
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Frmmoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = False And Option2.Value = False Then
MsgBox "«·—Ã«¡  ÕœÌœ ‰Ê⁄ «·ﬁÌ„… «·„«·Ì… Â· ÂÌ „—»Õ √„ „’—Ê›", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

If Text1.Text = "" Or Text1.Text = "0" Then
MsgBox "«·—Ã«¡ ≈œŒ«· «·ﬁÌ„… «·„«·Ì…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "«·—Ã«¡ ≈œŒ«· «·”»»", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

If IsNumeric(Text1.Text) = False Then
MsgBox "≈‰¯ «· ⁄»Ì— «·–Ì ›Ì Õﬁ· «·ﬁÌ„… «·„«·Ì… ·Ì” —ﬁ„« «·—Ã«¡  ’ÕÌÕÂ", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

Data1.Recordset.AddNew
If Option1.Value = True Then
Text3.Text = CBool(Option1.Value)
End If
If Option2.Value = True Then
Text3.Text = CBool(Option1.Value)
End If
Text4.Text = CDbl(Text1.Text)
Text5.Text = Text2.Text
Text6.Text = Date
On Error Resume Next
Data1.Recordset.MoveNext
Data1.Recordset.MovePrevious
Data1.Refresh
DBGrid1.ReBind
DBGrid1.Refresh
Text1.Text = "0"
Text2.Text = ""
Option1.Value = False
Option2.Value = False

End Sub



Private Sub Command4_Click()
Unload Me

End Sub




Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd
Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
Data1.RecordSource = "select * from in_out_price"
Data1.Refresh

End Sub

