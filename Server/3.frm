VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«⁄œ«œ«  «·»—‰«„Ã"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   Icon            =   "3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   5175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Frame Frame4 
         Height          =   2415
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   4575
         Begin VB.CommandButton Command3 
            Caption         =   " €ÌÌ— ﬂ·„… «·„—Ê—"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1920
            Width           =   4335
         End
         Begin VB.PictureBox SkinLabel6 
            Height          =   255
            Left            =   2640
            ScaleHeight     =   195
            ScaleWidth      =   1755
            TabIndex        =   25
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1440
            Width           =   2535
         End
         Begin VB.PictureBox SkinLabel5 
            Height          =   255
            Left            =   2640
            ScaleHeight     =   195
            ScaleWidth      =   1755
            TabIndex        =   23
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   960
            Width           =   2535
         End
         Begin VB.PictureBox SkinLabel4 
            Height          =   255
            Left            =   2760
            ScaleHeight     =   195
            ScaleWidth      =   1635
            TabIndex        =   21
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   480
            Width           =   2535
         End
         Begin VB.PictureBox SkinLabel3 
            Height          =   255
            Left            =   120
            ScaleHeight     =   195
            ScaleWidth      =   4275
            TabIndex        =   19
            Top             =   120
            Width           =   4335
         End
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      DataField       =   "statebr"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      DataField       =   "toolbr"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "run_startup"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   2.45745e5
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "minute_price"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Text            =   "Text4"
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
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "settings"
      RightToLeft     =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   7575
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "hour_price"
         DataSource      =   "Data1"
         Height          =   495
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   2.45745e5
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "                    ⁄—÷ ‘—Ìÿ «·Õ«·…"
         Height          =   435
         Left            =   5040
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "                  ⁄—÷ ‘—Ìÿ «·√œÊ« "
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "   ‘€Ì·  ·ﬁ«∆Ì ⁄‰œ › Õ «·‰Ÿ«„"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   3600
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   27
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«·€«¡ «·√„—"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„‹‹‹‹‹‹‹‹Ê«›‹‹‹‹‹‹‹‹‹ﬁ"
      Height          =   375
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   7575
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   2535
      End
      Begin VB.PictureBox SkinLabel1 
         Height          =   255
         Left            =   6480
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   2535
      End
      Begin VB.PictureBox SkinLabel2 
         Height          =   255
         Left            =   6480
         ScaleHeight     =   195
         ScaleWidth      =   915
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "«⁄œ«œ«  ⁄«„…"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "«⁄œ«œ«  „«·Ì¯…"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "≈⁄œ«œ«  «·Õ„«Ì…"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmSetting"
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
If Text1.Text = "" Or Text1.Text = "0" Or Text2.Text = "" Then
MsgBox "«·—Ã«¡ ≈œŒ«· ﬁÌ„… ”⁄— «·”«⁄…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If
If IsNumeric(Text1.Text) = False Then
MsgBox "≈‰ «· ⁄»Ì— «·–Ì ›Ì Õﬁ· ”⁄— «·”«⁄… ·Ì” —ﬁ„«", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

Data1.Recordset.Edit
Text3.Text = Text1.Text
Text4.Text = Text2.Text
Text5.Text = CBool(Check1.Value)
Text6.Text = CBool(Check2.Value)
Text7.Text = CBool(Check3.Value)
On Error Resume Next
Data1.Recordset.MoveNext
On Error Resume Next
Data1.Recordset.MovePrevious
Data1.Refresh
Unload Me
'·«“„ ‘€·…  ⁄„·  ÕœÌÀ  ·ﬁ«∆Ì ⁄‰œ  €ÌÌ— «·≈⁄œ«œ«  »‰«¡ ⁄·Ï «·≈⁄œ«œ«  «·ÃœÌœ…
Call frm_main.refreshsetting
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
If Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "«·—Ã«¡ «· √ﬂœ „‰  ⁄»∆… Ã„Ì⁄ «·ÕﬁÊ· «·„ÿ·Ê»…", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

If Text9.Text <> Text10.Text Then
MsgBox "ﬂ·„… «·„—Ê— «·ÃœÌœ… Ê  √ﬂÌœÂ« €Ì— „ ÿ«»ﬁÌ‰", 16, "‰Ÿ«„ «·„Œ «—"
Exit Sub
End If

Dim vadmin As String
Dim vpass As String
vadmin = "Admin"
vpass = Trim(Text8.Text)
With signin
.Data1.RecordSource = "select * from admin where aname='" & vadmin & "'"
.Data1.Refresh
If .txtaname.Text = vadmin Then
   If .txtapass.Text = vpass Then
     On Error Resume Next
    .Data1.Recordset.Edit
    .txtapass.Text = Trim(Text10.Text)
    On Error Resume Next
    .Data1.Recordset.MoveNext
    .Data1.Recordset.MovePrevious
    MsgBox " „  €ÌÌ— ﬂ·„… «·„—Ê— »‰Ã«Õ", 64, "‰Ÿ«„ «·„Œ «—"
   Else
    MsgBox "ﬂ·„… «·„—Ê— «·ﬁœÌ„… Œ«ÿ∆…", 16, "›‘·  ”ÃÌ· «·œŒÊ·"
   End If
End If
End With
On Error Resume Next
Unload signin
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd
Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
Data1.RecordSource = "select * from settings"
Data1.Refresh
refill
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(1).Selected = True Then
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
ElseIf TabStrip1.Tabs(2).Selected = True Then
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
ElseIf TabStrip1.Tabs(3).Selected = True Then
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
End If

End Sub

Public Function refill()
If Text3.Text = "" Then
Exit Function
End If

Text1.Text = Text3.Text
Text2.Text = Text4.Text
If CBool(Text5.Text) = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If CBool(Text6.Text) = True Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If CBool(Text7.Text) = True Then
Check3.Value = 1
Else
Check3.Value = 0
End If

End Function

Private Sub Text1_Change()
If IsNumeric(Text1.Text) = True Then
Text2.Text = Round(CDbl(Text1.Text) * 60 ^ -1, 2)
End If
End Sub

