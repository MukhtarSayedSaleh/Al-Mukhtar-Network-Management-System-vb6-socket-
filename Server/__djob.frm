VERSION 5.00
Begin VB.Form Endjob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈‰Â«¡ Ã·”… ⁄„·"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   ControlBox      =   0   'False
   Icon            =   "ˇˇdjob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "1"
      Top             =   4005
      Width           =   4095
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   "1"
      Top             =   4005
      Width           =   4095
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   4080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   2400
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "≈€·«ﬁ ›ﬁÿ"
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5760
      Width           =   8415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈€·«ﬁ „⁄  ”œÌœ «·Õ”«»"
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   4920
      Width           =   8415
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Text            =   "1"
      Top             =   2800
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Text            =   "1"
      Top             =   2800
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Text            =   "1"
      Top             =   1600
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "1"
      Top             =   1600
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   8535
   End
   Begin VB.PictureBox SkinLabel1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.PictureBox SkinLabel2 
      Height          =   375
      Left            =   4380
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
   End
   Begin VB.PictureBox SkinLabel3 
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.PictureBox SkinLabel4 
      Height          =   375
      Left            =   4380
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   2400
      Width           =   4095
   End
   Begin VB.PictureBox SkinLabel5 
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   2400
      Width           =   3975
   End
   Begin VB.PictureBox SkinLabel6 
      Height          =   375
      Left            =   4380
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   13
      Top             =   3600
      Width           =   3975
   End
   Begin VB.PictureBox SkinLabel7 
      Height          =   375
      Left            =   180
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   15
      Top             =   3600
      Width           =   3975
   End
End
Attribute VB_Name = "Endjob"
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
With Frmmoney
.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
.Data1.RecordSource = "select * from in_out_price"
.Data1.Refresh

.Data1.Recordset.AddNew
.Text3.Text = CBool(True)
.Text4.Text = CDbl(Text7.Text)
.Text5.Text = " Õ”«» «·ÃÂ«“ —ﬁ„ " & Text1.Text
.Text6.Text = Date
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious
.Data1.Refresh
End With
Unload Frmmoney
Unload Me

End Sub

Private Sub Command2_Click()
Dim m
m = MsgBox("Â· √‰  „ √ﬂ¯œ √‰ﬂ  —Ìœ ≈€·«ﬁ Â–Â «·Ê«ÃÂ… œÊ‰  ”œÌœ «·Õ”«» ø", 64 + vbYesNo, "‰Ÿ«„ «·„Œ «—")
If m = vbYes Then
Unload Me
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub

