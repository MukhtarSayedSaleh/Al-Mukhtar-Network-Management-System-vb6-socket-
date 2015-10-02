VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmsendfileall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈—”«· „·› ·Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   "frmsendfileall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmsendfileall.frx":29C12
      TabIndex        =   6
      Top             =   120
      Width           =   6855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3240
      OleObjectBlob   =   "frmsendfileall.frx":29C88
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "C:\abokamal.txt"
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«” ⁄—«÷"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "»œ¡ «·≈—”«·"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ‘€Ì· «·„·› ›Ì ﬂ· «·‘»ﬂ…"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "≈€·«ﬁ"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmsendfileall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sorcefile As String
Dim issending As Boolean


Private Sub Command1_Click()
cd1.Filter = "All Files |*.*|"
cd1.ShowOpen
sorcefile = cd1.FileName
Text1.Text = sorcefile
End Sub

Private Sub Command2_Click()
' ⁄—Ì› «·„ €Ì—«  Ê «·ÀÊ«» 
Dim ss As String
Dim i As Integer
Dim f As Integer
Dim p As Long
Const Size = 2048
f = FreeFile
'› Õ «·„·› Ê „⁄—›… ÕÃ„Â Ê  ⁄ÿÌ· «·√“—«—
Open sorcefile For Binary As f
bar1.Max = LOF(f)
bar1.Min = 0
bar1.Value = 0
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
'«·»œ¡ »«·≈—”«· - √Ê¯·  ⁄·Ì„…

For i = 1 To 40
If frm_main.wsk(i).State = 7 Then
issending = True
canbutton
frm_main.wsk(i).SendData "filesou"
End If
Next

'«·»œ¡ »«·≈—”«· «·„·› - «·œ›⁄« 
For p = 1 To LOF(f) \ Size
 ss = Space(Size)
 Get #f, , ss
  If issending = True Then
    For i = 1 To 40
    If frm_main.wsk(i).State = 7 Then
     frm_main.wsk(i).SendData "newfart" & ss
     bar1.Value = bar1.Value + Len(ss)
    End If
    Next
     DoEvents
  End If
Next

'«—”«· «·„·› - «·„—Õ·… «·À«‰Ì…
If LOF(f) Mod Size > 0 Then
ss = Space(LOF(f) Mod Size)
 If issending = True Then
  For i = 1 To 40
    If frm_main.wsk(i).State = 7 Then
     Get #f, , ss
     frm_main.wsk(i).SendData "newfart" & ss
     bar1.Value = bar1.Value + Len(ss)
     DoEvents
    End If
   Next
 End If
End If

If issending = True Then
For i = 1 To 40
If frm_main.wsk(i).State = 7 Then
frm_main.wsk(i).SendData "enffile" & Text1.Text
bar1.Value = LOF(f)
issending = False
canbutton
DoEvents
End If
Next
End If
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
MsgBox " „ ≈—”«· «·„·› »‰Ã«Õ ·Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…", 64, "·„⁄·Ê„« ﬂ"
bye:
Close f
End Sub

Private Sub Command3_Click()
For i = 1 To 40
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "runfile" & Text1.Text
End If
Next
End Sub

Private Sub Command4_Click()
If issending = True Then
For i = 1 To 40
If frm_main.wsk(nowcompnum).State = 7 Then
  frm_main.wsk(nowcompnum).SendData "stpsend"
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = False
issending = False
bar1.Value = 0
canbutton
End If
Next
MsgBox "›‘· ≈—”«· «·„·› »”»» ≈Ìﬁ«›  ‰›Ì– «·⁄„·Ì¯…", 16, "Œÿ« √À‰«¡ ‰ﬁ· «·„·›"
Else
Unload Me
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

issending = False
bar1.Value = 0
End Sub

Public Sub canbutton()
If issending = True Then
Command4.Caption = "≈·€«¡ «·≈—”«·"
Else
Command4.Caption = "«€·«ﬁ"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If issending = True Then
For i = 1 To 40
If frm_main.wsk(nowcompnum).State = 7 Then
  frm_main.wsk(nowcompnum).SendData "stpsend"
  Command1.Enabled = True
  Command2.Enabled = True
  Command3.Enabled = False
issending = False
bar1.Value = 0
canbutton
End If
Next
MsgBox "›‘· ≈—”«· «·„·› »”»» ≈Ìﬁ«›  ‰›Ì– «·⁄„·Ì¯…", 16, "Œÿ« √À‰«¡ ‰ﬁ· «·„·›"
Else
Unload Me
End If

End Sub
