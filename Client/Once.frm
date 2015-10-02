VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Once 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰Ÿ«„ «·„Œ «— ·≈œ«—… „ﬁ«ÂÌ «·≈‰ —‰  - ≈⁄œ«œ« «·‹ Client"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "Once.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1545
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "Once.frx":29C12
      Top             =   480
   End
   Begin ACTIVESKINLibCtl.SkinLabel label1 
      Height          =   255
      Left            =   3960
      OleObjectBlob   =   "Once.frx":29E46
      TabIndex        =   4
      Top             =   530
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel label2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "Once.frx":29ECC
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„‹Ê«›ﬁ"
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Once"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Function for RegWrite
Private Function RegWrite(ByVal Key1, ByVal SValue As String)
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.RegWrite Key1, SValue
End Function




Private Sub Command1_Click()
Open App.Path & ("\mokip.dll") For Output As #1
Write #1, (LOF(1)), Trim(Text1.Text)
Close #1
Unload Once
frm_closed.Show

End Sub

Private Sub Command2_Click()
Unload Once
End


End Sub

Private Sub Form_Load()
'„› «Õ «·—ÌÃ” —Ì
    RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\MokClient", App.Path & "\" & App.EXEName & ".exe"

'«·”ﬂ‰
Skin1.LoadSkin App.Path & ("\TopSecret.skn")
Skin1.ApplySkin Me.hwnd

'≈Œ›«¡ „‰ ‘—Ìÿ «·„Â«„
Dim style As Long
    Hide
    style = GetWindowLong(hwnd, GWL_EXSTYLE)

        If style And WS_EX_APPWINDOW Then
            style = style - WS_EX_APPWINDOW
        End If
    
        style = style Or WS_EX_APPWINDOW
    SetWindowLong hwnd, GWL_EXSTYLE, style
    App.TaskVisible = False

'·· Õﬁﬁ „‰ ÊÃÊœ „·› ≈⁄œ«œ  «·¬Ì »Ì
Dim file As String
file = App.Path & ("\mokip.dll")
If FileExists(file) = True Then
Unload Once
frm_closed.Show
Else
Text1.Text = "127.0.0.1"
Me.Show
StayOnTop Me
Exit Sub
End If


End Sub

Private Sub Label2_Click()

End Sub
