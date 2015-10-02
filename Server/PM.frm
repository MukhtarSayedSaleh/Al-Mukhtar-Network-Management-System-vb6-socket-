VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈—”«· —”«·… Œ«’…"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "PM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "PM.frx":29C12
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "„Ê«›ﬁ"
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "PM.frx":29E46
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "PM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If frm_main.wsk(nowcompnum).State = 7 Then
frm_main.wsk(nowcompnum).SendData "[pmmsg]" & Text1.Text
Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub
