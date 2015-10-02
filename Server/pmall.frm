VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form pmall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈—”«· —”«·… Œ«’¯… ·Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "pmall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "pmall.frx":29C12
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "pmall.frx":29E46
      Top             =   120
      Width           =   5415
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
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "≈·€«¡ «·√„—"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "pmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
For i = 1 To 40
If frm_main.wsk(i).State = 7 Then
frm_main.wsk(i).SendData "[pmmsg]" & Text1.Text
End If
Next
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

End Sub
