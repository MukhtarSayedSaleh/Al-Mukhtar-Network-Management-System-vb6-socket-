VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Al-Mokhtar System For Inernet Coffee $  ‰Ÿ«„ «·„Œ «— ·≈œ«—… „ﬁ«ÂÌ «·≈‰ —‰ "
   ClientHeight    =   10740
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   15240
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7080
      Top             =   5160
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   135
      Top             =   9960
      Width           =   5775
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "frm_main.frx":29C12
         TabIndex        =   136
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "frm_main.frx":29C73
         TabIndex        =   137
         Top             =   240
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frm_main.frx":29CD8
         TabIndex        =   138
         Top             =   240
         Width           =   2175
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frm_main.frx":29D4B
      TabIndex        =   132
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel labelopentimes 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "frm_main.frx":29DB5
      TabIndex        =   130
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7440
      OleObjectBlob   =   "frm_main.frx":29E2F
      Top             =   5160
   End
   Begin VB.Frame Frame2 
      Height          =   8895
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   15255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3225
         Left            =   120
         Picture         =   "frm_main.frx":2A063
         RightToLeft     =   -1  'True
         ScaleHeight     =   3225
         ScaleWidth      =   15210
         TabIndex        =   131
         Top             =   5400
         Width           =   15210
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   40
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   39
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   38
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   37
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   36
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   35
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   34
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   33
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   32
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   31
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   30
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   29
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   28
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   27
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   26
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   25
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   24
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   23
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   22
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   21
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   20
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   19
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   18
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   17
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   16
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   15
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   14
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   13
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   12
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   11
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   10
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   9
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   8
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   7
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   6
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   5
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   4
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   3
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   2
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.TextBox isClosed 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   1
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   40
         Left            =   12000
         Picture         =   "frm_main.frx":6FA0C
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   48
         Top             =   3840
         Width           =   975
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   0
         Left            =   4680
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   1
         Left            =   120
         Picture         =   "frm_main.frx":70875
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   2
         Left            =   1200
         Picture         =   "frm_main.frx":716DE
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   46
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   3
         Left            =   2280
         Picture         =   "frm_main.frx":72547
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   45
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   4
         Left            =   3360
         Picture         =   "frm_main.frx":733B0
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   44
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   5
         Left            =   4440
         Picture         =   "frm_main.frx":74219
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   43
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   6
         Left            =   5520
         Picture         =   "frm_main.frx":75082
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   7
         Left            =   6600
         Picture         =   "frm_main.frx":75EEB
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   8
         Left            =   7680
         Picture         =   "frm_main.frx":76D54
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   40
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   9
         Left            =   8760
         Picture         =   "frm_main.frx":77BBD
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   10
         Left            =   9840
         Picture         =   "frm_main.frx":78A26
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   11
         Left            =   10920
         Picture         =   "frm_main.frx":7988F
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   12
         Left            =   12000
         Picture         =   "frm_main.frx":7A6F8
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   13
         Left            =   13080
         Picture         =   "frm_main.frx":7B561
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   14
         Left            =   14160
         Picture         =   "frm_main.frx":7C3CA
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   15
         Left            =   120
         Picture         =   "frm_main.frx":7D233
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   33
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   16
         Left            =   1200
         Picture         =   "frm_main.frx":7E09C
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   32
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   17
         Left            =   2280
         Picture         =   "frm_main.frx":7EF05
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   31
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   18
         Left            =   3360
         Picture         =   "frm_main.frx":7FD6E
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   30
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   19
         Left            =   4440
         Picture         =   "frm_main.frx":80BD7
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   29
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   20
         Left            =   5520
         Picture         =   "frm_main.frx":81A40
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   28
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   21
         Left            =   6600
         Picture         =   "frm_main.frx":828A9
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   27
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   22
         Left            =   7680
         Picture         =   "frm_main.frx":83712
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   23
         Left            =   8760
         Picture         =   "frm_main.frx":8457B
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   25
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   24
         Left            =   9840
         Picture         =   "frm_main.frx":853E4
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   24
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   25
         Left            =   10920
         Picture         =   "frm_main.frx":8624D
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   23
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   26
         Left            =   12000
         Picture         =   "frm_main.frx":870B6
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   22
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   27
         Left            =   13080
         Picture         =   "frm_main.frx":87F1F
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   28
         Left            =   14160
         Picture         =   "frm_main.frx":88D88
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   20
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   29
         Left            =   120
         Picture         =   "frm_main.frx":89BF1
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   19
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   30
         Left            =   1200
         Picture         =   "frm_main.frx":8AA5A
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   18
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   31
         Left            =   2280
         Picture         =   "frm_main.frx":8B8C3
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   17
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   32
         Left            =   3360
         Picture         =   "frm_main.frx":8C72C
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   16
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   33
         Left            =   4440
         Picture         =   "frm_main.frx":8D595
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   15
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   34
         Left            =   5520
         Picture         =   "frm_main.frx":8E3FE
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   14
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   35
         Left            =   6600
         Picture         =   "frm_main.frx":8F267
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   13
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   36
         Left            =   7680
         Picture         =   "frm_main.frx":900D0
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   12
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   37
         Left            =   8760
         Picture         =   "frm_main.frx":90F39
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   11
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   38
         Left            =   9840
         Picture         =   "frm_main.frx":91DA2
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   10
         Top             =   3840
         Width           =   975
      End
      Begin VB.PictureBox comp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   39
         Left            =   10920
         Picture         =   "frm_main.frx":92C0B
         RightToLeft     =   -1  'True
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   9
         Top             =   3840
         Width           =   975
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   1
         Left            =   4800
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   2
         Left            =   4560
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   3
         Left            =   2040
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   4
         Left            =   2040
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   5
         Left            =   2040
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   6
         Left            =   2040
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   7
         Left            =   2160
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   8
         Left            =   3120
         Top             =   4560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   9
         Left            =   1320
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   10
         Left            =   2280
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   11
         Left            =   1920
         Top             =   4440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   12
         Left            =   1920
         Top             =   4920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   13
         Left            =   2400
         Top             =   5400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   14
         Left            =   1680
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   15
         Left            =   2040
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   16
         Left            =   2280
         Top             =   4920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   17
         Left            =   3240
         Top             =   5160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   18
         Left            =   3120
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   19
         Left            =   2400
         Top             =   5040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   20
         Left            =   3120
         Top             =   4680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   21
         Left            =   3000
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   22
         Left            =   3480
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   23
         Left            =   3600
         Top             =   4560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   24
         Left            =   3120
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   25
         Left            =   2520
         Top             =   4560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   26
         Left            =   2520
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   27
         Left            =   2280
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   28
         Left            =   2400
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   29
         Left            =   2280
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   30
         Left            =   2520
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   31
         Left            =   2400
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   32
         Left            =   2160
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   33
         Left            =   2040
         Top             =   4080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   34
         Left            =   2400
         Top             =   5040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   35
         Left            =   2520
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   36
         Left            =   2040
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   37
         Left            =   2160
         Top             =   3960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   38
         Left            =   2400
         Top             =   4440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   39
         Left            =   2520
         Top             =   4920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsk 
         Index           =   40
         Left            =   2040
         Top             =   3600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "frm_main.frx":93A74
         TabIndex        =   90
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   2
         Left            =   1200
         OleObjectBlob   =   "frm_main.frx":93ACD
         TabIndex        =   91
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   3
         Left            =   2280
         OleObjectBlob   =   "frm_main.frx":93B26
         TabIndex        =   92
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   4
         Left            =   3360
         OleObjectBlob   =   "frm_main.frx":93B7F
         TabIndex        =   93
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   5
         Left            =   4440
         OleObjectBlob   =   "frm_main.frx":93BD8
         TabIndex        =   94
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   6
         Left            =   5520
         OleObjectBlob   =   "frm_main.frx":93C31
         TabIndex        =   95
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   7
         Left            =   6600
         OleObjectBlob   =   "frm_main.frx":93C8A
         TabIndex        =   96
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   8
         Left            =   7680
         OleObjectBlob   =   "frm_main.frx":93CE3
         TabIndex        =   97
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   9
         Left            =   8760
         OleObjectBlob   =   "frm_main.frx":93D3C
         TabIndex        =   98
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   10
         Left            =   9840
         OleObjectBlob   =   "frm_main.frx":93D95
         TabIndex        =   99
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   11
         Left            =   10920
         OleObjectBlob   =   "frm_main.frx":93DEE
         TabIndex        =   100
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   12
         Left            =   12000
         OleObjectBlob   =   "frm_main.frx":93E47
         TabIndex        =   101
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   13
         Left            =   13080
         OleObjectBlob   =   "frm_main.frx":93EA0
         TabIndex        =   102
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   14
         Left            =   14160
         OleObjectBlob   =   "frm_main.frx":93EF9
         TabIndex        =   103
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   15
         Left            =   120
         OleObjectBlob   =   "frm_main.frx":93F52
         TabIndex        =   104
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   16
         Left            =   1200
         OleObjectBlob   =   "frm_main.frx":93FAB
         TabIndex        =   105
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   17
         Left            =   2280
         OleObjectBlob   =   "frm_main.frx":94004
         TabIndex        =   106
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   18
         Left            =   3360
         OleObjectBlob   =   "frm_main.frx":9405D
         TabIndex        =   107
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   19
         Left            =   4440
         OleObjectBlob   =   "frm_main.frx":940B6
         TabIndex        =   108
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   20
         Left            =   5520
         OleObjectBlob   =   "frm_main.frx":9410F
         TabIndex        =   109
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   21
         Left            =   6600
         OleObjectBlob   =   "frm_main.frx":94168
         TabIndex        =   110
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   22
         Left            =   7680
         OleObjectBlob   =   "frm_main.frx":941C1
         TabIndex        =   111
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   23
         Left            =   8760
         OleObjectBlob   =   "frm_main.frx":9421A
         TabIndex        =   112
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   24
         Left            =   9840
         OleObjectBlob   =   "frm_main.frx":94273
         TabIndex        =   113
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   25
         Left            =   10920
         OleObjectBlob   =   "frm_main.frx":942CC
         TabIndex        =   114
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   26
         Left            =   12000
         OleObjectBlob   =   "frm_main.frx":94325
         TabIndex        =   115
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   27
         Left            =   13080
         OleObjectBlob   =   "frm_main.frx":9437E
         TabIndex        =   116
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   28
         Left            =   14160
         OleObjectBlob   =   "frm_main.frx":943D7
         TabIndex        =   117
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   29
         Left            =   120
         OleObjectBlob   =   "frm_main.frx":94430
         TabIndex        =   118
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   30
         Left            =   1200
         OleObjectBlob   =   "frm_main.frx":94489
         TabIndex        =   119
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   31
         Left            =   2280
         OleObjectBlob   =   "frm_main.frx":944E2
         TabIndex        =   120
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   32
         Left            =   3360
         OleObjectBlob   =   "frm_main.frx":9453B
         TabIndex        =   121
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   33
         Left            =   4440
         OleObjectBlob   =   "frm_main.frx":94594
         TabIndex        =   122
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   34
         Left            =   5520
         OleObjectBlob   =   "frm_main.frx":945ED
         TabIndex        =   123
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   35
         Left            =   6600
         OleObjectBlob   =   "frm_main.frx":94646
         TabIndex        =   124
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   36
         Left            =   7680
         OleObjectBlob   =   "frm_main.frx":9469F
         TabIndex        =   125
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   37
         Left            =   8760
         OleObjectBlob   =   "frm_main.frx":946F8
         TabIndex        =   126
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   38
         Left            =   9840
         OleObjectBlob   =   "frm_main.frx":94751
         TabIndex        =   127
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   39
         Left            =   10920
         OleObjectBlob   =   "frm_main.frx":947AA
         TabIndex        =   128
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel Clabel 
         Height          =   495
         Index           =   40
         Left            =   12000
         OleObjectBlob   =   "frm_main.frx":94803
         TabIndex        =   129
         Top             =   3360
         Width           =   975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   9135
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   16113
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "√ÃÂ“… «·‘»ﬂ…"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«·√—»«Õ Ê «·„’«—Ì›"
         Height          =   855
         Left            =   8160
         Picture         =   "frm_main.frx":9485C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ÕÊ· «·‰Ÿ«„"
         Height          =   855
         Left            =   0
         Picture         =   "frm_main.frx":94C9E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "„Ê«÷Ì⁄ «· ⁄·Ì„« "
         Height          =   855
         Left            =   1440
         Picture         =   "frm_main.frx":950E0
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "≈œ«—… «·«‘ —«ﬂ« "
         Height          =   855
         Left            =   9600
         Picture         =   "frm_main.frx":95822
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "«·ﬂ«› —Ì«"
         Height          =   855
         Left            =   11040
         Picture         =   "frm_main.frx":95C64
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Õ„«Ì… «·‰Ÿ«„"
         Height          =   855
         Left            =   12480
         Picture         =   "frm_main.frx":960A6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "≈⁄œ«œ«  «·»—‰«„Ã"
         Height          =   855
         Left            =   13920
         Picture         =   "frm_main.frx":964E8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "frm_main.frx":9692A
      TabIndex        =   133
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   2760
      OleObjectBlob   =   "frm_main.frx":96994
      TabIndex        =   134
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu MNU_POPUP 
      Caption         =   ""
      NegotiatePosition=   3  'Right
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu openSelect 
         Caption         =   "› Õ «·ÃÂ«“"
      End
      Begin VB.Menu closeSelectes 
         Caption         =   "≈€·«ﬁ «·ÃÂ«“"
      End
      Begin VB.Menu asss 
         Caption         =   "-"
      End
      Begin VB.Menu EndLLL 
         Caption         =   "≈‰Â«¡ Ã·”… «·⁄„·"
      End
      Begin VB.Menu MNU_POPUPAAA 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_ppup_Advance 
         Caption         =   "«·√Ê«„—"
         Begin VB.Menu Screencapt 
            Caption         =   "·ﬁÿ… „‰ «·‘«‘…"
            Begin VB.Menu ForSelecteedscreencapture 
               Caption         =   "··ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu ForAllNetworkCapture 
               Caption         =   "·Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
         Begin VB.Menu ApplicationList 
            Caption         =   "ﬁ«∆„… «· ÿ»Ìﬁ« "
         End
         Begin VB.Menu OpenwebPage 
            Caption         =   "› Õ ’›Õ… «‰ —‰ "
         End
         Begin VB.Menu ppmnu_aaa 
            Caption         =   "-"
         End
         Begin VB.Menu SendFile 
            Caption         =   "«—”«· „·›"
            Begin VB.Menu SendOnly 
               Caption         =   "··ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu ForAllNetWork 
               Caption         =   "·Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
         Begin VB.Menu Pmessage 
            Caption         =   "≈—”«· —”«·… Œ«’…"
            Begin VB.Menu SMPPc 
               Caption         =   "··ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu SPM4all 
               Caption         =   "·Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
         Begin VB.Menu Refresh 
            Caption         =   " ÕœÌÀ"
            Begin VB.Menu SeRefOnly 
               Caption         =   "«·ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu RefRAll 
               Caption         =   "Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
      End
      Begin VB.Menu ppUp_Mnu_Control 
         Caption         =   "«· Õﬂ¯„"
         Begin VB.Menu Close_Clanet 
            Caption         =   "≈€·«ﬁ «·⁄„Ì·"
            Begin VB.Menu CClaientSelect 
               Caption         =   "›Ì «·ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu ForAllNetWorkClos 
               Caption         =   "›Ì Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
         Begin VB.Menu Reboot 
            Caption         =   "≈⁄«œ… ≈ﬁ·«⁄"
            Begin VB.Menu ForSelectedReboot 
               Caption         =   "«·ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu FallNetReboot 
               Caption         =   "Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
         Begin VB.Menu Logoff 
            Caption         =   "≈Ìﬁ«› «· ‘€Ì·"
            Begin VB.Menu SeLogOff 
               Caption         =   "«·ÃÂ«“ «·„Õœœ"
            End
            Begin VB.Menu AllLogOff 
               Caption         =   "Ã„Ì⁄ √ÃÂ“… «·‘»ﬂ…"
            End
         End
      End
   End
   Begin VB.Menu Mnu_help 
      Caption         =   "«· ⁄·Ì„« "
      Begin VB.Menu Mnu_Hlp_con 
         Caption         =   "„Ê«÷Ì⁄ «· ⁄·„Ì« "
      End
      Begin VB.Menu qqq 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Hlp_About 
         Caption         =   "ÕÊ· «·‰Ÿ«„"
      End
   End
   Begin VB.Menu mnu_networkmanagemaent 
      Caption         =   "√œÊ« "
      Begin VB.Menu Mnu_Mngmnt_cafeterea 
         Caption         =   "«·ﬂ«› —Ì«"
      End
      Begin VB.Menu Mnu_mng_cash 
         Caption         =   "«·√—»«Õ Ê «·„’«—Ì›"
      End
      Begin VB.Menu Mnu_Eshtrak 
         Caption         =   "«·«‘ —ﬂ«  Ê «·„‘ —ﬂÌ‰"
      End
      Begin VB.Menu bbb 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Setting 
         Caption         =   "≈⁄œ«œ  «·»—‰«„Ã"
      End
   End
   Begin VB.Menu mnu_view 
      Caption         =   "⁄—÷"
      Begin VB.Menu mnu_view_toolbars 
         Caption         =   "‘—Ìÿ «·√œÊ« "
      End
      Begin VB.Menu mnu_view_statiusbar 
         Caption         =   "‘—Ìÿ «·Õ«·…"
      End
   End
   Begin VB.Menu mnu_reports 
      Caption         =   "«· ﬁ«—Ì—"
      Begin VB.Menu mnu_perort_cash 
         Caption         =   "«· ﬁ«—Ì— «·„«·Ì¯…"
         Begin VB.Menu mnu_prt_cash_increment 
            Caption         =   "«·√—»«Õ"
         End
         Begin VB.Menu mnu_rpt_cash_decremenet 
            Caption         =   "«·„’—Ê›« "
         End
      End
      Begin VB.Menu mnu_rtp_aaa 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rpt_caft 
         Caption         =   "«·ﬂ«› —Ì«"
      End
      Begin VB.Menu mnu_rpt_bbb 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_rpt_users 
         Caption         =   "«·„‘ —ﬂÌ‰"
         Begin VB.Menu Ingen 
            Caption         =   "»‘ﬂ· ⁄«„"
         End
         Begin VB.Menu enddddaas 
            Caption         =   "«·«‘ —«ﬂ«  «·„‰ ÂÌ…"
         End
      End
   End
   Begin VB.Menu mnu_file 
      Caption         =   "„·›"
      Begin VB.Menu mnu_file_exit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AllLogOff_Click()
Dim i As Integer
For i = 1 To 40
If wsk(i).State = 7 Then
wsk(i).SendData "[shtdn]"
End If
Next
End Sub


Private Sub ApplicationList_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "taskmgr"
End If

End Sub

Private Sub CClaientSelect_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "clsclnt"
End If

End Sub



Private Sub closeSelectes_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[mokhatrclose]"
isClosed(nowcompnum).Text = "true"
End If
End Sub

Private Sub Command1_Click()
FrmSetting.Show
StayOnTop FrmSetting

End Sub

Private Sub Command2_Click()
With FrmSetting
.TabStrip1.Tabs(3).Selected = True
.Frame3.Visible = True
.Frame1.Visible = False
.Frame2.Visible = False
.Show
End With
StayOnTop FrmSetting
End Sub

Private Sub Command3_Click()
FrmCaftrea.Show
StayOnTop FrmCaftrea

End Sub

Private Sub Command4_Click()
frmUsers.Show
StayOnTop frmUsers

End Sub

Private Sub Command5_Click()
help.Show
StayOnTop help

End Sub

Private Sub Command6_Click()
About.Show
StayOnTop About

End Sub

Private Sub Command7_Click()
Frmmoney.Show
StayOnTop Frmmoney

End Sub

Private Sub comp_Click(Index As Integer)
Dim i As Integer
i = comp(Index).Index
  If wsk(i).State = 7 Then
    If isClosed(i).Text <> "true" Then
    Call CState(2, i)
    labelopentimes.Visible = True
    SkinLabel1.Visible = True
    SkinLabel2.Visible = True
    SkinLabel3.Visible = True
    labelopentimes.Caption = GetMins(Format(opentimes(i), "Short Time"), Format(Now, "Short Time"))
    FrmSetting.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
    FrmSetting.Data1.RecordSource = "select * from settings"
    FrmSetting.Data1.Refresh
    SkinLabel3.Caption = Round(CDbl(labelopentimes.Caption) * CDbl(FrmSetting.Text2.Text), 2)
    Unload FrmSetting
    End If
  End If
    
End Sub

Private Sub comp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
refreshcomp
nowcompnum = Index
If Button = 2 Then
 Dim i As Integer
 i = comp(Index).Index
 Call CState(2, i)
 Rem √Ê·« ›Ì Õ«· ﬂÊ‰Â „ ’·  ŸÂ— «·ﬁ«∆„… Ê „›⁄·… «·ŒÌ«—« 
 '›Ì Õ«· ﬂÊ‰Â €Ì— „ ’·  ŸÂ— „Ê „›⁄¯·…
  If wsk(i).State = 7 Then
    Mnu_ppup_Advance.Enabled = True
    EndLLL.Enabled = True
    ppUp_Mnu_Control.Enabled = True
    ' ›Ì Õ«·… „€·ﬁ √Ê „› ÊÕ (Ì⁄‰Ì ⁄—÷ «·›·«‘)
    ' Ê Õ”» «·‘—ÿ  ŸÂ— «·ﬁ«∆„ Ì‰ › Õ «·ÃÂ«“ √Ê ≈€·«ﬁÂ
    If isClosed(i).Text = "true" Then
     openSelect.Enabled = True
     EndLLL.Enabled = False
     closeSelectes.Enabled = False
    Else
     openSelect.Enabled = False
     EndLLL.Enabled = True
     closeSelectes.Enabled = True
    labelopentimes.Visible = True
    SkinLabel1.Visible = True
    SkinLabel2.Visible = True
    SkinLabel3.Visible = True
    labelopentimes.Caption = GetMins(Format(opentimes(i), "Short Time"), Format(Now, "Short Time"))
    FrmSetting.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
    FrmSetting.Data1.RecordSource = "select * from settings"
    FrmSetting.Data1.Refresh
    SkinLabel3.Caption = Round(CDbl(labelopentimes.Caption) * CDbl(FrmSetting.Text2.Text), 2)
    Unload FrmSetting
    End If
    frm_main.PopupMenu MNU_POPUP
  Else
         
     openSelect.Enabled = False
     closeSelectes.Enabled = False
     Mnu_ppup_Advance.Enabled = False
     ppUp_Mnu_Control.Enabled = False
     EndLLL.Enabled = False

     frm_main.PopupMenu MNU_POPUP
  End If
End If
End Sub


Private Sub enddddaas_Click()
With dv1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\mokdatabase.dll;Mode=Read|Write"
           .Commands(1).CommandType = adCmdText
          .Commands(1).CommandText = "select * from users where minutesp <=0"
           .Commands(1).Execute
            
         If .rsCommand1.State = 1 Then
         
           .rsCommand1.Close
         
         End If
         
End With
DataReport1.Sections(2).Controls("label1").Caption = " ﬁ—Ì— ⁄‰ «·«‘ —«ﬂ«  «·„‰ ÂÌ…"
DataReport1.Show

End Sub

Private Sub EndLLL_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[endjb]"
End If

End Sub

Private Sub FallNetReboot_Click()
Dim i As Integer
For i = 1 To 40
If wsk(i).State = 7 Then
wsk(i).SendData "[rebot]"
End If
Next
End Sub

Private Sub ForAllNetWork_Click()
frmsendfileall.Show
End Sub

Private Sub ForAllNetworkCapture_Click()
Dim i As Integer
For i = 1 To 40
  If wsk(i).State = 7 Then
wsk(i).SendData "[mokdesktopsnap]"
  End If
Next
End Sub

Private Sub ForAllNetWorkClos_Click()
Dim i As Integer
For i = 1 To 40
If wsk(i).State = 7 Then
wsk(i).SendData "clsclnt"
End If
Next
End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.Path & ("\green.skn")
Skin1.ApplySkin Me.hwnd

wsk(0).Close
wsk(0).LocalPort = 13770
wsk(0).Listen

refreshcomp
Dim i As Integer
i = 0
For i = 1 To 40
isClosed(i).Visible = False
Next
refreshsetting
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
For i = 0 To 40
wsk(i).Close
Next
End
Rem ≈€·«ﬁ «·»Ê— «  «·„› ÊÕ…
'«·Œ—ÊÃ „‰ «·»—‰«„Ã
Dim msgs
msgs = MsgBox("Â·  —Ìœ »«· √ﬂÌœ «·Œ—ÊÃ „‰ «·‰Ÿ«„ ø", vbYesNo + vbQuestion, "‰Ÿ«„ «·„Œ «— ·≈œ«—… „ﬁ«ÂÌ «·≈‰ —‰ ")
If msgs = vbYes Then
Unload Me
End
Else
Exit Sub
End If

End Sub

Private Sub ForSelectedReboot_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[rebot]"
End If

End Sub

Private Sub ForSelecteedscreencapture_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[mokdesktopsnap]"
End If

End Sub

Private Sub Frame2_Click()
refreshcomp
End Sub

Private Sub Ingen_Click()
With dv1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\mokdatabase.dll;Mode=Read|Write"
           .Commands(1).CommandType = adCmdText
          .Commands(1).CommandText = "select * from users"
           .Commands(1).Execute
            
         If .rsCommand1.State = 1 Then
         
           .rsCommand1.Close
         
         End If
         
End With
DataReport1.Sections(2).Controls("label1").Caption = " ﬁ—Ì— ⁄«„ ⁄‰ «·„‘ —ﬂÌ‰ Ê «·«‘ —«ﬂ« "
DataReport1.Show

End Sub

Private Sub Mnu_Eshtrak_Click()
frmUsers.Show
StayOnTop frmUsers
End Sub

Private Sub mnu_file_exit_Click()
Dim i As Integer
For i = 0 To 40
wsk(i).Close
Next
End

Rem ≈€·«ﬁ «·»Ê— «  «·„› ÊÕ…
'«·Œ—ÊÃ „‰ «·»—‰«„Ã
Dim msgs
msgs = MsgBox("Â·  —Ìœ »«· √ﬂÌœ «·Œ—ÊÃ „‰ «·‰Ÿ«„ ø", vbYesNo + vbQuestion, "‰Ÿ«„ «·„Œ «— ·≈œ«—… „ﬁ«ÂÌ «·≈‰ —‰ ")
If msgs = vbYes Then
Unload Me
End
Else
Exit Sub
End If
End Sub

Private Sub Mnu_Hlp_About_Click()
About.Show
StayOnTop About
End Sub

Private Sub Mnu_Hlp_con_Click()
help.Show
StayOnTop help

End Sub

Private Sub Mnu_mng_cash_Click()
Frmmoney.Show
StayOnTop Frmmoney

End Sub

Private Sub Mnu_Mngmnt_cafeterea_Click()
FrmCaftrea.Show
StayOnTop FrmCaftrea

End Sub


Private Sub mnu_prt_cash_increment_Click()
Dim sum As Double
Dim i As Integer
With Frmmoney
.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
.Data1.RecordSource = "select * from in_out_price where type=true order by date"
.Data1.Refresh
On Error Resume Next
.Data1.Recordset.MoveFirst
If .Text4.Text <> "" Then
sum = sum + CDbl(.Text4.Text)
End If
For i = 1 To .Data1.Recordset.RecordCount
On Error Resume Next
.Data1.Recordset.MoveNext
If .Text4.Text <> "" Then
sum = sum + CDbl(.Text4.Text)
End If
Next
End With
Unload Frmmoney
With dv1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\mokdatabase.dll;Mode=Read|Write"
           .Commands(2).CommandType = adCmdText
          .Commands(2).CommandText = "select * from in_out_price where type=true order by date"
           .Commands(2).Execute
            
         If .rsCommand2.State = 1 Then
         
           .rsCommand2.Close
         
         End If
         
End With
DataReport2.Sections(2).Controls("label1").Caption = " ﬁ—Ì— ⁄«„ ⁄‰ «·√—»«Õ"
DataReport2.Sections(5).Controls("label2").Caption = sum
DataReport2.Show

End Sub

Private Sub mnu_rpt_caft_Click()
With dv1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\mokdatabase.dll;Mode=Read|Write"
           .Commands(3).CommandType = adCmdText
          .Commands(3).CommandText = "select * from coffee"
           .Commands(3).Execute
            
         If .rsCommand3.State = 1 Then
         
           .rsCommand3.Close
         
         End If
         
End With
DataReport3.Sections(2).Controls("label1").Caption = " ﬁ—Ì— ⁄«„ ⁄‰ „Õ ÊÌ«  «·ﬂ«› —Ì«"
DataReport3.Show

End Sub

Private Sub mnu_rpt_cash_decremenet_Click()
Dim sum As Double
Dim i As Integer
With Frmmoney
.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
.Data1.RecordSource = "select * from in_out_price where type=false order by date"
.Data1.Refresh
On Error Resume Next
.Data1.Recordset.MoveFirst
If .Text4.Text <> "" Then
sum = sum + CDbl(.Text4.Text)
End If
For i = 1 To .Data1.Recordset.RecordCount
On Error Resume Next
.Data1.Recordset.MoveNext
If .Text4.Text <> "" Then
sum = sum + CDbl(.Text4.Text)
End If
Next
End With
Unload Frmmoney

With dv1
On Error Resume Next
           .Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4;Persist Security Info=False;Data Source=" & App.Path & "\mokdatabase.dll;Mode=Read|Write"
           .Commands(2).CommandType = adCmdText
          .Commands(2).CommandText = "select * from in_out_price where type=false order by date"
           .Commands(2).Execute
            
         If .rsCommand2.State = 1 Then
         
           .rsCommand2.Close
         
         End If
         
End With
DataReport2.Sections(2).Controls("label1").Caption = " ﬁ—Ì— ⁄«„ ⁄‰ «·„’—Ê›« "
DataReport2.Sections(5).Controls("label2").Caption = sum
DataReport2.Show

End Sub

Private Sub mnu_Setting_Click()
FrmSetting.Show
StayOnTop FrmSetting

End Sub

Private Sub mnu_view_statiusbar_Click()
If statebarstate = False Then
Frame3.Visible = True
mnu_view_statiusbar.Checked = True
statebarstate = True
Else
Frame3.Visible = False
mnu_view_statiusbar.Checked = False
statebarstate = False
End If

End Sub

Private Sub mnu_view_toolbars_Click()
If toolbarstate = False Then
Frame1.Visible = True
TabStrip1.Top = 1080
TabStrip1.Height = 9135
Frame2.Top = 960
Frame2.Height = 8895
Picture1.Top = 5400
toolbarstate = True
mnu_view_toolbars.Checked = True
Else
Frame1.Visible = False
TabStrip1.Top = 20
TabStrip1.Height = 10235
Frame2.Top = 5
Frame2.Height = 9925
Picture1.Top = 6600
toolbarstate = False
mnu_view_toolbars.Checked = False
End If
End Sub

Private Sub openSelect_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[mokhatropen]"
opentimes(nowcompnum) = Time
isClosed(nowcompnum).Text = "false"
End If

End Sub

Private Sub OpenwebPage_Click()
Dim mok
mok = InputBox("«·—Ã«¡ √œŒ· —«»ÿ «·’›Õ…", "› Õ ’›Õ… «‰ —‰ ", "www.abokamal.com")

If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[wbrse]" & mok
End If

End Sub

Private Sub Picture1_Click()
refreshcomp
End Sub

Private Sub RefRAll_Click()
Dim i As Integer
For i = 1 To 40
If wsk(i).State = 7 Then
wsk(i).SendData "refdesk"
End If
Next
End Sub

Private Sub SeLogOff_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "[shtdn]"
End If

End Sub

Private Sub SendOnly_Click()
frmSendFile.Show
End Sub

Private Sub SeRefOnly_Click()
If wsk(nowcompnum).State = 7 Then
wsk(nowcompnum).SendData "refdesk"
End If

End Sub


Private Sub SMPPc_Click()
PM.Show
End Sub

Private Sub SPM4all_Click()
pmall.Show
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(1).Selected = True Then
Frame2.Visible = True
refreshcomp
Else
Frame2.Visible = False
End If
End Sub

Private Sub Timer1_Timer()
SkinLabel4.Caption = Time
SkinLabel5.Caption = Date

End Sub

Private Sub wsk_Close(Index As Integer)
reclosecomp
End Sub

Private Sub wsk_Connect(Index As Integer)
    ' €ÌÌ— ‘ﬂ· «·√ÌﬁÊ‰« 
    Call CState(3, Index)
    ' €ÌÌ— «··Ì»· «·„ﬁ«»· ·ﬂ· Ã«Â“
    Clabel(Index).Caption = Index
     ' „«„ «·¬‰ «ŸÂ— «·›·«‘ ›Ì «·ﬂ·«‰Ì  (”ﬂ¯— «·ÃÂ«“)
     refreshcomp
End Sub

Private Sub wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Rem Ã⁄· «·Ê‰”Êﬂ 0 ﬁ«œ—… ⁄·Ï ÿÊ· ⁄·Ï «·≈” ﬁ»«· Ì⁄‰Ì ›Ì Õ«·…
'«‰ Ÿ«— œ«∆„… »ÕÌÀ  ⁄„· »«ﬁÌ «·ÊÌ‰”Êﬂ” ·· ‰”Ìﬁ „⁄ «·√ÃÂ“… «·√Œ—Ï ··⁄„· ⁄·Ï »Ê—  Ê«Õœ
Dim i As Integer
For i = 1 To 40
    If wsk(i).State = 0 Then
    ' ÕÊÌ· «·ÿ·»
    wsk(i).Close
    wsk(i).Accept requestID
    Exit For
    End If
Next

End Sub


Public Function refreshcomp()
Dim i As Integer
For i = 1 To 40
If wsk(i).State = 7 Then
Call CState(3, i)
End If
If wsk(i).State = 0 Then
Call CState(1, i)
End If
If wsk(i).State = 8 Or wsk(i).State = 9 Then
Call CState(4, i)
End If
Clabel(i).Caption = i
Next

If toolbarstate = True Then
mnu_view_toolbars.Checked = True
Else
mnu_view_toolbars.Checked = False
End If

End Function

Public Function reclosecomp()
Dim i As Integer
For i = 1 To 40
If wsk(i).State = 8 Or wsk(i).State = 9 Then
wsk(i).Close
End If
DoEvents
Next
refreshcomp

End Function

Private Sub wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim s As String
Dim ss As String
Dim sss As String
'«” ·«„ «·—”«·… „‰ «·⁄„Ì·
wsk(Index).GetData s
ss = Mid(s, 1, 7)
sss = Mid(s, 8, Len(s))
'›Ì Õ«· ﬂÊ‰ «·⁄„Ì· „€·ﬁ
If s = "ready" Then
 isClosed(Index).Text = "true"
 refreshcomp
End If
'⁄‰œ › Õ «·⁄„Ì·
If s = "opened" Then
 isClosed(Index).Text = "false"
 opentimes(Index) = Time
 refreshcomp
End If
'—”«·… ÿ·» › Õ «·ÃÂ«“
If s = "[openrequest]" Then
 reqmsg.SkinLabel1.Caption = " «·„” Œœ„ «·Ã«·” ⁄·Ï «·ÃÂ«“ —ﬁ„ " & Index & " Ìÿ·» „‰ ”Ì«œ ﬂ„ › Õ «·ÃÂ«“ «·„–ﬂÊ— "
 reqmsg.Show
End If
' ·ﬁÌ „·› ’Ê—… ”ÿÕ «·„ﬂ »
If ss = "newpart" Then
 Dim f As Integer
 Dim fs As String
 f = FreeFile
 Open "c:\desksnap.sna" For Binary As f
 Put #f, (LOF(f) + 1), sss
 Close f
End If
If ss = "endfile" Then
  Close f
  '⁄—÷ «·’Ê—…
  Dim asnap As New DeskSnap
  With asnap
  .Show
  .Caption = "·ﬁÿ… „‰ «·ÃÂ«“ —ﬁ„ " & Index
  .image1.Width = .Width
  .image1.Height = .Height
  .image1.Top = 0
  .image1.Left = 0
  .image1.Picture = LoadPicture("c:\desksnap.sna")
  End With
End If

'  ·ﬁÌ »Ì«‰«  «·œŒÊ· „‰ «”„ „” Œœ„ Ê »«”ÊÊ—œ
'› Õ «·ÃÂ«“ «·⁄„Ì· ·„” Œœ„ „⁄Ì¯‰
If ss = "opncomp" Then
 ' Õ’Ì· ﬂ·„… «·„—Ê— Ê «·ÌÊ“— «·„‘›¯—…
  Dim chn As Integer
  Dim chc As String
  Dim usrpass As String
  Dim iii As Integer
  iii = 0
  For iii = 9 To Len(s)
  chc = Mid(s, iii - 1, 2)
  If chc = "%%" Then
  UsrName = Mid(s, 8, iii - 9)
  usrpass = Mid(s, iii + 1, Len(s))
  End If
  Next
 
 '«· Õﬁﬁ „‰ ’Õ… «·„⁄·Ê„«  «·Œ«’… »«·œŒÊ·
 'ﬂ√œ„‰ √Ê·«
 With signin
 .Data1.RecordSource = "select * from admin where aname='" & UsrName & "'"
 .Data1.Refresh
 If .txtaname.Text = UsrName Then
  If .txtapass.Text = usrpass Then
    If wsk(Index).State = 7 Then
    wsk(Index).SendData "adlogok"
    Exit Sub
    End If
  End If
 End If
 End With
 'ﬂÌÊ“— ⁄«œÌ
 With frmUsers
  .Data1.RecordSource = "select * from users where uname='" & UsrName & "'"
  .Data1.Refresh
  If .Text1.Text = UsrName Then
   If .Text2.Text = usrpass Then
    If Int(.Text6.Text) > 1 Then
     If wsk(Index).State = 7 Then
       wsk(Index).SendData "usrlogn" & .Text6.Text
       Exit Sub
     End If
    End If
   End If
  End If
 End With
 '›‘·  ”ÃÌ· «·œŒÊ·
    If wsk(Index).State = 7 Then
     wsk(Index).SendData "loginff"
    End If
End If

' ⁄‰œ  ”ÃÌ· «·Œ—ÊÃ ›Ì «·ÃÂ«“ «·⁄„Ì·
If ss = "logoutt" Then
' Õ”«» «·Êﬁ  Ê ÿ—ÕÂ „‰ «·»«ﬁÌ
 Dim chn2 As Integer
 Dim chc2 As String
 Dim usrname2 As String
 Dim usedtime
 Dim iii2 As Integer
 iii2 = 0
 For iii2 = 9 To Len(s)
 chc2 = Mid(s, iii2 - 1, 2)
 If chc2 = "%%" Then
 usrname2 = Mid(s, 8, iii2 - 9)
 usedtime = Mid(s, iii2 + 1, Len(s))
 End If
 Next
With frmUsers
 .Data1.RecordSource = "select * from users where uname='" & usrname2 & "'"
 .Data1.Refresh
 If .Text1.Text = usrname2 Then
 If usrname2 <> "" Then
 .Data1.Recordset.Edit
 .Text6.Text = .Text6.Text - (usedtime \ 60)
 .Text5.Text = .Text5.Text - (.Text6.Text \ 60)
 On Error Resume Next
 .Data1.Recordset.MoveNext
 .Data1.Recordset.MovePrevious
 .Data1.Refresh
 End If
 End If
 Unload frmUsers
 Endjob.Show
 Endjob.Text1.Text = Index
 Endjob.Text2.Text = "„Ò‘ —ˆﬂ"
 Endjob.Text3.Text = "„Ò‘ —ˆﬂ"
 Endjob.Text4.Text = "„Ò‘ —ˆﬂ"
 Endjob.Text5.Text = "0.00"
 Endjob.Text6.Text = cafprice(Index)
 Endjob.Text7.Text = Round(CDbl(Endjob.Text5.Text) + CDbl(Endjob.Text6.Text), 2)
 isClosed(Index).Text = "true"
   
End With
End If

'≈‰Â«¡ Ã·”… «·⁄„·
If ss = "endjobs" Then
 Endjob.Show
 Endjob.Text1.Text = Index
 Endjob.Text2.Text = opentimes(Index)
 Endjob.Text3.Text = Time
 Endjob.Text4.Text = GetMins(Format(opentimes(Index), "Short Time"), Format(Now, "Short Time")) & " œﬁÌﬁ… "
     FrmSetting.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
     FrmSetting.Data1.RecordSource = "select * from settings"
     FrmSetting.Data1.Refresh
 Endjob.Text5.Text = Round(CDbl(GetMins(Format(opentimes(Index), "Short Time"), Format(Now, "Short Time"))) * CDbl(FrmSetting.Text2.Text), 2)
 Endjob.Text6.Text = cafprice(Index)
 Endjob.Text7.Text = Round(CDbl(Endjob.Text5.Text) + CDbl(Endjob.Text6.Text), 2)

 Unload FrmSetting
 
 isClosed(Index).Text = "true"
End If

'ﬁ«∆„… «· ÿ»Ìﬁ« 
If ss = "taskmgr" Then
 Dim a As Integer
 Dim fir As String
 Dim snow As String
 For a = 8 To Len(s)
 fir = Mid(s, a, 1)
 If fir = "%" Then
   If snow <> "" Then
    Taskman.List1.AddItem snow
    snow = ""
   End If
 Else
   snow = Trim(snow) & fir
 End If
 Next
 Taskman.Show
End If

'≈—”«· „Õ ÊÌ«  «·ﬂ«› —Ì«
If ss = "caftlst" Then
 Dim cafts As String
 Dim nc As Integer
 FrmCaftrea.Show
 FrmCaftrea.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
 FrmCaftrea.Data1.RecordSource = "select * from coffee"
 FrmCaftrea.Data1.Refresh
 FrmCaftrea.Data1.Recordset.MoveFirst
 cafts = FrmCaftrea.Text1.Text
 On Error Resume Next
 FrmCaftrea.Data1.Recordset.MoveNext
 If cafts = "" Then
  If wsk(Index).State = 7 Then
  wsk(Index).SendData "nocaftr"
  End If
 End If
 For nc = 1 To FrmCaftrea.Data1.Recordset.RecordCount
  cafts = cafts & "%" & FrmCaftrea.Text1.Text
  On Error Resume Next
  FrmCaftrea.Data1.Recordset.MoveNext
 Next
 Unload FrmCaftrea
 If wsk(Index).State = 7 Then
  wsk(Index).SendData "caflist" & cafts & "%"
 End If
End If

'⁄‰œ «” ·«„ ÿ·»Ì¯…
If ss = "reqcaft" Then
Dim qname As String
Dim chnq As Integer
Dim chcq As String
Dim passq As String
Dim iq As Integer
iq = 0
For iq = 9 To Len(s)
chcq = Mid(s, iq - 1, 2)
If chcq = "%%" Then
nameq = Mid(s, 8, iq - 9)
passq = Mid(s, iq + 1, Len(s))
With FrmCaftrea
.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
.Data1.RecordSource = "select * from coffee where nofp='" & nameq & "'"
.Data1.Refresh
If Int(passq) > Int(.Text3.Text) Then
 If frm_main.wsk(Index).State = 7 Then
  wsk(Index).SendData "cafnota"
 End If
End If
If Int(passq) <= Int(.Text3.Text) Then
Dim mok
mok = MsgBox(" «·„” Œœ„ «·Ã«·” ⁄·Ï «·ÃÂ«“ —ﬁ„ " & Index & " Ìÿ·» „‰ «·ﬂ«› —Ì« " & passq & " " & nameq & " Â· ” ﬁÊ„ » ‰›Ì– «·ÿ·» ø ", vbQuestion + vbYesNo, "‰Ÿ«„ «·„Œ «—")
If mok = vbYes Then
.Data1.Recordset.Edit
.Text3.Text = Int(.Text3.Text) - Int(passq)
cafprice(Index) = Int(cafprice(Index)) + (Int(passq) * Int(.Text2.Text))
On Error Resume Next
.Data1.Recordset.MoveNext
.Data1.Recordset.MovePrevious

If wsk(Index).State = 7 Then
wsk(Index).SendData "cafreok"
Unload FrmCaftrea
End If
Else
If wsk(Index).State = 7 Then
wsk(Index).SendData "cafreqn"
End If
End If
End If
End With
End If
Next

End If
End Sub


Public Sub refreshsetting()
'‘—Ìÿ «·√œÊ«  Ê ‘—Ìÿ «·Õ«·… Õ”» «·≈⁄œ«œ« 
    FrmSetting.Data1.DatabaseName = App.Path & ("\mokdatabase.dll")
    FrmSetting.Data1.RecordSource = "select * from settings"
    FrmSetting.Data1.Refresh
    toolbarstate = CBool(FrmSetting.Text6.Text)
    statebarstate = CBool(FrmSetting.Text7.Text)
    runinstartup = CBool(FrmSetting.Text5.Text)
    Unload FrmSetting
    If runinstartup = True Then
    '„› «Õ «·—ÌÃ” —Ì
    RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Mokserver", App.Path & "\" & App.EXEName & ".exe"

    Else
    DeleteRegValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "Mokserver"
    End If
If statebarstate = True Then
Frame3.Visible = True
mnu_view_statiusbar.Checked = True
Else
Frame3.Visible = False
mnu_view_statiusbar.Checked = False

End If

If toolbarstate = True Then
Frame1.Visible = True
TabStrip1.Top = 1080
TabStrip1.Height = 9135
Frame2.Top = 960
Frame2.Height = 8895
Picture1.Top = 5400
toolbarstate = True
mnu_view_toolbars.Checked = True
Else
Frame1.Visible = False
TabStrip1.Top = 20
TabStrip1.Height = 10235
Frame2.Top = 5
Frame2.Height = 9925
Picture1.Top = 6600
toolbarstate = False
mnu_view_toolbars.Checked = False
End If

End Sub

Private Function RegWrite(ByVal Key1, ByVal SValue As String)
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.RegWrite Key1, SValue
End Function


