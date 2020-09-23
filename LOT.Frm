VERSION 4.00
Begin VB.Form frmShowDraws 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "y"
   ClientHeight    =   5565
   ClientLeft      =   345
   ClientTop       =   720
   ClientWidth     =   8880
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808080&
   Height          =   5970
   Icon            =   "LOT.frx":0000
   Left            =   285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Top             =   375
   Width           =   9000
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7920
      Picture         =   "LOT.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   780
      TabIndex        =   105
      Top             =   120
      Width           =   780
   End
   Begin VB.PictureBox AddDraw 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   120
      Picture         =   "LOT.frx":0F38
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   104
      Top             =   5040
      Width           =   15
   End
   Begin VB.PictureBox DrawInfo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   4440
      Picture         =   "LOT.frx":137E
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   103
      Top             =   5040
      Width           =   15
   End
   Begin VB.PictureBox AmendDraw 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   2280
      Picture         =   "LOT.frx":17C4
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   102
      Top             =   5040
      Width           =   15
   End
   Begin VB.VScrollBar Scroll 
      Height          =   4095
      Left            =   8520
      Min             =   1
      TabIndex        =   90
      Top             =   840
      Value           =   1
      Width           =   255
   End
   Begin VB.PictureBox PicExit 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Index           =   0
      Left            =   6600
      Picture         =   "LOT.frx":1C0A
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   5040
      Width           =   15
   End
   Begin VB.TextBox txtLBBon 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtLBDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox txtLB6 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtLB5 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtLB4 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtLB3 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtLB2 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5280
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtLB1 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4800
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   960
      Picture         =   "LOT.frx":2050
      Top             =   60
      Width           =   6795
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   101
      Top             =   720
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   100
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   99
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   98
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   97
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   96
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   95
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   94
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   93
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   92
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label DrawNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   120
      TabIndex        =   91
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   89
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   88
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   87
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   86
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   85
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   84
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   83
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Lin10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   82
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   81
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   80
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   79
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   78
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   77
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   76
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   75
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lin9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   74
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   73
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   72
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   71
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   70
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   69
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   68
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   67
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Lin8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   66
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   65
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   64
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   63
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   62
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   61
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   60
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   59
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Lin7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   58
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   57
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   56
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   55
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   54
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   53
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   52
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   51
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Lin6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   50
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   49
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   48
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   47
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   46
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   45
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   44
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   43
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Lin5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   42
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   41
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   40
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   39
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   38
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   37
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   36
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   35
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Lin4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   34
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   33
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   32
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   31
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   30
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   29
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   28
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   27
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lin3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   26
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   25
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   24
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   23
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   22
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   21
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   20
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   19
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Lin2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   18
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   17
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7200
      TabIndex        =   16
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   15
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   14
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   13
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   11
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Lin1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw No.                  Date of draw                          1      2      3       4     5      6     Bonus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   8295
   End
End
Attribute VB_Name = "frmShowDraws"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim BAplc, L, M, ab
Dim DA%, DB%, DC, DD%, DE%, DF%, DBon, HDS
Dim B1%, B2%, B3%, B4%, B5%, B6%
Dim gPrevLine1
Dim NameToSearch
Dim Found As Integer
Dim fg As Integer
Dim brc As Integer
Dim BKF As Integer
Dim BUL As Integer
Private Sub AddDraw_Click()
Add
End Sub

Private Sub AmendDraw_Click()
AmendEntry
End Sub

Private Sub DrawInfo_Click()
On Error Resume Next
DrawInfo.Picture = frmLottery.ImageList1.ListImages(48).Picture
Dim dialogtype As Integer
Dim dialogtitle As String
Dim dialogmsg As String
Dim response As Integer
dialogtype = vbYesNo + vbExclamation
dialogtitle = "Lottery Draw Result"
dialogmsg = "Delete Lottery Draw Number: " + Str(FB) + "  Are You Sure?"
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
Dim TFN As Long
Dim TDS As DrawsInfo
Dim RCN As Long
Dim TRN
If Dir("Draws.TMP") = "Draws.TMP" Then
Kill "Draws.TMP"
End If
TFN = FreeFile
Open "Draws.TMP" For Random As TFN Len = FE
RCN = 1
TRN = 1
Do While RCN < FD + 1
If RCN <> FB Then
Get #FA, RCN, TDS
Put #TFN, TRN, TDS
TRN = TRN + 1
End If
RCN = RCN + 1
Loop
Close #TFN
Close
Kill "Results.DAT"
Name "Draws.TMP" As "Results.DAT"
FA = FreeFile
Open "Results.DAT" For Random As FA Len = FE
FD = FD - 1
If FD = 0 Then FD = 1
If FB > FD Then
FB = FD
Scroll.Max = Str(FD)
Scroll.Value = Str(FD)
End If
ShowCurrentRecord
ShowPrevLin1
ShowPrevLin2
ShowPrevLin3
ShowPrevLin4
ShowPrevLin5
ShowPrevlin6
ShowPrevLin7
ShowPrevLin8
ShowPrevLin9
ShowPrevLin10
Open "Prizes.DAT" For Random As ZA Len = ZE
frmShowDraws.Caption = Str(FD) + " Lottery draw results stored."
100
DrawInfo.Picture = frmLottery.ImageList1.ListImages(47).Picture
End Sub
Private Sub Form_Activate()
frmShowDraws.Left = (Screen.Width / 2) - (frmShowDraws.Width / 2)
frmShowDraws.Top = (Screen.Height / 2) - (frmShowDraws.Height / 2)
If FD = 0 Then
FD = 1
FB = 1
End If
If FD > 1 Then
FB = FD
End If
frmShowDraws.Caption = Str(FD) + " Lottery draw results stored."
ShowPrevLin1
ShowPrevLin2
ShowPrevLin3
ShowPrevLin4
ShowPrevLin5
ShowPrevlin6
ShowPrevLin7
ShowPrevLin8
ShowPrevLin9
ShowPrevLin10
ShowCurrentRecord
Scroll.Max = Str(FD)
Scroll.Value = Str(FD)
End Sub

Private Sub Form_Load()
PicExit(0).Picture = frmLottery.ImageList1.ListImages(49).Picture
AddDraw.Picture = frmLottery.ImageList1.ListImages(51).Picture
AmendDraw.Picture = frmLottery.ImageList1.ListImages(45).Picture
DrawInfo.Picture = frmLottery.ImageList1.ListImages(47).Picture
txtLB1.Enabled = False
txtLB2.Enabled = False
txtLB3.Enabled = False
txtLB4.Enabled = False
txtLB5.Enabled = False
txtLB6.Enabled = False
txtlbBon.Enabled = False
TxtLbDate.Enabled = False
End Sub
Public Sub SaveCurrentRecord()
'On Error GoTo 7
FC.Ball1 = txtLB1.Text
FC.Ball2 = txtLB2.Text
FC.Ball3 = txtLB3.Text
FC.Ball4 = txtLB4.Text
FC.ball5 = txtLB5.Text
FC.ball6 = txtLB6.Text
FC.BallBon = txtlbBon.Text
FC.Date = TxtLbDate.Text
Put #FA, FB, FC
6 GoTo 9
7 If BKF > 3 Then GoTo 9
MsgBox "A change in the Lottery Draws detected, recommend closing and re-starting Doctor Lottery. Your Lottery Draw Data has been saved.", vbExclamation
BKF = BKF + 1
9
End Sub
Public Sub ShowCurrentRecord()
On Error GoTo 7
Get #FA, FB, FC
txtLB1.Text = Trim(FC.Ball1)
txtLB2.Text = Trim(FC.Ball2)
txtLB3.Text = Trim(FC.Ball3)
txtLB4.Text = Trim(FC.Ball4)
txtLB5.Text = Trim(FC.ball5)
txtLB6.Text = Trim(FC.ball6)
txtlbBon.Text = Trim(FC.BallBon)
TxtLbDate.Text = Trim(FC.Date)
DrawNum(0).Caption = Str(FB)
ShowPrevLin1
ShowPrevLin2
ShowPrevLin3
ShowPrevLin4
ShowPrevLin5
ShowPrevlin6
ShowPrevLin7
ShowPrevLin8
ShowPrevLin9
ShowPrevLin10
6 GoTo 9
7 If BKF > 3 Then GoTo 9
MsgBox "A change in the Lottery Draws detected, recommend closing and re-starting Doctor Lottery. Your Lottery Draw Data has been saved.", vbExclamation
BKF = BKF + 1
FB = 1
9
End Sub


Public Function InputNewNumbers()
Dim UserInput
Dim response As Integer
B1% = 0: B2% = 0: B3% = 0: B4% = 0: B5% = 0: B6% = 0
100 UserInput = InputBox("Enter First Number,  Followed by [ENTER}")
If UserInput = "" Then
HDS = 1
End If
If UserInput = "0" Then
HDS = 1
End If
If HDS = 1 Then
FD = FD - 1
FB = FB - 1
GoTo 911
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 100
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 100
End If
DA% = UserInput
110 UserInput = InputBox("Enter Second Number")
If UserInput = "" Then
Beep
MsgBox "Invalid Number!", vbCritical
GoTo 110
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 110
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 110
End If
If UserInput = DA% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 110
End If
DB% = UserInput
120 UserInput = InputBox("Enter Third Number")
If UserInput = "" Then
Beep
MsgBox "Invalid Number!", vbCritical
GoTo 120
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 120
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 120
End If
If UserInput = DA% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 120
End If
If UserInput = DB% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 120
End If
DC = UserInput
130 UserInput = InputBox("Enter Fourth Number")
If UserInput = "" Then
Beep
MsgBox "Invalid Number!", vbCritical
GoTo 130
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 130
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 130
End If
If UserInput = DA% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 130
End If
If UserInput = DB% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 130
End If
If UserInput = DC Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 130
End If
DD% = UserInput

135 UserInput = InputBox("Enter Fifth Number")
If UserInput = "" Then
Beep
MsgBox "Invalid Number!", vbCritical
GoTo 135
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 135
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 135
End If
If UserInput = DA% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 135
End If
If UserInput = DB% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 135
End If
If UserInput = DC Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 135
End If
If UserInput = DD% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 135
End If
DE% = UserInput
150 UserInput = InputBox("Enter Sixth Number")
If UserInput = "" Then
Beep
MsgBox "Invalid Number!", vbCritical
GoTo 150
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 150
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 150
End If
If UserInput = DA% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 150
End If
If UserInput = DB% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 150
End If
If UserInput = DC Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 150
End If
If UserInput = DD% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 150
End If
If UserInput = DE% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 150
End If
DF% = UserInput
160 UserInput = InputBox("...And the Bonus Ball")
If UserInput = "" Then
Beep
MsgBox "Invalid Number!", vbCritical
GoTo 160
End If
If Not IsNumeric(UserInput) Then
MsgBox "Invalid Number!", vbCritical
GoTo 160
End If
If UserInput > 49 Or UserInput <= 0 Then
MsgBox "Lottery Numbers Range from  1 to 49 !", vbCritical
GoTo 160
End If
If UserInput = DA% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 160
End If
If UserInput = DB% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 160
End If
If UserInput = DC Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 160
End If
If UserInput = DD% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 160
End If
If UserInput = DE% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 160
End If
If UserInput = DF% Then
MsgBox "You can't have two of the same number!", vbCritical
GoTo 160
End If

L = 1
If DA% > DB% Then L = L + 1
If DA% > DC Then L = L + 1
If DA% > DD% Then L = L + 1
If DA% > DE% Then L = L + 1
If DA% > DF% Then L = L + 1
If L = 1 Then B1% = DA%
If L = 2 Then B2% = DA%
If L = 3 Then B3% = DA%
If L = 4 Then B4% = DA%
If L = 5 Then B5% = DA%
If L = 6 Then B6% = DA%


L = 1
If DB% > DA% Then L = L + 1
If DB% > DC Then L = L + 1
If DB% > DD% Then L = L + 1
If DB% > DE% Then L = L + 1
If DB% > DF% Then L = L + 1
If L = 1 Then B1% = DB%
If L = 2 Then B2% = DB%
If L = 3 Then B3% = DB%
If L = 4 Then B4% = DB%
If L = 5 Then B5% = DB%
If L = 6 Then B6% = DB%

L = 1
If DC > DA% Then L = L + 1
If DC > DB% Then L = L + 1
If DC > DD% Then L = L + 1
If DC > DE% Then L = L + 1
If DC > DF% Then L = L + 1
If L = 1 Then B1% = DC
If L = 2 Then B2% = DC
If L = 3 Then B3% = DC
If L = 4 Then B4% = DC
If L = 5 Then B5% = DC
If L = 6 Then B6% = DC

L = 1
If DD% > DA% Then L = L + 1
If DD% > DB% Then L = L + 1
If DD% > DC Then L = L + 1
If DD% > DE% Then L = L + 1
If DD% > DF% Then L = L + 1
If L = 1 Then B1% = DD%
If L = 2 Then B2% = DD%
If L = 3 Then B3% = DD%
If L = 4 Then B4% = DD%
If L = 5 Then B5% = DD%
If L = 6 Then B6% = DD%

L = 1
If DE% > DA% Then L = L + 1
If DE% > DB% Then L = L + 1
If DE% > DC Then L = L + 1
If DE% > DD% Then L = L + 1
If DE% > DF% Then L = L + 1
If L = 1 Then B1% = DE%
If L = 2 Then B2% = DE%
If L = 3 Then B3% = DE%
If L = 4 Then B4% = DE%
If L = 5 Then B5% = DE%
If L = 6 Then B6% = DE%



L = 1
If DF% > DA% Then L = L + 1
If DF% > DB% Then L = L + 1
If DF% > DC Then L = L + 1
If DF% > DD% Then L = L + 1
If DF% > DE% Then L = L + 1
If L = 1 Then B1% = DF%
If L = 2 Then B2% = DF%
If L = 3 Then B3% = DF%
If L = 4 Then B4% = DF%
If L = 5 Then B5% = DF%
If L = 6 Then B6% = DF%
txtLB1.Text = Str(B1%)
txtLB2.Text = Str(B2%)
txtLB3.Text = Str(B3%)
txtLB4.Text = Str(B4%)
txtLB5.Text = Str(B5%)
txtLB6.Text = Str(B6%)
DBon = UserInput
txtlbBon.Text = Str(DBon)
Dim DateInput, DayOfWeek, msg
170 DateInput = InputBox("Enter the Date of the Draw  (If not known enter 'NK')       Example -  January 14, 1999                                         The DAY of the draw is calculated automatically")
If DateInput = Trim("NK") Then
GoTo 175
End If
If DateInput = Trim("nk") Then
GoTo 175
End If
If DateInput = "" Then
MsgBox "Invalid Date!", vbCritical
GoTo 170
End If
If Not IsDate(DateInput) Then
MsgBox "Invalid Date!", vbCritical
GoTo 170
End If
DayOfWeek = Format(DateInput, "dddd")
msg = msg + DayOfWeek
DateInput = StrConv(DateInput, vbProperCase)
TxtLbDate = DayOfWeek + ", " + DateInput
MsgBox "This Entry has been saved"
Exit Function
175 TxtLbDate = Trim("NK")
MsgBox "This Entry has been saved"
911
End Function
Private Sub PicExit_Click(Index As Integer)
FB = 1
frmShowDraws.Hide
If FD = 0 Then FD = 1
If LsetAct = 1 Then
IPUT = FD
End If
End Sub



Private Sub Picture1_Click()
Dim JDA(0 To 6) As Integer
Dim JDX As String
Dim XYY As Printer
Dim HDA As Long
Dim PNR As Integer
'For Each XYY In Printers
'If XYY.Orientation = vbPRORPortrait Then
'Set Printer = XYY
'Exit For
'End If
'Next
PNR = 1
HDA = 500
Printer.Font = "Arial"
Printer.Font.Size = 14
Printer.CurrentX = 2300
Printer.Print "Doctor Lottery, Lottery Draw Results - Page " + Str(PNR)
Printer.Print ""
Printer.Print ""
Printer.Font = "Arial"
Printer.Font.Size = 11
FB = 1
For J = 1 To FD
Printer.Font = "Arial"
Printer.Font.Size = 11
Get #FA, FB, FC
JDA(1) = Str(FC.Ball1)
JDA(2) = Str(FC.Ball2)
JDA(3) = Str(FC.Ball3)
JDA(4) = Str(FC.Ball4)
JDA(5) = Str(FC.ball5)
JDA(6) = Str(FC.ball6)
JDA(0) = Str(FC.BallBon)
JDX = Trim(FC.Date)
Printer.Font = "Arial"
Printer.Font.Size = 11
Printer.CurrentX = 100
Printer.CurrentY = HDA
Printer.Print "    Draw Number: " + Str(FB)
Printer.CurrentX = 2400
Printer.CurrentY = HDA
Printer.Print JDX
Printer.CurrentX = 5900
Printer.CurrentY = HDA
Printer.Print Str(JDA(1))
Printer.CurrentX = 6400
Printer.CurrentY = HDA
Printer.Print Str(JDA(2))
Printer.CurrentX = 6900
Printer.CurrentY = HDA
Printer.Print Str(JDA(3))
Printer.CurrentX = 7400
Printer.CurrentY = HDA
Printer.Print Str(JDA(4))
Printer.CurrentX = 7900
Printer.CurrentY = HDA
Printer.Print Str(JDA(5))
Printer.CurrentX = 8400
Printer.CurrentY = HDA
Printer.Print Str(JDA(6))
Printer.CurrentX = 8900
Printer.CurrentY = HDA
Printer.Print "Bonus Ball"
Printer.CurrentX = 10150
Printer.CurrentY = HDA
Printer.Print Str(JDA(0))
FB = FB + 1
HDA = HDA + 240
If FB = 59 Or FB = 121 Or FB = 182 Or FB = 243 Or FB = 304 Or FB = 365 Or FB = 426 Or FB = 487 Or FB = 548 Or FB = 609 Or FB = 670 Or FB = 731 Or FB = 792 Or FB = 853 Or FB = 914 Or FB = 975 Or FB = 1036 Or FB = 1097 Or FB = 1158 Or FB = 1219 Or FB = 1280 Or FB = 1341 Or FB = 1402 Or FB = 1463 Then
PNR = PNR + 1
Printer.NewPage
Printer.Font = "Arial"
Printer.Font.Size = 11
Printer.CurrentY = 100
Printer.CurrentX = 100
Printer.Print "Page " + Str(PNR)
HDA = 400
Printer.CurrentY = HDA
End If
If FB > FD Then
Printer.EndDoc
GoTo 100
End If
Next
100 Picture1.Picture = frmLottery.ImageList1.ListImages(57).Picture
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Picture = frmLottery.ImageList1.ListImages(58).Picture
End Sub


Private Sub Scroll_Change()
FB = Scroll.Value
ShowCurrentRecord
End Sub

Private Sub Scroll_Scroll()
FB = Scroll.Value
ShowCurrentRecord
End Sub


Private Sub txtLB1_Click()
TextboxClear
txtLB1.BackColor = QBColor(13)
txtLB1.ForeColor = QBColor(15)
End Sub


Private Sub txtLB2_Click()
TextboxClear
txtLB2.BackColor = QBColor(13)
txtLB2.ForeColor = QBColor(15)
End Sub


Private Sub txtLB3_Click()
TextboxClear
txtLB3.BackColor = QBColor(13)
txtLB3.ForeColor = QBColor(15)
End Sub



Public Sub TextboxClear()
txtLB1.BackColor = QBColor(15)
txtLB2.BackColor = QBColor(15)
txtLB3.BackColor = QBColor(15)
txtLB4.BackColor = QBColor(15)
txtLB5.BackColor = QBColor(15)
txtLB6.BackColor = QBColor(15)
txtlbBon.BackColor = QBColor(15)
TxtLbDate.BackColor = QBColor(15)
txtLB1.ForeColor = QBColor(0)
txtLB2.ForeColor = QBColor(0)
txtLB3.ForeColor = QBColor(0)
txtLB4.ForeColor = QBColor(0)
txtLB5.ForeColor = QBColor(0)
txtLB6.ForeColor = QBColor(0)
txtlbBon.ForeColor = QBColor(0)
TxtLbDate.ForeColor = QBColor(0)
End Sub
Private Sub txtLB4_Click()
TextboxClear
txtLB4.BackColor = QBColor(13)
txtLB4.ForeColor = QBColor(15)
End Sub


Private Sub txtLB5_Click()
TextboxClear
txtLB5.BackColor = QBColor(13)
txtLB5.ForeColor = QBColor(15)
End Sub


Private Sub txtLB6_Click()
TextboxClear
txtLB6.BackColor = QBColor(13)
txtLB6.ForeColor = QBColor(15)
End Sub


Private Sub txtlbBon_Click()
TextboxClear
txtlbBon.BackColor = QBColor(13)
txtlbBon.ForeColor = QBColor(15)
End Sub


Private Sub TxtLbDate_Click()
TextboxClear
TxtLbDate.BackColor = QBColor(13)
TxtLbDate.ForeColor = QBColor(15)
End Sub






















Public Sub ShowPrevLin1()
gPrevLine1 = FB - 1
If FB <= 1 Then
gPrevLine1 = 1
BlankLine1
GoTo 5
End If
Get #FA, gPrevLine1, FC
Lin1(1).Caption = Trim(FC.Ball1)
Lin1(2).Caption = Trim(FC.Ball2)
Lin1(3).Caption = Trim(FC.Ball3)
Lin1(4).Caption = Trim(FC.Ball4)
Lin1(5).Caption = Trim(FC.ball5)
Lin1(6).Caption = Trim(FC.ball6)
Lin1(7).Caption = Trim(FC.BallBon)
Lin1(8).Caption = Trim(FC.Date)
DrawNum(1).Caption = Str(FB - 1)
5
End Sub
Public Sub BlankLine1()
For J = 1 To 8
Lin1(J).Caption = ""
Next
DrawNum(1).Caption = ""
End Sub

Public Sub ShowPrevLin2()
gPrevLine1 = FB - 2
If FB <= 2 Then
gPrevLine1 = 1
BlankLine2
GoTo 7
End If
Get #FA, gPrevLine1, FC
Lin2(1).Caption = Trim(FC.Ball1)
Lin2(2).Caption = Trim(FC.Ball2)
Lin2(3).Caption = Trim(FC.Ball3)
Lin2(4).Caption = Trim(FC.Ball4)
Lin2(5).Caption = Trim(FC.ball5)
Lin2(6).Caption = Trim(FC.ball6)
Lin2(7).Caption = Trim(FC.BallBon)
Lin2(8).Caption = Trim(FC.Date)
DrawNum(2).Caption = Str(FB - 2)
7
End Sub
Public Sub BlankLine2()
For J = 1 To 8
Lin2(J).Caption = ""
Next
DrawNum(2).Caption = ""
End Sub

Public Sub BlankLine3()
For J = 1 To 8
Lin3(J).Caption = ""
Next
DrawNum(3).Caption = ""
End Sub

Public Sub ShowPrevLin3()
gPrevLine1 = FB - 3
If FB <= 3 Then
gPrevLine1 = 1
BlankLine3
GoTo 7
End If
Get #FA, gPrevLine1, FC
Lin3(1).Caption = Trim(FC.Ball1)
Lin3(2).Caption = Trim(FC.Ball2)
Lin3(3).Caption = Trim(FC.Ball3)
Lin3(4).Caption = Trim(FC.Ball4)
Lin3(5).Caption = Trim(FC.ball5)
Lin3(6).Caption = Trim(FC.ball6)
Lin3(7).Caption = Trim(FC.BallBon)
Lin3(8).Caption = Trim(FC.Date)
DrawNum(3).Caption = Str(FB - 3)
7
End Sub
Public Sub BlankLine4()
For J = 1 To 8
Lin4(J).Caption = ""
Next
DrawNum(4).Caption = ""
End Sub

Public Sub BlankLine5()
For J = 1 To 8
Lin5(J).Caption = ""
Next
DrawNum(5).Caption = ""
End Sub

Public Sub ShowPrevLin4()
gPrevLine1 = FB - 4
If FB <= 4 Then
gPrevLine1 = 1
BlankLine4
GoTo 7
End If
Get #FA, gPrevLine1, FC
Lin4(1).Caption = Trim(FC.Ball1)
Lin4(2).Caption = Trim(FC.Ball2)
Lin4(3).Caption = Trim(FC.Ball3)
Lin4(4).Caption = Trim(FC.Ball4)
Lin4(5).Caption = Trim(FC.ball5)
Lin4(6).Caption = Trim(FC.ball6)
Lin4(7).Caption = Trim(FC.BallBon)
Lin4(8).Caption = Trim(FC.Date)
DrawNum(4).Caption = Str(FB - 4)
7
End Sub
Public Sub ShowPrevLin5()
gPrevLine1 = FB - 5
If FB <= 5 Then
gPrevLine1 = 1
BlankLine5
GoTo 7
Else
Get #FA, gPrevLine1, FC
Lin5(1).Caption = Trim(FC.Ball1)
Lin5(2).Caption = Trim(FC.Ball2)
Lin5(3).Caption = Trim(FC.Ball3)
Lin5(4).Caption = Trim(FC.Ball4)
Lin5(5).Caption = Trim(FC.ball5)
Lin5(6).Caption = Trim(FC.ball6)
Lin5(7).Caption = Trim(FC.BallBon)
Lin5(8).Caption = Trim(FC.Date)
DrawNum(5).Caption = Str(FB - 5)
7
End If
End Sub
Public Sub BlankLine6()
For J = 1 To 8
Lin6(J).Caption = ""
Next
DrawNum(6).Caption = ""
End Sub

Public Sub BlankLine7()
For J = 1 To 8
Lin7(J).Caption = ""
Next
DrawNum(7).Caption = ""
End Sub

Public Sub BlankLine8()
For J = 1 To 8
Lin8(J).Caption = ""
Next
DrawNum(8).Caption = ""
End Sub

Public Sub BlankLine9()
For J = 1 To 8
Lin9(J).Caption = ""
Next
DrawNum(9).Caption = ""
End Sub

Public Sub BlankLine10()
For J = 1 To 8
Lin10(J).Caption = ""
Next
DrawNum(10).Caption = ""
End Sub

Public Sub ShowPrevlin6()
gPrevLine1 = FB - 6
If FB <= 6 Then
gPrevLine1 = 1
BlankLine6
GoTo 7
Else
Get #FA, gPrevLine1, FC
Lin6(1).Caption = Trim(FC.Ball1)
Lin6(2).Caption = Trim(FC.Ball2)
Lin6(3).Caption = Trim(FC.Ball3)
Lin6(4).Caption = Trim(FC.Ball4)
Lin6(5).Caption = Trim(FC.ball5)
Lin6(6).Caption = Trim(FC.ball6)
Lin6(7).Caption = Trim(FC.BallBon)
Lin6(8).Caption = Trim(FC.Date)
DrawNum(6).Caption = Str(FB - 6)
7
End If
End Sub
Public Sub ShowPrevLin7()
gPrevLine1 = FB - 7
If FB <= 7 Then
gPrevLine1 = 1
BlankLine7
GoTo 7
Else
Get #FA, gPrevLine1, FC
Lin7(1).Caption = Trim(FC.Ball1)
Lin7(2).Caption = Trim(FC.Ball2)
Lin7(3).Caption = Trim(FC.Ball3)
Lin7(4).Caption = Trim(FC.Ball4)
Lin7(5).Caption = Trim(FC.ball5)
Lin7(6).Caption = Trim(FC.ball6)
Lin7(7).Caption = Trim(FC.BallBon)
Lin7(8).Caption = Trim(FC.Date)
DrawNum(7).Caption = Str(FB - 7)
7
End If
End Sub
Public Sub ShowPrevLin8()
gPrevLine1 = FB - 8
If FB <= 8 Then
gPrevLine1 = 1
BlankLine8
GoTo 7
Else
Get #FA, gPrevLine1, FC
Lin8(1).Caption = Trim(FC.Ball1)
Lin8(2).Caption = Trim(FC.Ball2)
Lin8(3).Caption = Trim(FC.Ball3)
Lin8(4).Caption = Trim(FC.Ball4)
Lin8(5).Caption = Trim(FC.ball5)
Lin8(6).Caption = Trim(FC.ball6)
Lin8(7).Caption = Trim(FC.BallBon)
Lin8(8).Caption = Trim(FC.Date)
DrawNum(8).Caption = Str(FB - 8)
7
End If
End Sub
Public Sub ShowPrevLin9()
gPrevLine1 = FB - 9
If FB <= 9 Then
gPrevLine1 = 1
BlankLine9
GoTo 7
Else
Get #FA, gPrevLine1, FC
Lin9(1).Caption = Trim(FC.Ball1)
Lin9(2).Caption = Trim(FC.Ball2)
Lin9(3).Caption = Trim(FC.Ball3)
Lin9(4).Caption = Trim(FC.Ball4)
Lin9(5).Caption = Trim(FC.ball5)
Lin9(6).Caption = Trim(FC.ball6)
Lin9(7).Caption = Trim(FC.BallBon)
Lin9(8).Caption = Trim(FC.Date)
DrawNum(9).Caption = Str(FB - 9)
7
End If
End Sub

Public Sub ShowPrevLin10()
gPrevLine1 = FB - 10
If FB <= 10 Then
gPrevLine1 = 1
BlankLine10
GoTo 7
Else
Get #FA, gPrevLine1, FC
Lin10(1).Caption = Trim(FC.Ball1)
Lin10(2).Caption = Trim(FC.Ball2)
Lin10(3).Caption = Trim(FC.Ball3)
Lin10(4).Caption = Trim(FC.Ball4)
Lin10(5).Caption = Trim(FC.ball5)
Lin10(6).Caption = Trim(FC.ball6)
Lin10(7).Caption = Trim(FC.BallBon)
Lin10(8).Caption = Trim(FC.Date)
DrawNum(10).Caption = Str(FB - 10)
7
End If
End Sub
Public Function CheckLotNumbers()
If Str(txtLB2.Text) = Str(txtLB1.Text) Then
MsgBox "First and second numbers are the same", vbCritical
End If
If Str(txtLB3.Text) = Str(txtLB1.Text) Then
MsgBox "First and third numbers are the same", vbCritical
End If
If Str(txtLB4.Text) = Str(txtLB1.Text) Then
MsgBox "First and fourth numbers are the same", vbCritical
End If
If Str(txtLB5.Text) = Str(txtLB1.Text) Then
MsgBox "First and fifth numbers are the same", vbCritical
End If
If Str(txtLB6.Text) = Str(txtLB1.Text) Then
MsgBox "First and sixth numbers are the same", vbCritical
End If
If Str(txtlbBon.Text) = Str(txtLB1.Text) Then
MsgBox "First and the bonus numbers are the same", vbCritical
End If
If Str(txtLB3.Text) = Str(txtLB2.Text) Then
MsgBox "Second and third numbers are the same!", vbCritical
End If
If Str(txtLB4.Text) = Str(txtLB2.Text) Then
MsgBox "Second and fourth numbers are the same!", vbCritical
End If
If Str(txtLB5.Text) = Str(txtLB2.Text) Then
MsgBox "Second and fifth numbers are the same!", vbCritical
End If
If Str(txtLB6.Text) = Str(txtLB2.Text) Then
MsgBox "Second and sixth numbers are the same!", vbCritical
End If
If Str(txtlbBon.Text) = Str(txtLB2.Text) Then
MsgBox "Second and the Bonus numbers are the same!", vbCritical
End If
If Str(txtLB4.Text) = Str(txtLB3.Text) Then
MsgBox "Third and fourth numbers are the same!", vbCritical
End If
If Str(txtLB5.Text) = Str(txtLB3.Text) Then
MsgBox "Third and fifth numbers are the same!", vbCritical
End If
If Str(txtLB6.Text) = Str(txtLB3.Text) Then
MsgBox "Third and sixth numbers are the same!", vbCritical
End If
If Str(txtlbBon.Text) = Str(txtLB3.Text) Then
MsgBox "Third and the bonus numbers are the same!", vbCritical
End If
If Str(txtLB5.Text) = Str(txtLB4.Text) Then
MsgBox "Fourth and fifth numbers are the same!", vbCritical
End If
If Str(txtLB6.Text) = Str(txtLB4.Text) Then
MsgBox "Fourth and sixth numbers are the same!", vbCritical
End If
If Str(txtlbBon.Text) = Str(txtLB4.Text) Then
MsgBox "Fourth and the bonus numbers are the same!", vbCritical
End If
If Str(txtLB6.Text) = Str(txtLB5.Text) Then
MsgBox "Fifth and sixth numbers are the same!", vbCritical
End If
If Str(txtlbBon.Text) = Str(txtLB5.Text) Then
MsgBox "Fifth and the bonus numbers are the same!", vbCritical
End If
End Function
Public Function AmendEntry()
TextboxClear
BAplc = BAplc + 1
If BAplc >= 2 Then
BAplc = 0
End If
If BAplc = 1 Then
AmendDraw.Picture = frmLottery.ImageList1.ListImages(46).Picture
txtLB1.Enabled = True
txtLB2.Enabled = True
txtLB3.Enabled = True
txtLB4.Enabled = True
txtLB5.Enabled = True
txtLB6.Enabled = True
txtlbBon.Enabled = True
TxtLbDate.Enabled = True
txtLB1.BackColor = QBColor(13)
txtLB1.ForeColor = QBColor(15)
txtLB1.SetFocus
End If
If BAplc = 0 Then
AmendDraw.Picture = frmLottery.ImageList1.ListImages(45).Picture
CheckLotNumbers
SaveCurrentRecord
txtLB1.Enabled = False
txtLB2.Enabled = False
txtLB3.Enabled = False
txtLB4.Enabled = False
txtLB5.Enabled = False
txtLB6.Enabled = False
txtlbBon.Enabled = False
TxtLbDate.Enabled = False
BUL = FB
FB = 1
FC.Ball1 = "3"
FC.Ball2 = "5"
FC.Ball3 = "14"
FC.Ball4 = "22"
FC.ball5 = "30"
FC.ball6 = "44"
FC.BallBon = "10"
FC.Date = "Saturday, November 19, 1994"
Put #FA, FB, FC
FB = BUL

End If

End Function
Public Function Add()
AddDraw.Picture = frmLottery.ImageList1.ListImages(52).Picture
FD = FD + 1
FB = FD
IPUT = FD
Scroll.Max = FD
Scroll.Value = FD
ShowCurrentRecord
ShowPrevLin1
ShowPrevLin2
ShowPrevLin3
ShowPrevLin4
ShowPrevLin5
ShowPrevlin6
ShowPrevLin7
ShowPrevLin8
ShowPrevLin9
ShowPrevLin10
InputNewNumbers
If HDS = 1 Then
ShowCurrentRecord
ShowPrevLin1
ShowPrevLin2
ShowPrevLin3
ShowPrevLin4
ShowPrevLin5
ShowPrevlin6
ShowPrevLin7
ShowPrevLin8
ShowPrevLin9
ShowPrevLin10
HDS = 0
GoTo 9
End If
SaveCurrentRecord
ShowCurrentRecord
ShowPrevLin1
ShowPrevLin2
ShowPrevLin3
ShowPrevLin4
ShowPrevLin5
ShowPrevlin6
ShowPrevLin7
ShowPrevLin8
ShowPrevLin9
ShowPrevLin10
If HDS = 1 Then GoTo 9
7 If BKF > 3 Then GoTo 9
FB = 1
9 frmShowDraws.Caption = Str(FD) + " Lottery draw results stored."
IPUT = FD
Scroll.Max = FD
Scroll.Value = FD

AddDraw.Picture = frmLottery.ImageList1.ListImages(51).Picture
End Function
