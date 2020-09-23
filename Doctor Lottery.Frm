VERSION 4.00
Begin VB.Form frmLottery 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doctor Lottery"
   ClientHeight    =   6690
   ClientLeft      =   2940
   ClientTop       =   3015
   ClientWidth     =   9660
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   7380
   HelpContextID   =   1
   Icon            =   "Doctor Lottery.frx":0000
   Left            =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Top             =   2385
   Width           =   9780
   Begin VB.PictureBox Tool4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1560
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   100
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL19 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      ScaleHeight     =   255
      ScaleWidth      =   1935
      TabIndex        =   98
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox Pshow 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   97
      Top             =   5200
      Width           =   495
   End
   Begin VB.VScrollBar Vspeed 
      Height          =   5535
      Left            =   9360
      Max             =   1000
      Min             =   1
      TabIndex        =   96
      Top             =   600
      Value           =   100
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8880
      Top             =   5280
   End
   Begin VB.PictureBox TOOL13 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   95
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL14 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   94
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   93
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL16 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   92
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL17 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   91
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL18 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   90
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL9 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   89
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL10 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   88
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox TOOL11 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   87
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox TOOL5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   86
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   85
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   84
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   83
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   82
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HelpContextID   =   14
      Left            =   1200
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   81
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox TOOL2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HelpContextID   =   13
      Left            =   840
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   80
      Top             =   120
      WhatsThisHelpID =   30
      Width           =   375
   End
   Begin VB.PictureBox TOOL1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HelpContextID   =   12
      Left            =   360
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   79
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox STEP 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   78
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox RUN 
      BorderStyle     =   0  'None
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   77
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdFast 
      Caption         =   "&Select Number of Runs"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   22
      Top             =   -1320
      Width           =   2055
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   49
      Left            =   8160
      TabIndex        =   175
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   48
      Left            =   7560
      TabIndex        =   174
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   47
      Left            =   6960
      TabIndex        =   173
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   46
      Left            =   6360
      TabIndex        =   172
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   45
      Left            =   8760
      TabIndex        =   171
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   44
      Left            =   8160
      TabIndex        =   170
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   43
      Left            =   7560
      TabIndex        =   169
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   42
      Left            =   6960
      TabIndex        =   168
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   41
      Left            =   6360
      TabIndex        =   167
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   40
      Left            =   8760
      TabIndex        =   166
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   39
      Left            =   8160
      TabIndex        =   165
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   38
      Left            =   7560
      TabIndex        =   164
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   37
      Left            =   6960
      TabIndex        =   163
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   36
      Left            =   6360
      TabIndex        =   162
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   35
      Left            =   8760
      TabIndex        =   161
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   34
      Left            =   8160
      TabIndex        =   160
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   33
      Left            =   7560
      TabIndex        =   159
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   32
      Left            =   6960
      TabIndex        =   158
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   31
      Left            =   6360
      TabIndex        =   157
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   30
      Left            =   8760
      TabIndex        =   156
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   29
      Left            =   8160
      TabIndex        =   155
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   28
      Left            =   7560
      TabIndex        =   154
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   27
      Left            =   6960
      TabIndex        =   153
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   26
      Left            =   6360
      TabIndex        =   152
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   25
      Left            =   8760
      TabIndex        =   151
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   24
      Left            =   8160
      TabIndex        =   150
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   23
      Left            =   7560
      TabIndex        =   149
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   22
      Left            =   6960
      TabIndex        =   148
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   21
      Left            =   6360
      TabIndex        =   147
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   20
      Left            =   8760
      TabIndex        =   146
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   19
      Left            =   8160
      TabIndex        =   145
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   18
      Left            =   7560
      TabIndex        =   144
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   17
      Left            =   6960
      TabIndex        =   143
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   16
      Left            =   6360
      TabIndex        =   142
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   15
      Left            =   8760
      TabIndex        =   141
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   14
      Left            =   8160
      TabIndex        =   140
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   13
      Left            =   7560
      TabIndex        =   139
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   12
      Left            =   6960
      TabIndex        =   138
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   11
      Left            =   6360
      TabIndex        =   137
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   10
      Left            =   8760
      TabIndex        =   136
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   8160
      TabIndex        =   135
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   7560
      TabIndex        =   134
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   6960
      TabIndex        =   133
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   6360
      TabIndex        =   132
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   8760
      TabIndex        =   131
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   8160
      TabIndex        =   130
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   7560
      TabIndex        =   129
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   6960
      TabIndex        =   128
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label TxtBall 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   6360
      TabIndex        =   127
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   126
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   125
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   124
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   123
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   122
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   121
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   120
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   119
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   118
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4200
      TabIndex        =   117
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   116
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   115
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label TxtOddsla 
      Alignment       =   2  'Center
      Caption         =   "Odds of winning: 1 in:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   114
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label txtOddsL 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Draws:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   113
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   112
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   111
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   110
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   109
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   108
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   107
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Ball 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   106
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label TxtOdds 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   103
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label TxtPwn 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   105
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label TxtPerl 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   104
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label TxtWn 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   102
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label txtNument 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   101
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label TxtBlock 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   8280
      TabIndex        =   99
      Top             =   600
      Width           =   1095
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   61
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":0736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":0A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":0C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":0D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":1598
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":1D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":2558
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":2B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":2D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":2ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3064
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":31FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3398
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3532
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":36CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3866
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":3ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":4068
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":4202
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":439C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":4536
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":47D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":4A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":4D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":4FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":5158
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":52F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":548C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":5626
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":57C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":595A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":5BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":5E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":6038
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":61D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":6810
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":6E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":748C
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":7ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":833C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":8BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":9420
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":9C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":A504
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":AD76
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":B5E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":BE5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":C61C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":CDDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":DC50
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":EAC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":F53C
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":FFB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":11FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Doctor Lottery.frx":14016
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1920
      X2              =   6240
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   2040
      TabIndex        =   76
      Top             =   5280
      Width           =   45
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   415
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   9135
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date of Actual Lottery Draw"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1935
      TabIndex        =   75
      Top             =   1850
      Width           =   4290
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   590
      Left            =   1920
      Top             =   1830
      Width           =   4335
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      Caption         =   "Progress..........."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   0
      TabIndex        =   74
      Top             =   6000
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      HelpContext     =   1
      HelpFile        =   "Lottery.HLP"
      HelpKey         =   "F1"
      MaxFileSize     =   2600
   End
   Begin VB.Label lblsumtext 
      Caption         =   "Sum"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5445
      TabIndex        =   73
      Top             =   3360
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblRanText 
      Alignment       =   2  'Center
      Caption         =   "Ran"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5415
      TabIndex        =   72
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5780
      TabIndex        =   71
      Top             =   3360
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.Label lblRan 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5780
      TabIndex        =   70
      Top             =   3120
      Visible         =   0   'False
      Width           =   475
   End
   Begin VB.Label lblNumDraws 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   69
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblLowRange 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   68
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblHighRange 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   67
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblLowSum 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   66
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblHighSum 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   65
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Lowest Range"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      TabIndex        =   64
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Highest Range"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      TabIndex        =   63
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Highest Sum"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   62
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Lowest Sum"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4680
      TabIndex        =   61
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblDrawDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1920
      TabIndex        =   60
      Top             =   2070
      Width           =   4335
   End
   Begin VB.Label lblLotteryStat 
      Alignment       =   2  'Center
      Caption         =   "(Random Numbers)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   59
      Top             =   2670
      Width           =   1815
   End
   Begin VB.Label lblPlayerStat 
      Alignment       =   2  'Center
      Caption         =   "(Hot Numbers)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   58
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblRangeF 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   57
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblAvRangeF 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   56
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblAvRange 
      Alignment       =   2  'Center
      Caption         =   "Average Range"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   55
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblRange 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Range:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   54
      Top             =   2040
      Width           =   555
   End
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   2160
      TabIndex        =   53
      Top             =   6000
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2160
      TabIndex        =   52
      Top             =   5760
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Sum:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   51
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblTotalF 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   50
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblAvtotal 
      Alignment       =   2  'Center
      Caption         =   "Average Sum"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4560
      TabIndex        =   49
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblAvtotalF 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5640
      TabIndex        =   48
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   47
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   46
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   3720
      TabIndex        =   45
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   5
      Left            =   4320
      TabIndex        =   44
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   6
      Left            =   4920
      TabIndex        =   43
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   5760
      TabIndex        =   42
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lblpb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   41
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblpb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   40
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblpb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   39
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblpb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   4
      Left            =   3720
      TabIndex        =   38
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblpb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   5
      Left            =   4320
      TabIndex        =   37
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblpb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   6
      Left            =   4920
      TabIndex        =   36
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lbllb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   35
      Top             =   2450
      Width           =   495
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -41-   - 42-    -43-     -44-    -45-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   8
      Left            =   6360
      TabIndex        =   34
      Top             =   4725
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -1-      -2-       -3-      -4-      -5-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   6360
      TabIndex        =   33
      Top             =   900
      Width           =   2820
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -6-      -7-       -8-      -9-     -10-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   32
      Top             =   1380
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "  -11-    - 12-    -13-     -14-    -15- "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   31
      Top             =   1845
      Width           =   2925
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "  -16-     -17-    -18-     -19-    -20-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   30
      Top             =   2325
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -21-    -22-    -23-     -24-    -25-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   29
      Top             =   2805
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -26-   - 27-    -28-     -29-    -30-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   6360
      TabIndex        =   28
      Top             =   3285
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -31-    -32-    -33-     -34-    -35-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   6360
      TabIndex        =   27
      Top             =   3770
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "   -36-   - 37-    -38-     -39-    -40-"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   6360
      TabIndex        =   26
      Top             =   4245
      Width           =   2865
   End
   Begin VB.Label lblBall1 
      AutoSize        =   -1  'True
      Caption         =   "  -46-     -47-     -48-    -49- "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   9
      Left            =   6360
      TabIndex        =   25
      Top             =   5200
      Width           =   2355
   End
   Begin VB.Label Txth 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6 Most Common Occurring Numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   24
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label txtc 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6 Least Common Occurring  Numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   23
      Top             =   5160
      Width           =   2900
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   " 1         2       3         4        5         6         Bonus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   21
      Top             =   2920
      Width           =   4290
   End
   Begin VB.Label lblProgWN 
      Alignment       =   1  'Right Justify
      Caption         =   "% Stakes won back..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2040
      TabIndex        =   19
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   2040
      TabIndex        =   18
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2040
      TabIndex        =   17
      Top             =   4560
      Width           =   45
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2040
      TabIndex        =   16
      Top             =   4320
      Width           =   45
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2040
      TabIndex        =   15
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label lblw 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   3840
      Width           =   45
   End
   Begin VB.Label lblWn1 
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   13
      Top             =   3150
      Width           =   75
   End
   Begin VB.Label lblJackpot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Jackpot! "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label lbl5Bonus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "5 + Bonus: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label lbl4Ball 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "4 Ball Match: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label lbl3Ball 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "3 Ball Match: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   945
   End
   Begin VB.Label lbl5Ball 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "5 Ball Match: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   945
   End
   Begin VB.Label lbl2Ball 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "2 Ball  Match:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   80
      TabIndex        =   7
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label lbl1Ball 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "1 Ball Match:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label LblSML 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "%  Stake Money Lost:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   1560
   End
   Begin VB.Label lblSMW 
      AutoSize        =   -1  'True
      Caption         =   "% Stake money Won Back:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1920
   End
   Begin VB.Label lblTWinn 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total Winnings So Far:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   690
      TabIndex        =   3
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label lblSpent 
      AutoSize        =   -1  'True
      Caption         =   "Total Spent on the Lottery:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1845
   End
   Begin VB.Label lblONos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Player Numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Label lblLottnos 
      Alignment       =   2  'Center
      Caption         =   "Lottery Numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   2445
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File   "
      Begin VB.Menu mnuActual 
         Caption         =   "View Actual Lottery Draw numbers"
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLotteryNumbers 
      Caption         =   "Lottery Numbers   "
      Begin VB.Menu mnuRandomNumbers 
         Caption         =   "Use RANDOM Lottery Numbers"
      End
      Begin VB.Menu mnuActualNumbers 
         Caption         =   "Use Actual Lottery Draw Results"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLottNos 
         Caption         =   "View Lottery Draw Numbers"
      End
   End
   Begin VB.Menu mnuPlayerNumbers 
      Caption         =   "Player Numbers   "
      Begin VB.Menu mnuEHotNums 
         Caption         =   "Enter HOT Numbers into the draw..."
      End
      Begin VB.Menu mnuEColdNums 
         Caption         =   "Enter COLD Numbers into the draw..."
      End
      Begin VB.Menu mnuRandom 
         Caption         =   "Enter RANDOM Numbers into the draw..."
      End
      Begin VB.Menu MnuAverage 
         Caption         =   "Enter AVERAGE 6 Numbers into the draw..."
      End
      Begin VB.Menu MnuFixPlayer 
         Caption         =   "Enter FIXED Numbers into the draw..."
      End
      Begin VB.Menu MnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSumFilter 
         Caption         =   "Sum Filter"
      End
      Begin VB.Menu mnuRangeFilter 
         Caption         =   "Range Filter"
      End
   End
   Begin VB.Menu mnuPrizes 
      Caption         =   "Prizes    "
      Begin VB.Menu mnuIntertPrizeVals 
         Caption         =   "Insert Own Prize Values"
      End
      Begin VB.Menu mnuPrizedefault 
         Caption         =   "Use Default Prize Values"
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStakeVal 
         Caption         =   "Change Value of Each Stake"
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "Game Card   "
      Begin VB.Menu MNUhotdISP 
         Caption         =   "Display HOTTEST 6 Lottery Numbers"
      End
      Begin VB.Menu mnuColdDisp 
         Caption         =   "Display COLDEST 6 Lottery Numbers"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayDisp 
         Caption         =   "Display Current PLAYER Numbers"
      End
      Begin VB.Menu mnuLotteryDisp 
         Caption         =   "Display Current LOTTERY Numbers"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu MnuApear 
         Caption         =   "Show number of times each ball has appeared"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuPercentage 
         Caption         =   "Show number of times as Percentage"
      End
      Begin VB.Menu MnuPositions 
         Caption         =   "&Show Lottery Ball Frequency Positions"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents            F1"
      End
      Begin VB.Menu mnuHow 
         Caption         =   "Getting Quick Help"
      End
      Begin VB.Menu mnuTipWiz 
         Caption         =   "Tip of the Day"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Restore all Custom Settings"
      End
      Begin VB.Menu mnuSuggestions 
         Caption         =   "Suggestions?"
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmLottery"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer
Dim AggTotal
Dim AGT As Integer
Dim AGTA As Long
Dim AGTO As Integer
Dim AVRANGEA As Long
Dim AVRANGE As Integer
Dim B(1 To 49) As Long
Dim BP As Long
Dim C As Integer
Dim CD As Integer
Dim CL As Currency
Dim COUNTER As Long
Dim D As Integer
Dim DA, DB, DC, DD, DE, DF
Dim DBon As Integer
Dim E(1 To 6) As Integer
Dim EHN As Integer
Dim ECN As Integer
Dim H(1 To 6) As Integer
Dim HD As Integer
Dim Hrange
Dim HSW As Integer
Dim Hsum
Dim L As Integer
Dim LD As Integer
Dim LRange As Integer
Dim Lsum
Dim M As Integer
Dim MA1 As Integer
Dim NUMENT As Currency
Dim ODD As Long
Dim ODDS As Integer
Dim PAGT As Integer
Dim PB(1 To 6) As Integer
Dim PC As Integer
Dim PD As Integer
Dim PLT As Currency
Dim Per(1 To 49) As Variant
Dim PEWON As Integer
Dim PR(0 To 6) As Currency
Dim Prange As Integer
Dim PWN As Currency
Dim Q As Integer
Dim R As Integer
Dim RANGE As Integer
Dim RangeFil As Integer
Dim RP(0 To 6) As Integer
Dim SD(1 To 6) As Integer
Dim SetRan
Dim SetFix
Dim Stack As Long
Dim StartKey
Dim Stopper
Dim Sumfil As Integer
Dim SW As Integer
Dim TB(0 To 6) As Currency
Dim WKK As Currency
Dim XA, XB, XC, XD, XE, XF
Dim XZ As Integer
Dim ZZX As Integer



Public Function LotBalls()
Randomize Timer
LB(1) = Int((49 * Rnd) + 1)
Do
Randomize Timer
LB(2) = Int((49 * Rnd) + 1)
Loop Until LB(2) <> LB(1)
Do
Randomize Timer
LB(3) = Int((49 * Rnd) + 1)
Loop Until LB(3) <> LB(2) And LB(3) <> LB(1)
Do
Randomize Timer
LB(4) = Int((49 * Rnd) + 1)
Loop Until LB(4) <> LB(3) And LB(4) <> LB(2) And LB(4) <> LB(1)
Do
Randomize Timer
LB(5) = Int((49 * Rnd) + 1)
Loop Until LB(5) <> LB(4) And LB(5) <> LB(3) And LB(5) <> LB(2) And LB(5) <> LB(1)
Do
Randomize Timer
LB(6) = Int((49 * Rnd) + 1)
Loop Until LB(6) <> LB(5) And LB(6) <> LB(4) And LB(6) <> LB(3) And LB(6) <> LB(2) And LB(6) <> LB(1)
Do
Randomize Timer
LB(0) = Int((49 * Rnd) + 1)
Loop Until LB(0) <> LB(6) And LB(0) <> LB(5) And LB(0) <> LB(4) And LB(0) <> LB(3) And LB(0) <> LB(2) And LB(0) <> LB(1)
SortLottery
lbllb(1).Caption = Str(LB(1))
lbllb(2).Caption = Str(LB(2))
lbllb(3).Caption = Str(LB(3))
lbllb(4).Caption = Str(LB(4))
lbllb(5).Caption = Str(LB(5))
lbllb(6).Caption = Str(LB(6))
lbllb(0).Caption = Str(LB(0))
If PC = 1 Then
NumCounter
End If
AGT = LB(1) + LB(2) + LB(3) + LB(4) + LB(5) + LB(6)
End Function



Public Function playerballs()
Randomize Timer
PB(1) = Int((49 * Rnd) + 1)
Do
Randomize Timer
PB(2) = Int((49 * Rnd) + 1)
Loop Until PB(2) <> PB(1)
Do
Randomize Timer
PB(3) = Int((49 * Rnd) + 1)
Loop Until PB(3) <> PB(2) And PB(3) <> PB(1)
Do
Randomize Timer
PB(4) = Int((49 * Rnd) + 1)
Loop Until PB(4) <> PB(3) And PB(4) <> PB(2) And PB(4) <> PB(1)
Do
Randomize Timer
PB(5) = Int((49 * Rnd) + 1)
Loop Until PB(5) <> PB(4) And PB(5) <> PB(3) And PB(5) <> PB(2) And PB(5) <> PB(1)
Do
Randomize Timer
PB(6) = Int((49 * Rnd) + 1)
Loop Until PB(6) <> PB(5) And PB(6) <> PB(4) And PB(6) <> PB(3) And PB(6) <> PB(2) And PB(6) <> PB(1)
PAGT = PB(1) + PB(2) + PB(3) + PB(4) + PB(5) + PB(6)
End Function
Public Function Matcher()
A = 12
D = 13
If PB(1) = LB(1) Then
MA1 = MA1 + 1
lbllb(1).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(1) Then
MA1 = MA1 + 1
lbllb(1).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(1) Then
MA1 = MA1 + 1
lbllb(1).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(1) Then
MA1 = MA1 + 1
lbllb(1).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(1) Then
MA1 = MA1 + 1
lbllb(1).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(1) Then
MA1 = MA1 + 1
lbllb(1).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(2) Then
MA1 = MA1 + 1
lbllb(2).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(2) Then
MA1 = MA1 + 1
lbllb(2).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(2) Then
MA1 = MA1 + 1
lbllb(2).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(2) Then
MA1 = MA1 + 1
lbllb(2).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(2) Then
MA1 = MA1 + 1
lbllb(2).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(2) Then
MA1 = MA1 + 1
lbllb(2).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(3) Then
MA1 = MA1 + 1
lbllb(3).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(3) Then
MA1 = MA1 + 1
lbllb(3).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(3) Then
MA1 = MA1 + 1
lbllb(3).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(3) Then
MA1 = MA1 + 1
lbllb(3).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(3) Then
MA1 = MA1 + 1
lbllb(3).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(3) Then
MA1 = MA1 + 1
lbllb(3).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(4) Then
MA1 = MA1 + 1
lbllb(4).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(4) Then
MA1 = MA1 + 1
lbllb(4).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(4) Then
MA1 = MA1 + 1
lbllb(4).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(4) Then
MA1 = MA1 + 1
lbllb(4).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(4) Then
MA1 = MA1 + 1
lbllb(4).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(4) Then
MA1 = MA1 + 1
lbllb(4).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(5) Then
MA1 = MA1 + 1
lbllb(5).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(5) Then
MA1 = MA1 + 1
lbllb(5).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(5) Then
MA1 = MA1 + 1
lbllb(5).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(5) Then
MA1 = MA1 + 1
lbllb(5).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(5) Then
MA1 = MA1 + 1
lbllb(5).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(5) Then
MA1 = MA1 + 1
lbllb(5).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(6) Then
MA1 = MA1 + 1
lbllb(6).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(6) Then
MA1 = MA1 + 1
lbllb(6).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(6) Then
MA1 = MA1 + 1
lbllb(6).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(6) Then
MA1 = MA1 + 1
lbllb(6).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(6) Then
MA1 = MA1 + 1
lbllb(6).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(6) Then
MA1 = MA1 + 1
lbllb(6).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
Select Case MA1
Case 1
TB(1) = TB(1) + 1
lblw(1).ForeColor = QBColor(D)
Case 2
TB(2) = TB(2) + 1
lblw(2).ForeColor = QBColor(D)
Case 3
lblw(3).ForeColor = QBColor(D)
TB(3) = TB(3) + 1
Case 4
lblw(4).ForeColor = QBColor(D)
TB(4) = TB(4) + 1
Case 5
lblw(5).ForeColor = QBColor(D)
TB(5) = TB(5) + 1
Bonus
Case 6
lblw(6).ForeColor = QBColor(D)
TB(6) = TB(6) + 1
End Select
If TB(6) > 300 Then
MsgBox "Too Many Jackpots!"
End
End If
Ball(1).Caption = Str(TB(1))
Ball(2).Caption = Str(TB(2))
Ball(3).Caption = Str(TB(3))
Ball(4).Caption = Str(TB(4))
Ball(5).Caption = Str(TB(5))
Ball(6).Caption = Str(TB(6))
PR(0) = TB(0) * PRIZE(0)
PR(1) = TB(1) * PRIZE(1)
PR(2) = TB(2) * PRIZE(2)
PR(3) = TB(3) * PRIZE(3)
PR(4) = TB(4) * PRIZE(4)
PR(5) = TB(5) * PRIZE(5)
PR(6) = TB(6) * PRIZE(6)
WKK = PR(1) + PR(2) + PR(3) + PR(4) + PR(5) + PR(0)
NUMENT = NUMENT + Stake
Stack = Stack + 1
CL = Stack - WKK
If Stack = 0 Then
Stack = 1
End If
PLT = CL / Stack * 100
If PLT > 32766 Then PLT = 32766
PWN = WKK / Stack * 100
If PWN > 100 Then PWN = 100
If PRIZE(1) >= 1 Then
ODD = TB(1) + TB(2) + TB(3) + TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(2) >= 1 And PRIZE(1) = 0 Then
ODD = TB(2) + TB(3) + TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(3) >= 1 And PRIZE(1) = 0 And PRIZE(2) = 0 Then
ODD = TB(3) + TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(4) >= 1 And PRIZE(1) = 0 And PRIZE(2) = 0 And PRIZE(3) = 0 Then
ODD = TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(5) >= 1 And PRIZE(1) = 0 And PRIZE(2) = 0 And PRIZE(3) = 0 And PRIZE(4) = 0 Then
ODD = TB(5) + TB(0) + TB(6)
End If
If ODD = 0 Then ODD = 1
ODDS = Stack / ODD
AGTA = AGTA + AGT
AGTO = AGTA / Stack
If Hsum < AGT Then
Hsum = AGT
End If
If Lsum > AGT Then
Lsum = AGT
End If
RangeCalc
txtwn.Caption = "£" + Str(WKK)
txtNument.Caption = "£" + Str(NUMENT)
If PLT <= 0 Then PLT = 0
txtPerl.Caption = Str(PLT) + "%"
txtPWn.Caption = Str(PWN) + "%"
If PWN <= 0 Then PWN = 0
ProgressBar1.Value = Str(PWN)
txtBlock.Caption = Str(IPUT)
txtOdds.Caption = Str(ODDS)
lblTotalF.Caption = Str(AGT)
lblAvtotalF.Caption = Str(AGTO)
lblRangeF.Caption = Str(RANGE)
lblAvRangeF.Caption = Str(AVRANGE)
lblHighRange.Caption = Str(Hrange)
lblLowRange.Caption = Str(LRange)
lblHighSum.Caption = Str(Hsum)
lblLowSum.Caption = Str(Lsum)
lblNumDraws.Caption = Str(Stack)
lblran.Caption = Str(Prange)
lblSum.Caption = Str(PAGT)
MA1 = 0
End Function
Public Function clc()
lbllb(1).BackColor = QBColor(15)
lbllb(2).BackColor = QBColor(15)
lbllb(3).BackColor = QBColor(15)
lbllb(4).BackColor = QBColor(15)
lbllb(5).BackColor = QBColor(15)
lbllb(6).BackColor = QBColor(15)
lbllb(0).BackColor = QBColor(15)
lblPb(1).BackColor = QBColor(15)
lblPb(2).BackColor = QBColor(15)
lblPb(3).BackColor = QBColor(15)
lblPb(4).BackColor = QBColor(15)
lblPb(5).BackColor = QBColor(15)
lblPb(6).BackColor = QBColor(15)
lblw(1).ForeColor = QBColor(0)
lblw(2).ForeColor = QBColor(0)
lblw(3).ForeColor = QBColor(0)
lblw(4).ForeColor = QBColor(0)
lblw(5).ForeColor = QBColor(0)
lblw(0).ForeColor = QBColor(0)
lblw(6).ForeColor = QBColor(0)
CLC2
End Function
Public Function Bonus()
If LB(0) = PB(1) Then
TB(0) = TB(0) + 1
TB(5) = TB(0) - 1
WKK = WKK - PRIZE(5)
WKK = WKK + PRIZE(0)
lbllb(0).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
lblw(0).ForeColor = QBColor(D)
lblw(5).ForeColor = QBColor(0)
End If
If LB(0) = PB(2) Then
TB(0) = TB(0) + 1
TB(5) = TB(5) - 1
WKK = WKK - PRIZE(5)
WKK = WKK + PRIZE(0)
lbllb(0).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
lblw(0).ForeColor = QBColor(D)
lblw(5).ForeColor = QBColor(0)
End If
If LB(0) = PB(3) Then
TB(0) = TB(0) + 1
TB(5) = TB(5) - 1
WKK = WKK - PRIZE(5)
WKK = WKK + PRIZE(0)
lbllb(0).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
lblw(0).ForeColor = QBColor(D)
lblw(5).ForeColor = QBColor(0)
End If
If LB(0) = PB(4) Then
TB(0) = TB(0) + 1
TB(5) = TB(5) - 1
WKK = WKK - PRIZE(5)
WKK = WKK + PRIZE(0)
lbllb(0).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
lblw(0).ForeColor = QBColor(D)
lblw(5).ForeColor = QBColor(0)
End If
If LB(0) = PB(5) Then
TB(0) = TB(0) + 1
TB(5) = TB(5) - 1
WKK = WKK - PRIZE(5)
WKK = WKK + PRIZE(0)
lbllb(0).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
lblw(5).ForeColor = QBColor(D)
lblw(0).ForeColor = QBColor(0)
End If
If LB(0) = PB(6) Then
TB(0) = TB(0) + 1
TB(5) = TB(5) - 1
WKK = WKK + PRIZE(5)
WKK = WKK - PRIZE(0)
lbllb(0).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
lblw(0).ForeColor = QBColor(D)
lblw(5).ForeColor = QBColor(0)
End If
End Function



Private Sub Ball_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub

Private Sub Form_Activate()
frmLottery.Left = (Screen.Width / 2) - (frmLottery.Width / 2)
frmLottery.Top = (Screen.Height / 2) - (frmLottery.Height / 2)
End Sub

Private Sub Form_Load()
On Error Resume Next
mnuRandom.Checked = True
mnuRandomNumbers.Checked = True
Vspeed.Visible = False
Timer1.Enabled = False
lblsumtext.Visible = False
cmdFast.Visible = False
lblSum.Visible = False
lblRanText.Visible = False
lblran.Visible = False
mnuLotteryDisp.Checked = True
TOOL1.Picture = ImageList1.ListImages(4).Picture
TOOL2.Picture = ImageList1.ListImages(20).Picture
TOOL3.Picture = ImageList1.ListImages(33).Picture
TOOL4.Picture = ImageList1.ListImages(2).Picture
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(20).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
TOOL9.Picture = ImageList1.ListImages(25).Picture
TOOL10.Picture = ImageList1.ListImages(29).Picture
TOOL11.Picture = ImageList1.ListImages(27).Picture
TOOL12.Picture = ImageList1.ListImages(31).Picture
TOOL13.Picture = ImageList1.ListImages(13).Picture
TOOL14.Picture = ImageList1.ListImages(18).Picture
TOOL15.Picture = ImageList1.ListImages(23).Picture
TOOL16.Picture = ImageList1.ListImages(15).Picture
TOOL17.Picture = ImageList1.ListImages(39).Picture
TOOL18.Picture = ImageList1.ListImages(6).Picture
TOOL19.Picture = ImageList1.ListImages(55).Picture
STEP.Picture = ImageList1.ListImages(41).Picture
RUN.Picture = ImageList1.ListImages(43).Picture
Pshow.Picture = ImageList1.ListImages(53).Picture
LsetRan = 1
LD = 1
SW = 1
LsetAct = 0
Sumfil = 0
RangeFil = 0
PAGT = 0
Hrange = 0
LRange = 49
Hsum = 0
Lsum = 279
COUNTER = 1
NUMENT = 0
PC = 1
lblPlayerStat.Caption = "(Random Numbers)"
LotBalls
Order
playerballs
hotballdisp
Coldballdisp
lblPb(1).Caption = Str(PB(1))
lblPb(2).Caption = Str(PB(2))
lblPb(3).Caption = Str(PB(3))
lblPb(4).Caption = Str(PB(4))
lblPb(5).Caption = Str(PB(5))
lblPb(6).Caption = Str(PB(6))
'LblRb1.Caption = Str(RB(1)) + " " + Str(RB(2)) + " " + Str(RB(3)) + " " + Str(RB(4)) + " " + Str(RB(5)) + " " + Str(RB(6))
'LblRb1.Caption = Str(LB(1)) + " " + Str(LB(2)) + " " + Str(LB(3)) + " " + Str(LB(4)) + " " + Str(LB(5)) + " " + Str(LB(6))
Matcher
NumDisplay
FE = Len(FC)
FA = FreeFile
Open "Results.DAT" For Random As FA Len = FE
FD = FileLen("Results.DAT") / FE
FB = 1
If FD = 0 Then
FD = 1
End If
ZE = Len(ZC)
ZA = FreeFile
Open "Prizes.DAT" For Random As ZA Len = ZE
If ZD = 0 Then ZD = 1
Get #ZA, 300, ZC
If ZC.Priz6 = 0 Then
PRIZE(1) = 0
PRIZE(2) = 0
PRIZE(3) = 10
PRIZE(4) = 70
PRIZE(5) = 1500
PRIZE(0) = 100000
PRIZE(6) = 8000000
Stake = 1
SF(1) = 1
SF(2) = 2
SF(3) = 3
SF(4) = 4
SF(5) = 5
SF(6) = 6
ZZ = 100
USL = 200
LSL = 100
IPUT = 200
URL = 40
LRL = 10
Else
PRIZE(1) = Trim(ZC.Priz1)
PRIZE(2) = Trim(ZC.Priz2)
PRIZE(3) = Trim(ZC.Priz3)
PRIZE(4) = Trim(ZC.Priz4)
PRIZE(5) = Trim(ZC.Priz5)
PRIZE(0) = Trim(ZC.Priz0)
PRIZE(6) = Trim(ZC.Priz6)
Stake = Trim(ZC.Stak)
SF(1) = Trim(ZC.PL1)
SF(2) = Trim(ZC.PL2)
SF(3) = Trim(ZC.PL3)
SF(4) = Trim(ZC.PL4)
SF(5) = Trim(ZC.PL5)
SF(6) = Trim(ZC.PL6)
ZZ = Trim(ZC.ZZ1)
USL = Trim(ZC.UPSUML)
LSL = Trim(ZC.LOSUML)
IPUT = Trim(ZC.NCDR)
URL = Trim(ZC.UPRANL)
LRL = Trim(ZC.LORANL)
End If
lblw(1).Caption = "£" + Str(PRIZE(1))
lblw(2).Caption = "£" + Str(PRIZE(2))
lblw(3).Caption = "£" + Str(PRIZE(3))
lblw(4).Caption = "£" + Str(PRIZE(4))
lblw(5).Caption = "£" + Str(PRIZE(5))
lblw(0).Caption = "£" + Str(PRIZE(0))
lblw(6).Caption = "£" + Str(PRIZE(6))
Timer1.Interval = ZZ
Vspeed.Value = ZZ
txtBlock.Caption = Str(IPUT)
If Stake <= 0 Then Stake = 1
displayNums
For J = 0 To 6
Ball(J).Caption = Str(TB(J))
Next
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 42
FrmQuickHelp.Show
End If
End Sub


Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 41
FrmQuickHelp.Show
End If
End Sub


Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 38
FrmQuickHelp.Show
End If
End Sub


Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 39
FrmQuickHelp.Show
End If
End Sub


Private Sub lbl1Ball_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub lbl2Ball_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub lbl3Ball_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub

Private Sub lbl4Ball_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub lbl5Ball_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub lbl5Bonus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub lblAvRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 37
FrmQuickHelp.Show
End If
End Sub


Private Sub lblAvRangeF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 37
FrmQuickHelp.Show
End If
End Sub


Private Sub lblAvtotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 40
FrmQuickHelp.Show
End If
End Sub


Private Sub lblAvtotalF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 40
FrmQuickHelp.Show
End If
End Sub


Private Sub lblBall1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 46
FrmQuickHelp.Show
End If
End Sub


Private Sub lblDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 20
FrmQuickHelp.Show
End If
End Sub


Private Sub lblDrawDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 20
FrmQuickHelp.Show
End If
End Sub

Private Sub lblHighRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 38
FrmQuickHelp.Show
End If
End Sub


Private Sub lblHighSum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 41
FrmQuickHelp.Show
End If
End Sub


Private Sub lblJackpot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub lbllb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 34
FrmQuickHelp.Show
End If
End Sub


Private Sub lblLotteryStat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 34
FrmQuickHelp.Show
End If
End Sub


Private Sub lblLottnos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 34
FrmQuickHelp.Show
End If
End Sub


Private Sub lblLowRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 39
FrmQuickHelp.Show
End If
End Sub


Private Sub lblLowSum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 42
FrmQuickHelp.Show
End If
End Sub


Private Sub lblNumDraws_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 30
FrmQuickHelp.Show
End If
End Sub


Private Sub lblONos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 35
FrmQuickHelp.Show
End If
End Sub


Private Sub lblPb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 34
FrmQuickHelp.Show
End If
End Sub


Private Sub lblPlayerStat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 34
FrmQuickHelp.Show
End If
End Sub


Private Sub lblProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 45
FrmQuickHelp.Show
End If
End Sub


Private Sub lblProgWN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 26
FrmQuickHelp.Show
End If
End Sub


Private Sub lblran_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 32
FrmQuickHelp.Show
End If
End Sub


Private Sub lblRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 29
FrmQuickHelp.Show
End If
End Sub


Private Sub lblRangeF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 29
FrmQuickHelp.Show
End If
End Sub


Private Sub lblRanText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 32
FrmQuickHelp.Show
End If

End Sub


Private Sub LblSML_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 26
FrmQuickHelp.Show
End If
End Sub


Private Sub lblSMW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 27
FrmQuickHelp.Show
End If
End Sub


Private Sub lblSpent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 24
FrmQuickHelp.Show
End If
End Sub

Private Sub lblSum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 33
FrmQuickHelp.Show
End If
End Sub


Private Sub lblsumtext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 33
FrmQuickHelp.Show
End If
End Sub


Private Sub lblTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 28
FrmQuickHelp.Show
End If
End Sub


Private Sub lblTotalF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 28
FrmQuickHelp.Show
End If
End Sub


Private Sub lblTWinn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 25
FrmQuickHelp.Show
End If
End Sub


Private Sub lblw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 36
FrmQuickHelp.Show
End If
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuActual_Click()
TOOL4.Picture = ImageList1.ListImages(3).Picture
frmShowDraws.Show 1
txtBlock.Caption = Str(IPUT)
TOOL4.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub mnuActualNumbers_Click()
MActual
End Sub



Private Sub MnuApear_Click()
SW = 1
MnuPercentage.Checked = False
MnuApear.Checked = True

NumDisplay
End Sub

Private Sub MnuAverage_Click()
MAverage
End Sub

Private Sub mnuColdDisp_Click()
DispCold
End Sub

Private Sub mnuContents_Click()
'Dim msg
'Err.Clear
'On Error Resume Next
'Err.Raise 6 ' Generate "Overflow" error.
'If Err.Number <> 0 Then
'msg = "Press F1 or Help to see " & Err.HelpFile & " topic for" & _
'" the following HelpContext: " & Err.HelpContext
'MsgBox msg, , "Error: " & Err.Description, Err.HelpFile, _
'Err.HelpContext
'End If
CommonDialog1.HelpFile = "lottery.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub
Private Sub mnuEColdNums_Click()
MCold
End Sub

Private Sub mnuEHotNums_Click()
MHot
End Sub

Private Sub mnuExit_Click()
ZC.ZZ1 = ZZ
Put #ZA, 300, ZC
Close
frmLottery.Hide
frmorders.Show 1
End
End Sub


Private Sub MnuFixPlayer_Click()
MFixed
End Sub

Private Sub MNUhotdISP_Click()
DispHot
End Sub

Private Sub mnuHow_Click()
HPP = 100
FrmQuickHelp.Show
End Sub

Private Sub mnuIntertPrizeVaLS_Click()
TOOL9.Picture = ImageList1.ListImages(26).Picture
frmPrizes.Show 1
For J = 0 To 6
lblw(J).Caption = "£" + Str(PRIZE(J))
Next
TOOL9.Picture = ImageList1.ListImages(25).Picture
End Sub

Private Sub mnuLotteryDisp_Click()
DispLott
End Sub

Private Sub MnuPercentage_Click()
SW = 0
MnuPercentage.Checked = True
MnuApear.Checked = False

NumDisplay

End Sub

Private Sub mnuPlayDisp_Click()
DispPlay
End Sub

Private Sub MnuPositions_Click()
HSW = 1
frmPositions.Show
End Sub

Private Sub mnuPrizedefault_Click()
TOOL9.Picture = ImageList1.ListImages(25).Picture
PRIZE(1) = 0
PRIZE(2) = 0
PRIZE(3) = 10
PRIZE(4) = 70
PRIZE(5) = 1500
PRIZE(0) = 100000
PRIZE(6) = 8000000
For J = 0 To 6
lblw(J).Caption = "£" + Str(PRIZE(J))
Next
End Sub

Private Sub mnuRandom_Click()
MRandom
End Sub
Private Sub mnuRandomNumbers_Click()
MLRandom
End Sub
Private Sub mnuRangeFilter_Click()
MRangeFilter
End Sub

Private Sub mnuSettings_Click()
Dim dialogtype As Integer
Dim dialogtitle As String
Dim dialogmsg As String
Dim response As Integer
dialogtype = vbYesNo + vbExclamation
dialogtitle = "Restore all Custom Settings"
dialogmsg = "Are you sure?"
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
PRIZE(1) = 0
PRIZE(2) = 0
PRIZE(3) = 10
PRIZE(4) = 70
PRIZE(5) = 1500
PRIZE(0) = 100000
PRIZE(6) = 8000000
Stake = 1
SF(1) = 1
SF(2) = 2
SF(3) = 3
SF(4) = 4
SF(5) = 5
SF(6) = 6
ZZ = 100
Timer1.Interval = ZZ
Vspeed.Value = ZZ
For J = 0 To 6
lblw(J).Caption = "£" + Str(PRIZE(J))
Next
LSL = 100
USL = 200
URL = 40
LRL = 10
IPUT = 300
txtBlock.Caption = Str(IPUT)
100
End Sub

Private Sub mnuStakeVal_Click()
FrmStake.Show 1
GoTo 1000
Dim UserInput
UserInput = InputBox("Enter the Value of each betting Stake: in £")
If UserInput = "" Then
Exit Sub
End If
If UserInput <= 0 Then
MsgBox "You must have a Monetary Value!"
Exit Sub
End If
If Not IsNumeric(UserInput) Then
MsgBox "You must have a Monetary Value!"
Exit Sub
End If
If NUMENT < 2 Then
Stake = UserInput
SystemRestart
End If
Stake = CCur(UserInput)


1000
End Sub

Private Sub mnuSuggestions_Click()
frmSuggestions.Show 1
End Sub

Private Sub mnuSumFilter_Click()
MSumFilter
End Sub


Private Sub mnuTipWiz_Click()
FrmTips.Show 1
End Sub

Private Sub mnuViewLottNos_Click()
TOOL4.Picture = ImageList1.ListImages(3).Picture
frmShowDraws.Show 1
txtBlock.Caption = Str(IPUT)
TOOL4.Picture = ImageList1.ListImages(2).Picture
End Sub



Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 26
FrmQuickHelp.Show
End If
End Sub

Private Sub ProgressBar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 45
FrmQuickHelp.Show
End If
End Sub

Private Sub Pshow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
ZZX = ZZX + 1
If ZZX >= 2 Then ZZX = 0
If ZZX = 0 Then
Pshow.Picture = ImageList1.ListImages(53).Picture
Timer1.Enabled = False
Vspeed.Visible = False
End If
If ZZX = 1 Then
Pshow.Picture = ImageList1.ListImages(54).Picture
Timer1.Enabled = True
Vspeed.Visible = True
End If
End If
If Button = vbRightButton Then
HPP = 21
FrmQuickHelp.Show
End If
End Sub


Private Sub RUN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
RUN.Picture = ImageList1.ListImages(44).Picture
If LsetRan = 1 Then
RunRandom
End If
If LsetAct = 1 Then
SystemRestart
RunActual
End If
Beep
RUN.Picture = ImageList1.ListImages(43).Picture
End If
If Button = vbRightButton Then
HPP = 22
FrmQuickHelp.Show
End If
End Sub


Private Sub STEP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
STEP.Picture = ImageList1.ListImages(42).Picture
If ECN = 0 And EHN = 0 Then GoTo 112
If ECN = 1 Or EHN = 1 Then
StartKey = StartKey + 1
If StartKey < 2 And LsetAct = 1 Then
NUMENT = 0
End If
End If
112
If LsetRan = 1 Then
StepRan
End If
If LsetAct = 1 Then
StepActual
End If
displayNums
STEP.Picture = ImageList1.ListImages(41).Picture
End If
If Button = vbRightButton Then
HPP = 23
FrmQuickHelp.Show
End If
End Sub
Private Sub Timer1_Timer()
If ECN = 0 And EHN = 0 Then GoTo 112
If ECN = 1 Or EHN = 1 Then
StartKey = StartKey + 1
If StartKey < 2 And LsetAct = 1 Then
NUMENT = 0
End If
End If
112
If LsetRan = 1 Then
StepRan
End If
If LsetAct = 1 Then
StepActual
End If
displayNums

End Sub

Private Sub TOOL1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
TOOL1.Picture = ImageList1.ListImages(5).Picture
SystemRestart
TOOL1.Picture = ImageList1.ListImages(4).Picture
End If
If Button = vbRightButton Then
HPP = 2
FrmQuickHelp.Show
End If
End Sub

Private Sub TOOL10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MSumFilter
End If
If Button = vbRightButton Then
HPP = 1
FrmQuickHelp.Show
End If
End Sub








Private Sub TOOL11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MRangeFilter
End If
If Button = vbRightButton Then
HPP = 10
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
FrmTips.Show 1
End If
If Button = vbRightButton Then
HPP = 12
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
CommonDialog1.HelpFile = "lottery.hlp"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End If
If Button = vbRightButton Then
HPP = 13
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
DispLott
End If
If Button = vbRightButton Then
HPP = 14
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
DispPlay
End If
If Button = vbRightButton Then
HPP = 15
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
DispHot
End If
If Button = vbRightButton Then
HPP = 16
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
DispCold
End If
If Button = vbRightButton Then
HPP = 17
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Timer1.Enabled = False
ZC.ZZ1 = ZZ
Put #ZA, 300, ZC
Close
frmLottery.Hide
frmorders.Show 1
End
End If
If Button = vbRightButton Then
HPP = 18
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
If LsetAct = 1 Then
MsgBox "You are running ACTUAL Lottery Draws, there are " + Str(IPUT) + " of them"
GoTo 100
Else
TOOL19.Picture = ImageList1.ListImages(56).Picture
End If
FrmRuns.Show 1
txtBlock.Caption = Str(IPUT)
100
TOOL19.Picture = ImageList1.ListImages(55).Picture
End If
If Button = vbRightButton Then
HPP = 19
FrmQuickHelp.Show
End If

End Sub

Private Sub TOOL2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MLRandom
End If
If Button = vbRightButton Then
HPP = 3
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MActual
End If
If Button = vbRightButton Then
HPP = 4
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
TOOL4.Picture = ImageList1.ListImages(3).Picture
frmShowDraws.Show 1
txtBlock.Caption = Str(IPUT)
TOOL4.Picture = ImageList1.ListImages(2).Picture
End If
If Button = vbRightButton Then
HPP = 5
FrmQuickHelp.Show
End If
End Sub
Private Sub TOOL5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MHot
End If
If Button = vbRightButton Then
HPP = 6
FrmQuickHelp.Show
End If
End Sub

Private Sub TOOL6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MCold
End If
If Button = vbRightButton Then
HPP = 7
FrmQuickHelp.Show
End If
End Sub

Private Sub TOOL7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MRandom
End If
If Button = vbRightButton Then
HPP = 8
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
MFixed
End If
If Button = vbRightButton Then
HPP = 9
FrmQuickHelp.Show
End If
End Sub


Private Sub TOOL9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
TOOL9.Picture = ImageList1.ListImages(26).Picture
frmPrizes.Show 1
For J = 0 To 6
lblw(J).Caption = "£" + Str(PRIZE(J))
Next
TOOL9.Picture = ImageList1.ListImages(25).Picture
End If
If Button = vbRightButton Then
HPP = 11
FrmQuickHelp.Show
End If





End Sub

Private Sub txtBall_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 46
FrmQuickHelp.Show
End If
End Sub

Private Sub txtBlock_Change()
txtBlock.Caption = Str(IPUT)
End Sub



Public Function NumCounter()
If LB(1) = 1 Then B(1) = B(1) + 1
If LB(2) = 1 Then B(1) = B(1) + 1
If LB(3) = 1 Then B(1) = B(1) + 1
If LB(4) = 1 Then B(1) = B(1) + 1
If LB(5) = 1 Then B(1) = B(1) + 1
If LB(6) = 1 Then B(1) = B(1) + 1
If LB(1) = 2 Then B(2) = B(2) + 1
If LB(2) = 2 Then B(2) = B(2) + 1
If LB(3) = 2 Then B(2) = B(2) + 1
If LB(4) = 2 Then B(2) = B(2) + 1
If LB(5) = 2 Then B(2) = B(2) + 1
If LB(6) = 2 Then B(2) = B(2) + 1
If LB(1) = 3 Then B(3) = B(3) + 1
If LB(2) = 3 Then B(3) = B(3) + 1
If LB(3) = 3 Then B(3) = B(3) + 1
If LB(4) = 3 Then B(3) = B(3) + 1
If LB(5) = 3 Then B(3) = B(3) + 1
If LB(6) = 3 Then B(3) = B(3) + 1
If LB(1) = 4 Then B(4) = B(4) + 1
If LB(2) = 4 Then B(4) = B(4) + 1
If LB(3) = 4 Then B(4) = B(4) + 1
If LB(4) = 4 Then B(4) = B(4) + 1
If LB(5) = 4 Then B(4) = B(4) + 1
If LB(6) = 4 Then B(4) = B(4) + 1
If LB(1) = 5 Then B(5) = B(5) + 1
If LB(2) = 5 Then B(5) = B(5) + 1
If LB(3) = 5 Then B(5) = B(5) + 1
If LB(4) = 5 Then B(5) = B(5) + 1
If LB(5) = 5 Then B(5) = B(5) + 1
If LB(6) = 5 Then B(5) = B(5) + 1
If LB(1) = 6 Then B(6) = B(6) + 1
If LB(2) = 6 Then B(6) = B(6) + 1
If LB(3) = 6 Then B(6) = B(6) + 1
If LB(4) = 6 Then B(6) = B(6) + 1
If LB(5) = 6 Then B(6) = B(6) + 1
If LB(6) = 6 Then B(6) = B(6) + 1
If LB(1) = 7 Then B(7) = B(7) + 1
If LB(2) = 7 Then B(7) = B(7) + 1
If LB(3) = 7 Then B(7) = B(7) + 1
If LB(4) = 7 Then B(7) = B(7) + 1
If LB(5) = 7 Then B(7) = B(7) + 1
If LB(6) = 7 Then B(7) = B(7) + 1
If LB(1) = 8 Then B(8) = B(8) + 1
If LB(2) = 8 Then B(8) = B(8) + 1
If LB(3) = 8 Then B(8) = B(8) + 1
If LB(4) = 8 Then B(8) = B(8) + 1
If LB(5) = 8 Then B(8) = B(8) + 1
If LB(6) = 8 Then B(8) = B(8) + 1
If LB(1) = 9 Then B(9) = B(9) + 1
If LB(2) = 9 Then B(9) = B(9) + 1
If LB(3) = 9 Then B(9) = B(9) + 1
If LB(4) = 9 Then B(9) = B(9) + 1
If LB(5) = 9 Then B(9) = B(9) + 1
If LB(6) = 9 Then B(9) = B(9) + 1
If LB(1) = 10 Then B(10) = B(10) + 1
If LB(2) = 10 Then B(10) = B(10) + 1
If LB(3) = 10 Then B(10) = B(10) + 1
If LB(4) = 10 Then B(10) = B(10) + 1
If LB(5) = 10 Then B(10) = B(10) + 1
If LB(6) = 10 Then B(10) = B(10) + 1
If LB(1) = 11 Then B(11) = B(11) + 1
If LB(2) = 11 Then B(11) = B(11) + 1
If LB(3) = 11 Then B(11) = B(11) + 1
If LB(4) = 11 Then B(11) = B(11) + 1
If LB(5) = 11 Then B(11) = B(11) + 1
If LB(6) = 11 Then B(11) = B(11) + 1
If LB(1) = 12 Then B(12) = B(12) + 1
If LB(2) = 12 Then B(12) = B(12) + 1
If LB(3) = 12 Then B(12) = B(12) + 1
If LB(4) = 12 Then B(12) = B(12) + 1
If LB(5) = 12 Then B(12) = B(12) + 1
If LB(6) = 12 Then B(12) = B(12) + 1
If LB(1) = 13 Then B(13) = B(13) + 1
If LB(2) = 13 Then B(13) = B(13) + 1
If LB(3) = 13 Then B(13) = B(13) + 1
If LB(4) = 13 Then B(13) = B(13) + 1
If LB(5) = 13 Then B(13) = B(13) + 1
If LB(6) = 13 Then B(13) = B(13) + 1
If LB(1) = 14 Then B(14) = B(14) + 1
If LB(2) = 14 Then B(14) = B(14) + 1
If LB(3) = 14 Then B(14) = B(14) + 1
If LB(4) = 14 Then B(14) = B(14) + 1
If LB(5) = 14 Then B(14) = B(14) + 1
If LB(6) = 14 Then B(14) = B(14) + 1
If LB(1) = 15 Then B(15) = B(15) + 1
If LB(2) = 15 Then B(15) = B(15) + 1
If LB(3) = 15 Then B(15) = B(15) + 1
If LB(4) = 15 Then B(15) = B(15) + 1
If LB(5) = 15 Then B(15) = B(15) + 1
If LB(6) = 15 Then B(15) = B(15) + 1
If LB(1) = 16 Then B(16) = B(16) + 1
If LB(2) = 16 Then B(16) = B(16) + 1
If LB(3) = 16 Then B(16) = B(16) + 1
If LB(4) = 16 Then B(16) = B(16) + 1
If LB(5) = 16 Then B(16) = B(16) + 1
If LB(6) = 16 Then B(16) = B(16) + 1
If LB(1) = 17 Then B(17) = B(17) + 1
If LB(2) = 17 Then B(17) = B(17) + 1
If LB(3) = 17 Then B(17) = B(17) + 1
If LB(4) = 17 Then B(17) = B(17) + 1
If LB(5) = 17 Then B(17) = B(17) + 1
If LB(6) = 17 Then B(17) = B(17) + 1
If LB(1) = 18 Then B(18) = B(18) + 1
If LB(2) = 18 Then B(18) = B(18) + 1
If LB(3) = 18 Then B(18) = B(18) + 1
If LB(4) = 18 Then B(18) = B(18) + 1
If LB(5) = 18 Then B(18) = B(18) + 1
If LB(6) = 18 Then B(18) = B(18) + 1
If LB(1) = 19 Then B(19) = B(19) + 1
If LB(2) = 19 Then B(19) = B(19) + 1
If LB(3) = 19 Then B(19) = B(19) + 1
If LB(4) = 19 Then B(19) = B(19) + 1
If LB(5) = 19 Then B(19) = B(19) + 1
If LB(6) = 19 Then B(19) = B(19) + 1
If LB(1) = 20 Then B(20) = B(20) + 1
If LB(2) = 20 Then B(20) = B(20) + 1
If LB(3) = 20 Then B(20) = B(20) + 1
If LB(4) = 20 Then B(20) = B(20) + 1
If LB(5) = 20 Then B(20) = B(20) + 1
If LB(6) = 20 Then B(20) = B(20) + 1
If LB(1) = 21 Then B(21) = B(21) + 1
If LB(2) = 21 Then B(21) = B(21) + 1
If LB(3) = 21 Then B(21) = B(21) + 1
If LB(4) = 21 Then B(21) = B(21) + 1
If LB(5) = 21 Then B(21) = B(21) + 1
If LB(6) = 21 Then B(21) = B(21) + 1
If LB(1) = 22 Then B(22) = B(22) + 1
If LB(2) = 22 Then B(22) = B(22) + 1
If LB(3) = 22 Then B(22) = B(22) + 1
If LB(4) = 22 Then B(22) = B(22) + 1
If LB(5) = 22 Then B(22) = B(22) + 1
If LB(6) = 22 Then B(22) = B(22) + 1
If LB(1) = 23 Then B(23) = B(23) + 1
If LB(2) = 23 Then B(23) = B(23) + 1
If LB(3) = 23 Then B(23) = B(23) + 1
If LB(4) = 23 Then B(23) = B(23) + 1
If LB(5) = 23 Then B(23) = B(23) + 1
If LB(6) = 23 Then B(23) = B(23) + 1
If LB(1) = 24 Then B(24) = B(24) + 1
If LB(2) = 24 Then B(24) = B(24) + 1
If LB(3) = 24 Then B(24) = B(24) + 1
If LB(4) = 24 Then B(24) = B(24) + 1
If LB(5) = 24 Then B(24) = B(24) + 1
If LB(6) = 24 Then B(24) = B(24) + 1
If LB(1) = 25 Then B(25) = B(25) + 1
If LB(2) = 25 Then B(25) = B(25) + 1
If LB(3) = 25 Then B(25) = B(25) + 1
If LB(4) = 25 Then B(25) = B(25) + 1
If LB(5) = 25 Then B(25) = B(25) + 1
If LB(6) = 25 Then B(25) = B(25) + 1
If LB(1) = 26 Then B(26) = B(26) + 1
If LB(2) = 26 Then B(26) = B(26) + 1
If LB(3) = 26 Then B(26) = B(26) + 1
If LB(4) = 26 Then B(26) = B(26) + 1
If LB(5) = 26 Then B(26) = B(26) + 1
If LB(6) = 26 Then B(26) = B(26) + 1
If LB(1) = 27 Then B(27) = B(27) + 1
If LB(2) = 27 Then B(27) = B(27) + 1
If LB(3) = 27 Then B(27) = B(27) + 1
If LB(4) = 27 Then B(27) = B(27) + 1
If LB(5) = 27 Then B(27) = B(27) + 1
If LB(6) = 27 Then B(27) = B(27) + 1
If LB(1) = 28 Then B(28) = B(28) + 1
If LB(2) = 28 Then B(28) = B(28) + 1
If LB(3) = 28 Then B(28) = B(28) + 1
If LB(4) = 28 Then B(28) = B(28) + 1
If LB(5) = 28 Then B(28) = B(28) + 1
If LB(6) = 28 Then B(28) = B(28) + 1
If LB(1) = 29 Then B(29) = B(29) + 1
If LB(2) = 29 Then B(29) = B(29) + 1
If LB(3) = 29 Then B(29) = B(29) + 1
If LB(4) = 29 Then B(29) = B(29) + 1
If LB(5) = 29 Then B(29) = B(29) + 1
If LB(6) = 29 Then B(29) = B(29) + 1
If LB(1) = 30 Then B(30) = B(30) + 1
If LB(2) = 30 Then B(30) = B(30) + 1
If LB(3) = 30 Then B(30) = B(30) + 1
If LB(4) = 30 Then B(30) = B(30) + 1
If LB(5) = 30 Then B(30) = B(30) + 1
If LB(6) = 30 Then B(30) = B(30) + 1
If LB(1) = 31 Then B(31) = B(31) + 1
If LB(2) = 31 Then B(31) = B(31) + 1
If LB(3) = 31 Then B(31) = B(31) + 1
If LB(4) = 31 Then B(31) = B(31) + 1
If LB(5) = 31 Then B(31) = B(31) + 1
If LB(6) = 31 Then B(31) = B(31) + 1
If LB(1) = 32 Then B(32) = B(32) + 1
If LB(2) = 32 Then B(32) = B(32) + 1
If LB(3) = 32 Then B(32) = B(32) + 1
If LB(4) = 32 Then B(32) = B(32) + 1
If LB(5) = 32 Then B(32) = B(32) + 1
If LB(6) = 32 Then B(32) = B(32) + 1
If LB(1) = 33 Then B(33) = B(33) + 1
If LB(2) = 33 Then B(33) = B(33) + 1
If LB(3) = 33 Then B(33) = B(33) + 1
If LB(4) = 33 Then B(33) = B(33) + 1
If LB(5) = 33 Then B(33) = B(33) + 1
If LB(6) = 33 Then B(33) = B(33) + 1
If LB(1) = 34 Then B(34) = B(34) + 1
If LB(2) = 34 Then B(34) = B(34) + 1
If LB(3) = 34 Then B(34) = B(34) + 1
If LB(4) = 34 Then B(34) = B(34) + 1
If LB(5) = 34 Then B(34) = B(34) + 1
If LB(6) = 34 Then B(34) = B(34) + 1
If LB(1) = 35 Then B(35) = B(35) + 1
If LB(2) = 35 Then B(35) = B(35) + 1
If LB(3) = 35 Then B(35) = B(35) + 1
If LB(4) = 35 Then B(35) = B(35) + 1
If LB(5) = 35 Then B(35) = B(35) + 1
If LB(6) = 35 Then B(35) = B(35) + 1
If LB(1) = 36 Then B(36) = B(36) + 1
If LB(2) = 36 Then B(36) = B(36) + 1
If LB(3) = 36 Then B(36) = B(36) + 1
If LB(4) = 36 Then B(36) = B(36) + 1
If LB(5) = 36 Then B(36) = B(36) + 1
If LB(6) = 36 Then B(36) = B(36) + 1
If LB(1) = 37 Then B(37) = B(37) + 1
If LB(2) = 37 Then B(37) = B(37) + 1
If LB(3) = 37 Then B(37) = B(37) + 1
If LB(4) = 37 Then B(37) = B(37) + 1
If LB(5) = 37 Then B(37) = B(37) + 1
If LB(6) = 37 Then B(37) = B(37) + 1
If LB(1) = 38 Then B(38) = B(38) + 1
If LB(2) = 38 Then B(38) = B(38) + 1
If LB(3) = 38 Then B(38) = B(38) + 1
If LB(4) = 38 Then B(38) = B(38) + 1
If LB(5) = 38 Then B(38) = B(38) + 1
If LB(6) = 38 Then B(38) = B(38) + 1
If LB(1) = 39 Then B(39) = B(39) + 1
If LB(2) = 39 Then B(39) = B(39) + 1
If LB(3) = 39 Then B(39) = B(39) + 1
If LB(4) = 39 Then B(39) = B(39) + 1
If LB(5) = 39 Then B(39) = B(39) + 1
If LB(6) = 39 Then B(39) = B(39) + 1
If LB(1) = 40 Then B(40) = B(40) + 1
If LB(2) = 40 Then B(40) = B(40) + 1
If LB(3) = 40 Then B(40) = B(40) + 1
If LB(4) = 40 Then B(40) = B(40) + 1
If LB(5) = 40 Then B(40) = B(40) + 1
If LB(6) = 40 Then B(40) = B(40) + 1
If LB(1) = 41 Then B(41) = B(41) + 1
If LB(2) = 41 Then B(41) = B(41) + 1
If LB(3) = 41 Then B(41) = B(41) + 1
If LB(4) = 41 Then B(41) = B(41) + 1
If LB(5) = 41 Then B(41) = B(41) + 1
If LB(6) = 41 Then B(41) = B(41) + 1
If LB(1) = 42 Then B(42) = B(42) + 1
If LB(2) = 42 Then B(42) = B(42) + 1
If LB(3) = 42 Then B(42) = B(42) + 1
If LB(4) = 42 Then B(42) = B(42) + 1
If LB(5) = 42 Then B(42) = B(42) + 1
If LB(6) = 42 Then B(42) = B(42) + 1
If LB(1) = 43 Then B(43) = B(43) + 1
If LB(2) = 43 Then B(43) = B(43) + 1
If LB(3) = 43 Then B(43) = B(43) + 1
If LB(4) = 43 Then B(43) = B(43) + 1
If LB(5) = 43 Then B(43) = B(43) + 1
If LB(6) = 43 Then B(43) = B(43) + 1
If LB(1) = 44 Then B(44) = B(44) + 1
If LB(2) = 44 Then B(44) = B(44) + 1
If LB(3) = 44 Then B(44) = B(44) + 1
If LB(4) = 44 Then B(44) = B(44) + 1
If LB(5) = 44 Then B(44) = B(44) + 1
If LB(6) = 44 Then B(44) = B(44) + 1
If LB(1) = 45 Then B(45) = B(45) + 1
If LB(2) = 45 Then B(45) = B(45) + 1
If LB(3) = 45 Then B(45) = B(45) + 1
If LB(4) = 45 Then B(45) = B(45) + 1
If LB(5) = 45 Then B(45) = B(45) + 1
If LB(6) = 45 Then B(45) = B(45) + 1
If LB(1) = 46 Then B(46) = B(46) + 1
If LB(2) = 46 Then B(46) = B(46) + 1
If LB(3) = 46 Then B(46) = B(46) + 1
If LB(4) = 46 Then B(46) = B(46) + 1
If LB(5) = 46 Then B(46) = B(46) + 1
If LB(6) = 46 Then B(46) = B(46) + 1
If LB(1) = 47 Then B(47) = B(47) + 1
If LB(2) = 47 Then B(47) = B(47) + 1
If LB(3) = 47 Then B(47) = B(47) + 1
If LB(4) = 47 Then B(47) = B(47) + 1
If LB(5) = 47 Then B(47) = B(47) + 1
If LB(6) = 47 Then B(47) = B(47) + 1
If LB(1) = 48 Then B(48) = B(48) + 1
If LB(2) = 48 Then B(48) = B(48) + 1
If LB(3) = 48 Then B(48) = B(48) + 1
If LB(4) = 48 Then B(48) = B(48) + 1
If LB(5) = 48 Then B(48) = B(48) + 1
If LB(6) = 48 Then B(48) = B(48) + 1
If LB(1) = 49 Then B(49) = B(49) + 1
If LB(2) = 49 Then B(49) = B(49) + 1
If LB(3) = 49 Then B(49) = B(49) + 1
If LB(4) = 49 Then B(49) = B(49) + 1
If LB(5) = 49 Then B(49) = B(49) + 1
If LB(6) = 49 Then B(49) = B(49) + 1
NumCounter2
End Function
Public Function Order()
M = 1: L = 1
If B(1) <= B(2) Then L = L + 1
If B(1) <= B(3) Then L = L + 1
If B(1) <= B(4) Then L = L + 1
If B(1) <= B(5) Then L = L + 1
If B(1) <= B(6) Then L = L + 1
If B(1) <= B(7) Then L = L + 1
If B(1) <= B(8) Then L = L + 1
If B(1) <= B(9) Then L = L + 1
If B(1) <= B(10) Then L = L + 1
If B(1) <= B(11) Then L = L + 1
If B(1) <= B(12) Then L = L + 1
If B(1) <= B(13) Then L = L + 1
If B(1) <= B(14) Then L = L + 1
If B(1) <= B(15) Then L = L + 1
If B(1) <= B(16) Then L = L + 1
If B(1) <= B(17) Then L = L + 1
If B(1) <= B(18) Then L = L + 1
If B(1) <= B(19) Then L = L + 1
If B(1) <= B(20) Then L = L + 1
If B(1) <= B(21) Then L = L + 1
If B(1) <= B(22) Then L = L + 1
If B(1) <= B(23) Then L = L + 1
If B(1) <= B(24) Then L = L + 1
If B(1) <= B(25) Then L = L + 1
If B(1) <= B(26) Then L = L + 1
If B(1) <= B(27) Then L = L + 1
If B(1) <= B(28) Then L = L + 1
If B(1) <= B(29) Then L = L + 1
If B(1) <= B(30) Then L = L + 1
If B(1) <= B(31) Then L = L + 1
If B(1) <= B(32) Then L = L + 1
If B(1) <= B(33) Then L = L + 1
If B(1) <= B(34) Then L = L + 1
If B(1) <= B(35) Then L = L + 1
If B(1) <= B(36) Then L = L + 1
If B(1) <= B(37) Then L = L + 1
If B(1) <= B(38) Then L = L + 1
If B(1) <= B(39) Then L = L + 1
If B(1) <= B(40) Then L = L + 1
If B(1) <= B(41) Then L = L + 1
If B(1) <= B(42) Then L = L + 1
If B(1) <= B(43) Then L = L + 1
If B(1) <= B(44) Then L = L + 1
If B(1) <= B(45) Then L = L + 1
If B(1) <= B(46) Then L = L + 1
If B(1) <= B(47) Then L = L + 1
If B(1) <= B(48) Then L = L + 1
If B(1) <= B(49) Then L = L + 1
 orderdisp
M = 2: L = 1
If B(2) < B(1) Then L = L + 1
If B(2) <= B(3) Then L = L + 1
If B(2) <= B(4) Then L = L + 1
If B(2) <= B(5) Then L = L + 1
If B(2) <= B(6) Then L = L + 1
If B(2) <= B(7) Then L = L + 1
If B(2) <= B(8) Then L = L + 1
If B(2) <= B(9) Then L = L + 1
If B(2) <= B(10) Then L = L + 1
If B(2) <= B(11) Then L = L + 1
If B(2) <= B(12) Then L = L + 1
If B(2) <= B(13) Then L = L + 1
If B(2) <= B(14) Then L = L + 1
If B(2) <= B(15) Then L = L + 1
If B(2) <= B(16) Then L = L + 1
If B(2) <= B(17) Then L = L + 1
If B(2) <= B(18) Then L = L + 1
If B(2) <= B(19) Then L = L + 1
If B(2) <= B(20) Then L = L + 1
If B(2) <= B(21) Then L = L + 1
If B(2) <= B(22) Then L = L + 1
If B(2) <= B(23) Then L = L + 1
If B(2) <= B(24) Then L = L + 1
If B(2) <= B(25) Then L = L + 1
If B(2) <= B(26) Then L = L + 1
If B(2) <= B(27) Then L = L + 1
If B(2) <= B(28) Then L = L + 1
If B(2) <= B(29) Then L = L + 1
If B(2) <= B(30) Then L = L + 1
If B(2) <= B(31) Then L = L + 1
If B(2) <= B(32) Then L = L + 1
If B(2) <= B(33) Then L = L + 1
If B(2) <= B(34) Then L = L + 1
If B(2) <= B(35) Then L = L + 1
If B(2) <= B(36) Then L = L + 1
If B(2) <= B(37) Then L = L + 1
If B(2) <= B(38) Then L = L + 1
If B(2) <= B(39) Then L = L + 1
If B(2) <= B(40) Then L = L + 1
If B(2) <= B(41) Then L = L + 1
If B(2) <= B(42) Then L = L + 1
If B(2) <= B(43) Then L = L + 1
If B(2) <= B(44) Then L = L + 1
If B(2) <= B(45) Then L = L + 1
If B(2) <= B(46) Then L = L + 1
If B(2) <= B(47) Then L = L + 1
If B(2) <= B(48) Then L = L + 1
If B(2) <= B(49) Then L = L + 1
 orderdisp
M = 3: L = 1
If B(3) < B(1) Then L = L + 1
If B(3) < B(2) Then L = L + 1
If B(3) <= B(4) Then L = L + 1
If B(3) <= B(5) Then L = L + 1
If B(3) <= B(6) Then L = L + 1
If B(3) <= B(7) Then L = L + 1
If B(3) <= B(8) Then L = L + 1
If B(3) <= B(9) Then L = L + 1
If B(3) <= B(10) Then L = L + 1
If B(3) <= B(11) Then L = L + 1
If B(3) <= B(12) Then L = L + 1
If B(3) <= B(13) Then L = L + 1
If B(3) <= B(14) Then L = L + 1
If B(3) <= B(15) Then L = L + 1
If B(3) <= B(16) Then L = L + 1
If B(3) <= B(17) Then L = L + 1
If B(3) <= B(18) Then L = L + 1
If B(3) <= B(19) Then L = L + 1
If B(3) <= B(20) Then L = L + 1
If B(3) <= B(21) Then L = L + 1
If B(3) <= B(22) Then L = L + 1
If B(3) <= B(23) Then L = L + 1
If B(3) <= B(24) Then L = L + 1
If B(3) <= B(25) Then L = L + 1
If B(3) <= B(26) Then L = L + 1
If B(3) <= B(27) Then L = L + 1
If B(3) <= B(28) Then L = L + 1
If B(3) <= B(29) Then L = L + 1
If B(3) <= B(30) Then L = L + 1
If B(3) <= B(31) Then L = L + 1
If B(3) <= B(32) Then L = L + 1
If B(3) <= B(33) Then L = L + 1
If B(3) <= B(34) Then L = L + 1
If B(3) <= B(35) Then L = L + 1
If B(3) <= B(36) Then L = L + 1
If B(3) <= B(37) Then L = L + 1
If B(3) <= B(38) Then L = L + 1
If B(3) <= B(39) Then L = L + 1
If B(3) <= B(40) Then L = L + 1
If B(3) <= B(41) Then L = L + 1
If B(3) <= B(42) Then L = L + 1
If B(3) <= B(43) Then L = L + 1
If B(3) <= B(44) Then L = L + 1
If B(3) <= B(45) Then L = L + 1
If B(3) <= B(46) Then L = L + 1
If B(3) <= B(47) Then L = L + 1
If B(3) <= B(48) Then L = L + 1
If B(3) <= B(49) Then L = L + 1
 orderdisp
M = 4: L = 1
If B(4) < B(1) Then L = L + 1
If B(4) < B(2) Then L = L + 1
If B(4) < B(3) Then L = L + 1
If B(4) <= B(5) Then L = L + 1
If B(4) <= B(6) Then L = L + 1
If B(4) <= B(7) Then L = L + 1
If B(4) <= B(8) Then L = L + 1
If B(4) <= B(9) Then L = L + 1
If B(4) <= B(10) Then L = L + 1
If B(4) <= B(11) Then L = L + 1
If B(4) <= B(12) Then L = L + 1
If B(4) <= B(13) Then L = L + 1
If B(4) <= B(14) Then L = L + 1
If B(4) <= B(15) Then L = L + 1
If B(4) <= B(16) Then L = L + 1
If B(4) <= B(17) Then L = L + 1
If B(4) <= B(18) Then L = L + 1
If B(4) <= B(19) Then L = L + 1
If B(4) <= B(20) Then L = L + 1
If B(4) <= B(21) Then L = L + 1
If B(4) <= B(22) Then L = L + 1
If B(4) <= B(23) Then L = L + 1
If B(1) <= B(24) Then L = L + 1
If B(4) <= B(25) Then L = L + 1
If B(4) <= B(26) Then L = L + 1
If B(4) <= B(27) Then L = L + 1
If B(4) <= B(28) Then L = L + 1
If B(4) <= B(29) Then L = L + 1
If B(4) <= B(30) Then L = L + 1
If B(4) <= B(31) Then L = L + 1
If B(4) <= B(32) Then L = L + 1
If B(4) <= B(33) Then L = L + 1
If B(4) <= B(34) Then L = L + 1
If B(4) <= B(35) Then L = L + 1
If B(4) <= B(36) Then L = L + 1
If B(4) <= B(37) Then L = L + 1
If B(4) <= B(38) Then L = L + 1
If B(4) <= B(39) Then L = L + 1
If B(4) <= B(40) Then L = L + 1
If B(4) <= B(41) Then L = L + 1
If B(4) <= B(42) Then L = L + 1
If B(4) <= B(43) Then L = L + 1
If B(4) <= B(44) Then L = L + 1
If B(4) <= B(45) Then L = L + 1
If B(4) <= B(46) Then L = L + 1
If B(4) <= B(47) Then L = L + 1
If B(4) <= B(48) Then L = L + 1
If B(4) <= B(49) Then L = L + 1
 orderdisp
M = 5: L = 1
If B(5) < B(1) Then L = L + 1
If B(5) < B(2) Then L = L + 1
If B(5) < B(3) Then L = L + 1
If B(5) < B(4) Then L = L + 1
If B(5) <= B(6) Then L = L + 1
If B(5) <= B(7) Then L = L + 1
If B(5) <= B(8) Then L = L + 1
If B(5) <= B(9) Then L = L + 1
If B(5) <= B(10) Then L = L + 1
If B(5) <= B(11) Then L = L + 1
If B(5) <= B(12) Then L = L + 1
If B(5) <= B(13) Then L = L + 1
If B(5) <= B(14) Then L = L + 1
If B(5) <= B(15) Then L = L + 1
If B(5) <= B(16) Then L = L + 1
If B(5) <= B(17) Then L = L + 1
If B(5) <= B(18) Then L = L + 1
If B(5) <= B(19) Then L = L + 1
If B(5) <= B(20) Then L = L + 1
If B(5) <= B(21) Then L = L + 1
If B(5) <= B(22) Then L = L + 1
If B(5) <= B(23) Then L = L + 1
If B(5) <= B(24) Then L = L + 1
If B(5) <= B(25) Then L = L + 1
If B(5) <= B(26) Then L = L + 1
If B(5) <= B(27) Then L = L + 1
If B(5) <= B(28) Then L = L + 1
If B(5) <= B(29) Then L = L + 1
If B(5) <= B(30) Then L = L + 1
If B(5) <= B(31) Then L = L + 1
If B(5) <= B(32) Then L = L + 1
If B(5) <= B(33) Then L = L + 1
If B(5) <= B(34) Then L = L + 1
If B(5) <= B(35) Then L = L + 1
If B(5) <= B(36) Then L = L + 1
If B(5) <= B(37) Then L = L + 1
If B(5) <= B(38) Then L = L + 1
If B(5) <= B(39) Then L = L + 1
If B(5) <= B(40) Then L = L + 1
If B(5) <= B(41) Then L = L + 1
If B(5) <= B(42) Then L = L + 1
If B(5) <= B(43) Then L = L + 1
If B(5) <= B(44) Then L = L + 1
If B(5) <= B(45) Then L = L + 1
If B(5) <= B(46) Then L = L + 1
If B(5) <= B(47) Then L = L + 1
If B(5) <= B(48) Then L = L + 1
If B(5) <= B(49) Then L = L + 1
 orderdisp
M = 6: L = 1
If B(6) < B(1) Then L = L + 1
If B(6) < B(2) Then L = L + 1
If B(6) < B(3) Then L = L + 1
If B(6) < B(4) Then L = L + 1
If B(6) < B(5) Then L = L + 1
If B(6) <= B(7) Then L = L + 1
If B(6) <= B(8) Then L = L + 1
If B(6) <= B(9) Then L = L + 1
If B(6) <= B(10) Then L = L + 1
If B(6) <= B(11) Then L = L + 1
If B(6) <= B(12) Then L = L + 1
If B(6) <= B(13) Then L = L + 1
If B(6) <= B(14) Then L = L + 1
If B(6) <= B(15) Then L = L + 1
If B(6) <= B(16) Then L = L + 1
If B(6) <= B(17) Then L = L + 1
If B(6) <= B(18) Then L = L + 1
If B(6) <= B(19) Then L = L + 1
If B(6) <= B(20) Then L = L + 1
If B(6) <= B(21) Then L = L + 1
If B(6) <= B(22) Then L = L + 1
If B(6) <= B(23) Then L = L + 1
If B(6) <= B(24) Then L = L + 1
If B(6) <= B(25) Then L = L + 1
If B(6) <= B(26) Then L = L + 1
If B(6) <= B(27) Then L = L + 1
If B(6) <= B(28) Then L = L + 1
If B(6) <= B(29) Then L = L + 1
If B(6) <= B(30) Then L = L + 1
If B(6) <= B(31) Then L = L + 1
If B(6) <= B(32) Then L = L + 1
If B(6) <= B(33) Then L = L + 1
If B(6) <= B(34) Then L = L + 1
If B(6) <= B(35) Then L = L + 1
If B(6) <= B(36) Then L = L + 1
If B(6) <= B(37) Then L = L + 1
If B(6) <= B(38) Then L = L + 1
If B(6) <= B(39) Then L = L + 1
If B(6) <= B(40) Then L = L + 1
If B(6) <= B(41) Then L = L + 1
If B(6) <= B(42) Then L = L + 1
If B(6) <= B(43) Then L = L + 1
If B(6) <= B(44) Then L = L + 1
If B(6) <= B(45) Then L = L + 1
If B(6) <= B(46) Then L = L + 1
If B(6) <= B(47) Then L = L + 1
If B(6) <= B(48) Then L = L + 1
If B(6) <= B(49) Then L = L + 1
 orderdisp
M = 7: L = 1
If B(7) < B(1) Then L = L + 1
If B(7) < B(2) Then L = L + 1
If B(7) < B(3) Then L = L + 1
If B(7) < B(4) Then L = L + 1
If B(7) < B(5) Then L = L + 1
If B(7) < B(6) Then L = L + 1
If B(7) <= B(8) Then L = L + 1
If B(7) <= B(9) Then L = L + 1
If B(7) <= B(10) Then L = L + 1
If B(7) <= B(11) Then L = L + 1
If B(7) <= B(12) Then L = L + 1
If B(7) <= B(13) Then L = L + 1
If B(7) <= B(14) Then L = L + 1
If B(7) <= B(15) Then L = L + 1
If B(7) <= B(16) Then L = L + 1
If B(7) <= B(17) Then L = L + 1
If B(7) <= B(18) Then L = L + 1
If B(7) <= B(19) Then L = L + 1
If B(7) <= B(20) Then L = L + 1
If B(7) <= B(21) Then L = L + 1
If B(7) <= B(22) Then L = L + 1
If B(7) <= B(23) Then L = L + 1
If B(7) <= B(24) Then L = L + 1
If B(7) <= B(25) Then L = L + 1
If B(7) <= B(26) Then L = L + 1
If B(7) <= B(27) Then L = L + 1
If B(7) <= B(28) Then L = L + 1
If B(7) <= B(29) Then L = L + 1
If B(7) <= B(30) Then L = L + 1
If B(7) <= B(31) Then L = L + 1
If B(7) <= B(32) Then L = L + 1
If B(7) <= B(33) Then L = L + 1
If B(7) <= B(34) Then L = L + 1
If B(7) <= B(35) Then L = L + 1
If B(7) <= B(36) Then L = L + 1
If B(7) <= B(37) Then L = L + 1
If B(7) <= B(38) Then L = L + 1
If B(7) <= B(39) Then L = L + 1
If B(7) <= B(40) Then L = L + 1
If B(7) <= B(41) Then L = L + 1
If B(7) <= B(42) Then L = L + 1
If B(7) <= B(43) Then L = L + 1
If B(7) <= B(44) Then L = L + 1
If B(7) <= B(45) Then L = L + 1
If B(7) <= B(46) Then L = L + 1
If B(7) <= B(47) Then L = L + 1
If B(7) <= B(48) Then L = L + 1
If B(7) <= B(49) Then L = L + 1
 orderdisp
M = 8: L = 1
If B(8) < B(1) Then L = L + 1
If B(8) < B(2) Then L = L + 1
If B(8) < B(3) Then L = L + 1
If B(8) < B(4) Then L = L + 1
If B(8) < B(5) Then L = L + 1
If B(8) < B(6) Then L = L + 1
If B(8) < B(7) Then L = L + 1
If B(8) <= B(9) Then L = L + 1
If B(8) <= B(10) Then L = L + 1
If B(8) <= B(11) Then L = L + 1
If B(8) <= B(12) Then L = L + 1
If B(8) <= B(13) Then L = L + 1
If B(8) <= B(14) Then L = L + 1
If B(8) <= B(15) Then L = L + 1
If B(8) <= B(16) Then L = L + 1
If B(8) <= B(17) Then L = L + 1
If B(8) <= B(18) Then L = L + 1
If B(8) <= B(19) Then L = L + 1
If B(8) <= B(20) Then L = L + 1
If B(8) <= B(21) Then L = L + 1
If B(8) <= B(22) Then L = L + 1
If B(8) <= B(23) Then L = L + 1
If B(8) <= B(24) Then L = L + 1
If B(8) <= B(25) Then L = L + 1
If B(8) <= B(26) Then L = L + 1
If B(8) <= B(27) Then L = L + 1
If B(8) <= B(28) Then L = L + 1
If B(8) <= B(29) Then L = L + 1
If B(8) <= B(30) Then L = L + 1
If B(8) <= B(31) Then L = L + 1
If B(8) <= B(32) Then L = L + 1
If B(8) <= B(33) Then L = L + 1
If B(8) <= B(34) Then L = L + 1
If B(8) <= B(35) Then L = L + 1
If B(8) <= B(36) Then L = L + 1
If B(8) <= B(37) Then L = L + 1
If B(8) <= B(38) Then L = L + 1
If B(8) <= B(39) Then L = L + 1
If B(8) <= B(40) Then L = L + 1
If B(8) <= B(41) Then L = L + 1
If B(8) <= B(42) Then L = L + 1
If B(8) <= B(43) Then L = L + 1
If B(8) <= B(44) Then L = L + 1
If B(8) <= B(45) Then L = L + 1
If B(8) <= B(46) Then L = L + 1
If B(8) <= B(47) Then L = L + 1
If B(8) <= B(48) Then L = L + 1
If B(8) <= B(49) Then L = L + 1
 orderdisp
M = 9: L = 1
If B(9) < B(1) Then L = L + 1
If B(9) < B(2) Then L = L + 1
If B(9) < B(3) Then L = L + 1
If B(9) < B(4) Then L = L + 1
If B(9) < B(5) Then L = L + 1
If B(9) < B(6) Then L = L + 1
If B(9) < B(7) Then L = L + 1
If B(9) < B(8) Then L = L + 1
If B(9) <= B(10) Then L = L + 1
If B(9) <= B(11) Then L = L + 1
If B(9) <= B(12) Then L = L + 1
If B(9) <= B(13) Then L = L + 1
If B(9) <= B(14) Then L = L + 1
If B(9) <= B(15) Then L = L + 1
If B(9) <= B(16) Then L = L + 1
If B(9) <= B(17) Then L = L + 1
If B(9) <= B(18) Then L = L + 1
If B(9) <= B(19) Then L = L + 1
If B(9) <= B(20) Then L = L + 1
If B(9) <= B(21) Then L = L + 1
If B(9) <= B(22) Then L = L + 1
If B(9) <= B(23) Then L = L + 1
If B(9) <= B(24) Then L = L + 1
If B(9) <= B(25) Then L = L + 1
If B(9) <= B(26) Then L = L + 1
If B(9) <= B(27) Then L = L + 1
If B(9) <= B(28) Then L = L + 1
If B(9) <= B(29) Then L = L + 1
If B(9) <= B(30) Then L = L + 1
If B(9) <= B(31) Then L = L + 1
If B(9) <= B(32) Then L = L + 1
If B(9) <= B(33) Then L = L + 1
If B(9) <= B(34) Then L = L + 1
If B(9) <= B(35) Then L = L + 1
If B(9) <= B(36) Then L = L + 1
If B(9) <= B(37) Then L = L + 1
If B(9) <= B(38) Then L = L + 1
If B(9) <= B(39) Then L = L + 1
If B(9) <= B(40) Then L = L + 1
If B(9) <= B(41) Then L = L + 1
If B(9) <= B(42) Then L = L + 1
If B(9) <= B(43) Then L = L + 1
If B(9) <= B(44) Then L = L + 1
If B(9) <= B(45) Then L = L + 1
If B(9) <= B(46) Then L = L + 1
If B(9) <= B(47) Then L = L + 1
If B(9) <= B(48) Then L = L + 1
If B(9) <= B(49) Then L = L + 1
 orderdisp
M = 10: L = 1
If B(10) < B(1) Then L = L + 1
If B(10) < B(2) Then L = L + 1
If B(10) < B(3) Then L = L + 1
If B(10) < B(4) Then L = L + 1
If B(10) < B(5) Then L = L + 1
If B(10) < B(6) Then L = L + 1
If B(10) < B(7) Then L = L + 1
If B(10) < B(8) Then L = L + 1
If B(10) < B(9) Then L = L + 1
If B(10) <= B(11) Then L = L + 1
If B(10) <= B(12) Then L = L + 1
If B(10) <= B(13) Then L = L + 1
If B(10) <= B(14) Then L = L + 1
If B(10) <= B(15) Then L = L + 1
If B(10) <= B(16) Then L = L + 1
If B(10) <= B(17) Then L = L + 1
If B(10) <= B(18) Then L = L + 1
If B(10) <= B(19) Then L = L + 1
If B(10) <= B(20) Then L = L + 1
If B(10) <= B(21) Then L = L + 1
If B(10) <= B(22) Then L = L + 1
If B(10) <= B(23) Then L = L + 1
If B(10) <= B(24) Then L = L + 1
If B(10) <= B(25) Then L = L + 1
If B(10) <= B(26) Then L = L + 1
If B(10) <= B(27) Then L = L + 1
If B(10) <= B(28) Then L = L + 1
If B(10) <= B(29) Then L = L + 1
If B(10) <= B(30) Then L = L + 1
If B(10) <= B(31) Then L = L + 1
If B(10) <= B(32) Then L = L + 1
If B(10) <= B(33) Then L = L + 1
If B(10) <= B(34) Then L = L + 1
If B(10) <= B(35) Then L = L + 1
If B(10) <= B(36) Then L = L + 1
If B(10) <= B(37) Then L = L + 1
If B(10) <= B(38) Then L = L + 1
If B(10) <= B(39) Then L = L + 1
If B(10) <= B(40) Then L = L + 1
If B(10) <= B(41) Then L = L + 1
If B(10) <= B(42) Then L = L + 1
If B(10) <= B(43) Then L = L + 1
If B(10) <= B(44) Then L = L + 1
If B(10) <= B(45) Then L = L + 1
If B(10) <= B(46) Then L = L + 1
If B(10) <= B(47) Then L = L + 1
If B(10) <= B(48) Then L = L + 1
If B(10) <= B(49) Then L = L + 1
 orderdisp
M = 11: L = 1
If B(11) < B(1) Then L = L + 1
If B(11) < B(2) Then L = L + 1
If B(11) < B(3) Then L = L + 1
If B(11) < B(4) Then L = L + 1
If B(11) < B(5) Then L = L + 1
If B(11) < B(6) Then L = L + 1
If B(11) < B(7) Then L = L + 1
If B(11) < B(8) Then L = L + 1
If B(11) < B(9) Then L = L + 1
If B(11) < B(10) Then L = L + 1
If B(11) <= B(12) Then L = L + 1
If B(11) <= B(13) Then L = L + 1
If B(11) <= B(14) Then L = L + 1
If B(11) <= B(15) Then L = L + 1
If B(11) <= B(16) Then L = L + 1
If B(11) <= B(17) Then L = L + 1
If B(11) <= B(18) Then L = L + 1
If B(11) <= B(19) Then L = L + 1
If B(11) <= B(20) Then L = L + 1
If B(11) <= B(21) Then L = L + 1
If B(11) <= B(22) Then L = L + 1
If B(11) <= B(23) Then L = L + 1
If B(11) <= B(24) Then L = L + 1
If B(11) <= B(25) Then L = L + 1
If B(11) <= B(26) Then L = L + 1
If B(11) <= B(27) Then L = L + 1
If B(11) <= B(28) Then L = L + 1
If B(11) <= B(29) Then L = L + 1
If B(11) <= B(30) Then L = L + 1
If B(11) <= B(31) Then L = L + 1
If B(11) <= B(32) Then L = L + 1
If B(11) <= B(33) Then L = L + 1
If B(11) <= B(34) Then L = L + 1
If B(11) <= B(35) Then L = L + 1
If B(11) <= B(36) Then L = L + 1
If B(11) <= B(37) Then L = L + 1
If B(11) <= B(38) Then L = L + 1
If B(11) <= B(39) Then L = L + 1
If B(11) <= B(40) Then L = L + 1
If B(11) <= B(41) Then L = L + 1
If B(11) <= B(42) Then L = L + 1
If B(11) <= B(43) Then L = L + 1
If B(11) <= B(44) Then L = L + 1
If B(11) <= B(45) Then L = L + 1
If B(11) <= B(46) Then L = L + 1
If B(11) <= B(47) Then L = L + 1
If B(11) <= B(48) Then L = L + 1
If B(11) <= B(49) Then L = L + 1
 orderdisp
M = 12: L = 1
If B(12) < B(1) Then L = L + 1
If B(12) < B(2) Then L = L + 1
If B(12) < B(3) Then L = L + 1
If B(12) < B(4) Then L = L + 1
If B(12) < B(5) Then L = L + 1
If B(12) < B(6) Then L = L + 1
If B(12) < B(7) Then L = L + 1
If B(12) < B(8) Then L = L + 1
If B(12) < B(9) Then L = L + 1
If B(12) < B(10) Then L = L + 1
If B(12) < B(11) Then L = L + 1
If B(12) <= B(13) Then L = L + 1
If B(12) <= B(14) Then L = L + 1
If B(12) <= B(15) Then L = L + 1
If B(12) <= B(16) Then L = L + 1
If B(12) <= B(17) Then L = L + 1
If B(12) <= B(18) Then L = L + 1
If B(12) <= B(19) Then L = L + 1
If B(12) <= B(20) Then L = L + 1
If B(12) <= B(21) Then L = L + 1
If B(12) <= B(22) Then L = L + 1
If B(12) <= B(23) Then L = L + 1
If B(12) <= B(24) Then L = L + 1
If B(12) <= B(25) Then L = L + 1
If B(12) <= B(26) Then L = L + 1
If B(12) <= B(27) Then L = L + 1
If B(12) <= B(28) Then L = L + 1
If B(12) <= B(29) Then L = L + 1
If B(12) <= B(30) Then L = L + 1
If B(12) <= B(31) Then L = L + 1
If B(12) <= B(32) Then L = L + 1
If B(12) <= B(33) Then L = L + 1
If B(12) <= B(34) Then L = L + 1
If B(12) <= B(35) Then L = L + 1
If B(12) <= B(36) Then L = L + 1
If B(12) <= B(37) Then L = L + 1
If B(12) <= B(38) Then L = L + 1
If B(12) <= B(39) Then L = L + 1
If B(12) <= B(40) Then L = L + 1
If B(12) <= B(41) Then L = L + 1
If B(12) <= B(42) Then L = L + 1
If B(12) <= B(43) Then L = L + 1
If B(12) <= B(44) Then L = L + 1
If B(12) <= B(45) Then L = L + 1
If B(12) <= B(46) Then L = L + 1
If B(12) <= B(47) Then L = L + 1
If B(12) <= B(48) Then L = L + 1
If B(12) <= B(49) Then L = L + 1
 orderdisp
M = 13: L = 1
If B(13) < B(1) Then L = L + 1
If B(13) < B(2) Then L = L + 1
If B(13) < B(3) Then L = L + 1
If B(13) < B(4) Then L = L + 1
If B(13) < B(5) Then L = L + 1
If B(13) < B(6) Then L = L + 1
If B(13) < B(7) Then L = L + 1
If B(13) < B(8) Then L = L + 1
If B(13) < B(9) Then L = L + 1
If B(13) < B(10) Then L = L + 1
If B(13) < B(11) Then L = L + 1
If B(13) < B(12) Then L = L + 1
If B(13) <= B(14) Then L = L + 1
If B(13) <= B(15) Then L = L + 1
If B(13) <= B(16) Then L = L + 1
If B(13) <= B(17) Then L = L + 1
If B(13) <= B(18) Then L = L + 1
If B(13) <= B(19) Then L = L + 1
If B(13) <= B(20) Then L = L + 1
If B(13) <= B(21) Then L = L + 1
If B(13) <= B(22) Then L = L + 1
If B(13) <= B(23) Then L = L + 1
If B(13) <= B(24) Then L = L + 1
If B(13) <= B(25) Then L = L + 1
If B(13) <= B(26) Then L = L + 1
If B(13) <= B(27) Then L = L + 1
If B(13) <= B(28) Then L = L + 1
If B(13) <= B(29) Then L = L + 1
If B(13) <= B(30) Then L = L + 1
If B(13) <= B(31) Then L = L + 1
If B(13) <= B(32) Then L = L + 1
If B(13) <= B(33) Then L = L + 1
If B(13) <= B(34) Then L = L + 1
If B(13) <= B(35) Then L = L + 1
If B(13) <= B(36) Then L = L + 1
If B(13) <= B(37) Then L = L + 1
If B(13) <= B(38) Then L = L + 1
If B(13) <= B(39) Then L = L + 1
If B(13) <= B(40) Then L = L + 1
If B(13) <= B(41) Then L = L + 1
If B(13) <= B(42) Then L = L + 1
If B(13) <= B(43) Then L = L + 1
If B(13) <= B(44) Then L = L + 1
If B(13) <= B(45) Then L = L + 1
If B(13) <= B(46) Then L = L + 1
If B(13) <= B(47) Then L = L + 1
If B(13) <= B(48) Then L = L + 1
If B(13) <= B(49) Then L = L + 1
 orderdisp
M = 14: L = 1
If B(14) < B(1) Then L = L + 1
If B(14) < B(2) Then L = L + 1
If B(14) < B(3) Then L = L + 1
If B(14) < B(4) Then L = L + 1
If B(14) < B(5) Then L = L + 1
If B(14) < B(6) Then L = L + 1
If B(14) < B(7) Then L = L + 1
If B(14) < B(8) Then L = L + 1
If B(14) < B(9) Then L = L + 1
If B(14) < B(10) Then L = L + 1
If B(14) < B(11) Then L = L + 1
If B(14) < B(12) Then L = L + 1
If B(14) < B(13) Then L = L + 1
If B(14) <= B(15) Then L = L + 1
If B(14) <= B(16) Then L = L + 1
If B(14) <= B(17) Then L = L + 1
If B(14) <= B(18) Then L = L + 1
If B(14) <= B(19) Then L = L + 1
If B(14) <= B(20) Then L = L + 1
If B(14) <= B(21) Then L = L + 1
If B(14) <= B(22) Then L = L + 1
If B(14) <= B(23) Then L = L + 1
If B(14) <= B(24) Then L = L + 1
If B(14) <= B(25) Then L = L + 1
If B(14) <= B(26) Then L = L + 1
If B(14) <= B(27) Then L = L + 1
If B(14) <= B(28) Then L = L + 1
If B(14) <= B(29) Then L = L + 1
If B(14) <= B(30) Then L = L + 1
If B(14) <= B(31) Then L = L + 1
If B(14) <= B(32) Then L = L + 1
If B(14) <= B(33) Then L = L + 1
If B(14) <= B(34) Then L = L + 1
If B(14) <= B(35) Then L = L + 1
If B(14) <= B(36) Then L = L + 1
If B(14) <= B(37) Then L = L + 1
If B(14) <= B(38) Then L = L + 1
If B(14) <= B(39) Then L = L + 1
If B(14) <= B(40) Then L = L + 1
If B(14) <= B(41) Then L = L + 1
If B(14) <= B(42) Then L = L + 1
If B(14) <= B(43) Then L = L + 1
If B(14) <= B(44) Then L = L + 1
If B(14) <= B(45) Then L = L + 1
If B(14) <= B(46) Then L = L + 1
If B(14) <= B(47) Then L = L + 1
If B(14) <= B(48) Then L = L + 1
If B(14) <= B(49) Then L = L + 1
 orderdisp
M = 15: L = 1
If B(15) < B(1) Then L = L + 1
If B(15) < B(2) Then L = L + 1
If B(15) < B(3) Then L = L + 1
If B(15) < B(4) Then L = L + 1
If B(15) < B(5) Then L = L + 1
If B(15) < B(6) Then L = L + 1
If B(15) < B(7) Then L = L + 1
If B(15) < B(8) Then L = L + 1
If B(15) < B(9) Then L = L + 1
If B(15) < B(10) Then L = L + 1
If B(15) < B(11) Then L = L + 1
If B(15) < B(12) Then L = L + 1
If B(15) < B(13) Then L = L + 1
If B(15) < B(14) Then L = L + 1
If B(15) <= B(16) Then L = L + 1
If B(15) <= B(17) Then L = L + 1
If B(15) <= B(18) Then L = L + 1
If B(15) <= B(19) Then L = L + 1
If B(15) <= B(20) Then L = L + 1
If B(15) <= B(21) Then L = L + 1
If B(15) <= B(22) Then L = L + 1
If B(15) <= B(23) Then L = L + 1
If B(15) <= B(24) Then L = L + 1
If B(15) <= B(25) Then L = L + 1
If B(15) <= B(26) Then L = L + 1
If B(15) <= B(27) Then L = L + 1
If B(15) <= B(28) Then L = L + 1
If B(15) <= B(29) Then L = L + 1
If B(15) <= B(30) Then L = L + 1
If B(15) <= B(31) Then L = L + 1
If B(15) <= B(32) Then L = L + 1
If B(15) <= B(33) Then L = L + 1
If B(15) <= B(34) Then L = L + 1
If B(15) <= B(35) Then L = L + 1
If B(15) <= B(36) Then L = L + 1
If B(15) <= B(37) Then L = L + 1
If B(15) <= B(38) Then L = L + 1
If B(15) <= B(39) Then L = L + 1
If B(15) <= B(40) Then L = L + 1
If B(15) <= B(41) Then L = L + 1
If B(15) <= B(42) Then L = L + 1
If B(15) <= B(43) Then L = L + 1
If B(15) <= B(44) Then L = L + 1
If B(15) <= B(45) Then L = L + 1
If B(15) <= B(46) Then L = L + 1
If B(15) <= B(47) Then L = L + 1
If B(15) <= B(48) Then L = L + 1
If B(15) <= B(49) Then L = L + 1
 orderdisp
M = 16: L = 1
If B(16) < B(1) Then L = L + 1
If B(16) < B(2) Then L = L + 1
If B(16) < B(3) Then L = L + 1
If B(16) < B(4) Then L = L + 1
If B(16) < B(5) Then L = L + 1
If B(16) < B(6) Then L = L + 1
If B(16) < B(7) Then L = L + 1
If B(16) < B(8) Then L = L + 1
If B(16) < B(9) Then L = L + 1
If B(16) < B(10) Then L = L + 1
If B(16) < B(11) Then L = L + 1
If B(16) < B(12) Then L = L + 1
If B(16) < B(13) Then L = L + 1
If B(16) < B(14) Then L = L + 1
If B(16) < B(15) Then L = L + 1
If B(16) <= B(17) Then L = L + 1
If B(16) <= B(18) Then L = L + 1
If B(16) <= B(19) Then L = L + 1
If B(16) <= B(20) Then L = L + 1
If B(16) <= B(21) Then L = L + 1
If B(16) <= B(22) Then L = L + 1
If B(16) <= B(23) Then L = L + 1
If B(16) <= B(24) Then L = L + 1
If B(16) <= B(25) Then L = L + 1
If B(16) <= B(26) Then L = L + 1
If B(16) <= B(27) Then L = L + 1
If B(16) <= B(28) Then L = L + 1
If B(16) <= B(29) Then L = L + 1
If B(16) <= B(30) Then L = L + 1
If B(16) <= B(31) Then L = L + 1
If B(16) <= B(32) Then L = L + 1
If B(16) <= B(33) Then L = L + 1
If B(16) <= B(34) Then L = L + 1
If B(16) <= B(35) Then L = L + 1
If B(16) <= B(36) Then L = L + 1
If B(16) <= B(37) Then L = L + 1
If B(16) <= B(38) Then L = L + 1
If B(16) <= B(39) Then L = L + 1
If B(16) <= B(40) Then L = L + 1
If B(16) <= B(41) Then L = L + 1
If B(16) <= B(42) Then L = L + 1
If B(16) <= B(43) Then L = L + 1
If B(16) <= B(44) Then L = L + 1
If B(16) <= B(45) Then L = L + 1
If B(16) <= B(46) Then L = L + 1
If B(16) <= B(47) Then L = L + 1
If B(16) <= B(48) Then L = L + 1
If B(16) <= B(49) Then L = L + 1
 orderdisp
M = 17: L = 1
If B(17) < B(1) Then L = L + 1
If B(17) < B(2) Then L = L + 1
If B(17) < B(3) Then L = L + 1
If B(17) < B(4) Then L = L + 1
If B(17) < B(5) Then L = L + 1
If B(17) < B(6) Then L = L + 1
If B(17) < B(7) Then L = L + 1
If B(17) < B(8) Then L = L + 1
If B(17) < B(9) Then L = L + 1
If B(17) < B(10) Then L = L + 1
If B(17) < B(11) Then L = L + 1
If B(17) < B(12) Then L = L + 1
If B(17) < B(13) Then L = L + 1
If B(17) < B(14) Then L = L + 1
If B(17) < B(15) Then L = L + 1
If B(17) < B(16) Then L = L + 1
If B(17) <= B(18) Then L = L + 1
If B(17) <= B(19) Then L = L + 1
If B(17) <= B(20) Then L = L + 1
If B(17) <= B(21) Then L = L + 1
If B(17) <= B(22) Then L = L + 1
If B(17) <= B(23) Then L = L + 1
If B(17) <= B(24) Then L = L + 1
If B(17) <= B(25) Then L = L + 1
If B(17) <= B(26) Then L = L + 1
If B(17) <= B(27) Then L = L + 1
If B(17) <= B(28) Then L = L + 1
If B(17) <= B(29) Then L = L + 1
If B(17) <= B(30) Then L = L + 1
If B(17) <= B(31) Then L = L + 1
If B(17) <= B(32) Then L = L + 1
If B(17) <= B(33) Then L = L + 1
If B(17) <= B(34) Then L = L + 1
If B(17) <= B(35) Then L = L + 1
If B(17) <= B(36) Then L = L + 1
If B(17) <= B(37) Then L = L + 1
If B(17) <= B(38) Then L = L + 1
If B(17) <= B(39) Then L = L + 1
If B(17) <= B(40) Then L = L + 1
If B(17) <= B(41) Then L = L + 1
If B(17) <= B(42) Then L = L + 1
If B(17) <= B(43) Then L = L + 1
If B(17) <= B(44) Then L = L + 1
If B(17) <= B(45) Then L = L + 1
If B(17) <= B(46) Then L = L + 1
If B(17) <= B(47) Then L = L + 1
If B(17) <= B(48) Then L = L + 1
If B(17) <= B(49) Then L = L + 1
 orderdisp
M = 18: L = 1
If B(18) < B(1) Then L = L + 1
If B(18) < B(2) Then L = L + 1
If B(18) < B(3) Then L = L + 1
If B(18) < B(4) Then L = L + 1
If B(18) < B(5) Then L = L + 1
If B(18) < B(6) Then L = L + 1
If B(18) < B(7) Then L = L + 1
If B(18) < B(8) Then L = L + 1
If B(18) < B(9) Then L = L + 1
If B(18) < B(10) Then L = L + 1
If B(18) < B(11) Then L = L + 1
If B(18) < B(12) Then L = L + 1
If B(18) < B(13) Then L = L + 1
If B(18) < B(14) Then L = L + 1
If B(18) < B(15) Then L = L + 1
If B(18) < B(16) Then L = L + 1
If B(18) < B(17) Then L = L + 1
If B(18) <= B(19) Then L = L + 1
If B(18) <= B(20) Then L = L + 1
If B(18) <= B(21) Then L = L + 1
If B(18) <= B(22) Then L = L + 1
If B(18) <= B(23) Then L = L + 1
If B(18) <= B(24) Then L = L + 1
If B(18) <= B(25) Then L = L + 1
If B(18) <= B(26) Then L = L + 1
If B(18) <= B(27) Then L = L + 1
If B(18) <= B(28) Then L = L + 1
If B(18) <= B(29) Then L = L + 1
If B(18) <= B(30) Then L = L + 1
If B(18) <= B(31) Then L = L + 1
If B(18) <= B(32) Then L = L + 1
If B(18) <= B(33) Then L = L + 1
If B(18) <= B(34) Then L = L + 1
If B(18) <= B(35) Then L = L + 1
If B(18) <= B(36) Then L = L + 1
If B(18) <= B(37) Then L = L + 1
If B(18) <= B(38) Then L = L + 1
If B(18) <= B(39) Then L = L + 1
If B(18) <= B(40) Then L = L + 1
If B(18) <= B(41) Then L = L + 1
If B(18) <= B(42) Then L = L + 1
If B(18) <= B(43) Then L = L + 1
If B(18) <= B(44) Then L = L + 1
If B(18) <= B(45) Then L = L + 1
If B(18) <= B(46) Then L = L + 1
If B(18) <= B(47) Then L = L + 1
If B(18) <= B(48) Then L = L + 1
If B(18) <= B(49) Then L = L + 1
 orderdisp
M = 19: L = 1
If B(19) < B(1) Then L = L + 1
If B(19) < B(2) Then L = L + 1
If B(19) < B(3) Then L = L + 1
If B(19) < B(4) Then L = L + 1
If B(19) < B(5) Then L = L + 1
If B(19) < B(6) Then L = L + 1
If B(19) < B(7) Then L = L + 1
If B(19) < B(8) Then L = L + 1
If B(19) < B(9) Then L = L + 1
If B(19) < B(10) Then L = L + 1
If B(19) < B(11) Then L = L + 1
If B(19) < B(12) Then L = L + 1
If B(19) < B(13) Then L = L + 1
If B(19) < B(14) Then L = L + 1
If B(19) < B(15) Then L = L + 1
If B(19) < B(16) Then L = L + 1
If B(19) < B(17) Then L = L + 1
If B(19) < B(18) Then L = L + 1
If B(19) <= B(20) Then L = L + 1
If B(19) <= B(21) Then L = L + 1
If B(19) <= B(22) Then L = L + 1
If B(19) <= B(23) Then L = L + 1
If B(19) <= B(24) Then L = L + 1
If B(19) <= B(25) Then L = L + 1
If B(19) <= B(26) Then L = L + 1
If B(19) <= B(27) Then L = L + 1
If B(19) <= B(28) Then L = L + 1
If B(19) <= B(29) Then L = L + 1
If B(19) <= B(30) Then L = L + 1
If B(19) <= B(31) Then L = L + 1
If B(19) <= B(32) Then L = L + 1
If B(19) <= B(33) Then L = L + 1
If B(19) <= B(34) Then L = L + 1
If B(19) <= B(35) Then L = L + 1
If B(19) <= B(36) Then L = L + 1
If B(19) <= B(37) Then L = L + 1
If B(19) <= B(38) Then L = L + 1
If B(19) <= B(39) Then L = L + 1
If B(19) <= B(40) Then L = L + 1
If B(19) <= B(41) Then L = L + 1
If B(19) <= B(42) Then L = L + 1
If B(19) <= B(43) Then L = L + 1
If B(19) <= B(44) Then L = L + 1
If B(19) <= B(45) Then L = L + 1
If B(19) <= B(46) Then L = L + 1
If B(19) <= B(47) Then L = L + 1
If B(19) <= B(48) Then L = L + 1
If B(19) <= B(49) Then L = L + 1
 orderdisp
M = 20: L = 1
If B(20) < B(1) Then L = L + 1
If B(20) < B(2) Then L = L + 1
If B(20) < B(3) Then L = L + 1
If B(20) < B(4) Then L = L + 1
If B(20) < B(5) Then L = L + 1
If B(20) < B(6) Then L = L + 1
If B(20) < B(7) Then L = L + 1
If B(20) < B(8) Then L = L + 1
If B(20) < B(9) Then L = L + 1
If B(20) < B(10) Then L = L + 1
If B(20) < B(11) Then L = L + 1
If B(20) < B(12) Then L = L + 1
If B(20) < B(13) Then L = L + 1
If B(20) < B(14) Then L = L + 1
If B(20) < B(15) Then L = L + 1
If B(20) < B(16) Then L = L + 1
If B(20) < B(17) Then L = L + 1
If B(20) < B(18) Then L = L + 1
If B(20) < B(19) Then L = L + 1
If B(20) <= B(21) Then L = L + 1
If B(20) <= B(22) Then L = L + 1
If B(20) <= B(23) Then L = L + 1
If B(20) <= B(24) Then L = L + 1
If B(20) <= B(25) Then L = L + 1
If B(20) <= B(26) Then L = L + 1
If B(20) <= B(27) Then L = L + 1
If B(20) <= B(28) Then L = L + 1
If B(20) <= B(29) Then L = L + 1
If B(20) <= B(30) Then L = L + 1
If B(20) <= B(31) Then L = L + 1
If B(20) <= B(32) Then L = L + 1
If B(20) <= B(33) Then L = L + 1
If B(20) <= B(34) Then L = L + 1
If B(20) <= B(35) Then L = L + 1
If B(20) <= B(36) Then L = L + 1
If B(20) <= B(37) Then L = L + 1
If B(20) <= B(38) Then L = L + 1
If B(20) <= B(39) Then L = L + 1
If B(20) <= B(40) Then L = L + 1
If B(20) <= B(41) Then L = L + 1
If B(20) <= B(42) Then L = L + 1
If B(20) <= B(43) Then L = L + 1
If B(20) <= B(44) Then L = L + 1
If B(20) <= B(45) Then L = L + 1
If B(20) <= B(46) Then L = L + 1
If B(20) <= B(47) Then L = L + 1
If B(20) <= B(48) Then L = L + 1
If B(20) <= B(49) Then L = L + 1
 orderdisp
Order2
 End Function

Public Function orderdisp()
If L = 1 Then P(1) = M
If L = 2 Then P(2) = M
If L = 3 Then P(3) = M
If L = 4 Then P(4) = M
If L = 5 Then P(5) = M
If L = 6 Then P(6) = M
If L = 7 Then P(7) = M
If L = 8 Then P(8) = M
If L = 9 Then P(9) = M
If L = 10 Then P(10) = M
If L = 11 Then P(11) = M
If L = 12 Then P(12) = M
If L = 13 Then P(13) = M
If L = 14 Then P(14) = M
If L = 15 Then P(15) = M
If L = 16 Then P(16) = M
If L = 17 Then P(17) = M
If L = 18 Then P(18) = M
If L = 19 Then P(19) = M
If L = 20 Then P(20) = M
If L = 21 Then P(21) = M
If L = 22 Then P(22) = M
If L = 23 Then P(23) = M
If L = 24 Then P(24) = M
If L = 25 Then P(25) = M
If L = 26 Then P(26) = M
If L = 27 Then P(27) = M
If L = 28 Then P(28) = M
If L = 29 Then P(29) = M
If L = 30 Then P(30) = M
If L = 31 Then P(31) = M
If L = 32 Then P(32) = M
If L = 33 Then P(33) = M
If L = 34 Then P(34) = M
If L = 35 Then P(35) = M
If L = 36 Then P(36) = M
If L = 37 Then P(37) = M
If L = 38 Then P(38) = M
If L = 39 Then P(39) = M
If L = 40 Then P(40) = M
If L = 41 Then P(41) = M
If L = 42 Then P(42) = M
If L = 43 Then P(43) = M
If L = 44 Then P(44) = M
If L = 45 Then P(45) = M
If L = 46 Then P(46) = M
If L = 47 Then P(47) = M
If L = 48 Then P(48) = M
If L = 49 Then P(49) = M
End Function

Public Sub displayNums()
If AV = 0 Then
txtH(1).Caption = Str(P(1))
txtH(2).Caption = Str(P(2))
txtH(3).Caption = Str(P(3))
txtH(4).Caption = Str(P(4))
txtH(5).Caption = Str(P(5))
txtH(6).Caption = Str(P(6))
txtC(1).Caption = Str(P(49))
txtC(2).Caption = Str(P(48))
txtC(3).Caption = Str(P(47))
txtC(4).Caption = Str(P(46))
txtC(5).Caption = Str(P(45))
txtC(6).Caption = Str(P(44))
Else
txtH(1).Caption = Str(RB(1))
txtH(2).Caption = Str(RB(2))
txtH(3).Caption = Str(RB(3))
txtH(4).Caption = Str(RB(4))
txtH(5).Caption = Str(RB(5))
txtH(6).Caption = Str(RB(6))
'LblRb1.Caption = Str(RB(1)) + " " + Str(RB(2)) + " " + Str(RB(3)) + " " + Str(RB(4)) + " " + Str(RB(5)) + " " + Str(RB(6))
End If
End Sub
Public Sub NumDisplay()
R = 15
C = 14
If CD = 1 Then
If P(49) = 1 Then txtBall(1).BackColor = QBColor(11)
If P(49) = 2 Then txtBall(2).BackColor = QBColor(11)
If P(49) = 3 Then txtBall(3).BackColor = QBColor(11)
If P(49) = 4 Then txtBall(4).BackColor = QBColor(11)
If P(49) = 5 Then txtBall(5).BackColor = QBColor(11)
If P(49) = 6 Then txtBall(6).BackColor = QBColor(11)
If P(49) = 7 Then txtBall(7).BackColor = QBColor(11)
If P(49) = 8 Then txtBall(8).BackColor = QBColor(11)
If P(49) = 9 Then txtBall(9).BackColor = QBColor(11)
If P(49) = 10 Then txtBall(10).BackColor = QBColor(11)
If P(49) = 11 Then txtBall(11).BackColor = QBColor(11)
If P(49) = 12 Then txtBall(12).BackColor = QBColor(11)
If P(49) = 13 Then txtBall(13).BackColor = QBColor(11)
If P(49) = 14 Then txtBall(14).BackColor = QBColor(11)
If P(49) = 15 Then txtBall(15).BackColor = QBColor(11)
If P(49) = 16 Then txtBall(16).BackColor = QBColor(11)
If P(49) = 17 Then txtBall(17).BackColor = QBColor(11)
If P(49) = 18 Then txtBall(18).BackColor = QBColor(11)
If P(49) = 19 Then txtBall(19).BackColor = QBColor(11)
If P(49) = 20 Then txtBall(20).BackColor = QBColor(11)
If P(49) = 21 Then txtBall(21).BackColor = QBColor(11)
If P(49) = 22 Then txtBall(22).BackColor = QBColor(11)
If P(49) = 23 Then txtBall(23).BackColor = QBColor(11)
If P(49) = 24 Then txtBall(24).BackColor = QBColor(11)
If P(49) = 25 Then txtBall(25).BackColor = QBColor(11)
If P(49) = 26 Then txtBall(26).BackColor = QBColor(11)
If P(49) = 27 Then txtBall(27).BackColor = QBColor(11)
If P(49) = 28 Then txtBall(28).BackColor = QBColor(11)
If P(49) = 29 Then txtBall(29).BackColor = QBColor(11)
If P(49) = 30 Then txtBall(30).BackColor = QBColor(11)
If P(49) = 31 Then txtBall(31).BackColor = QBColor(11)
If P(49) = 32 Then txtBall(32).BackColor = QBColor(11)
If P(49) = 33 Then txtBall(33).BackColor = QBColor(11)
If P(49) = 34 Then txtBall(34).BackColor = QBColor(11)
If P(49) = 35 Then txtBall(35).BackColor = QBColor(11)
If P(49) = 36 Then txtBall(36).BackColor = QBColor(11)
If P(49) = 37 Then txtBall(37).BackColor = QBColor(11)
If P(49) = 38 Then txtBall(38).BackColor = QBColor(11)
If P(49) = 39 Then txtBall(39).BackColor = QBColor(11)
If P(49) = 40 Then txtBall(40).BackColor = QBColor(11)
If P(49) = 41 Then txtBall(41).BackColor = QBColor(11)
If P(49) = 42 Then txtBall(42).BackColor = QBColor(11)
If P(49) = 43 Then txtBall(43).BackColor = QBColor(11)
If P(49) = 44 Then txtBall(44).BackColor = QBColor(11)
If P(49) = 45 Then txtBall(45).BackColor = QBColor(11)
If P(49) = 46 Then txtBall(46).BackColor = QBColor(11)
If P(49) = 47 Then txtBall(47).BackColor = QBColor(11)
If P(49) = 48 Then txtBall(48).BackColor = QBColor(11)
If P(49) = 49 Then txtBall(49).BackColor = QBColor(11)
If P(48) = 1 Then txtBall(1).BackColor = QBColor(11)
If P(48) = 2 Then txtBall(2).BackColor = QBColor(11)
If P(48) = 3 Then txtBall(3).BackColor = QBColor(11)
If P(48) = 4 Then txtBall(4).BackColor = QBColor(11)
If P(48) = 5 Then txtBall(5).BackColor = QBColor(11)
If P(48) = 6 Then txtBall(6).BackColor = QBColor(11)
If P(48) = 7 Then txtBall(7).BackColor = QBColor(11)
If P(48) = 8 Then txtBall(8).BackColor = QBColor(11)
If P(48) = 9 Then txtBall(9).BackColor = QBColor(11)
If P(48) = 10 Then txtBall(10).BackColor = QBColor(11)
If P(48) = 11 Then txtBall(11).BackColor = QBColor(11)
If P(48) = 12 Then txtBall(12).BackColor = QBColor(11)
If P(48) = 13 Then txtBall(13).BackColor = QBColor(11)
If P(48) = 14 Then txtBall(14).BackColor = QBColor(11)
If P(48) = 15 Then txtBall(15).BackColor = QBColor(11)
If P(48) = 16 Then txtBall(16).BackColor = QBColor(11)
If P(48) = 17 Then txtBall(17).BackColor = QBColor(11)
If P(48) = 18 Then txtBall(18).BackColor = QBColor(11)
If P(48) = 19 Then txtBall(19).BackColor = QBColor(11)
If P(48) = 20 Then txtBall(20).BackColor = QBColor(11)
If P(48) = 21 Then txtBall(21).BackColor = QBColor(11)
If P(48) = 22 Then txtBall(22).BackColor = QBColor(11)
If P(48) = 23 Then txtBall(23).BackColor = QBColor(11)
If P(48) = 24 Then txtBall(24).BackColor = QBColor(11)
If P(48) = 25 Then txtBall(25).BackColor = QBColor(11)
If P(48) = 26 Then txtBall(26).BackColor = QBColor(11)
If P(48) = 27 Then txtBall(27).BackColor = QBColor(11)
If P(48) = 28 Then txtBall(28).BackColor = QBColor(11)
If P(48) = 29 Then txtBall(29).BackColor = QBColor(11)
If P(48) = 30 Then txtBall(30).BackColor = QBColor(11)
If P(48) = 31 Then txtBall(31).BackColor = QBColor(11)
If P(48) = 32 Then txtBall(32).BackColor = QBColor(11)
If P(48) = 33 Then txtBall(33).BackColor = QBColor(11)
If P(48) = 34 Then txtBall(34).BackColor = QBColor(11)
If P(48) = 35 Then txtBall(35).BackColor = QBColor(11)
If P(48) = 36 Then txtBall(36).BackColor = QBColor(11)
If P(48) = 37 Then txtBall(37).BackColor = QBColor(11)
If P(48) = 38 Then txtBall(38).BackColor = QBColor(11)
If P(48) = 39 Then txtBall(39).BackColor = QBColor(11)
If P(48) = 40 Then txtBall(40).BackColor = QBColor(11)
If P(48) = 41 Then txtBall(41).BackColor = QBColor(11)
If P(48) = 42 Then txtBall(42).BackColor = QBColor(11)
If P(48) = 43 Then txtBall(43).BackColor = QBColor(11)
If P(48) = 44 Then txtBall(44).BackColor = QBColor(11)
If P(48) = 45 Then txtBall(45).BackColor = QBColor(11)
If P(48) = 46 Then txtBall(46).BackColor = QBColor(11)
If P(48) = 47 Then txtBall(47).BackColor = QBColor(11)
If P(48) = 48 Then txtBall(48).BackColor = QBColor(11)
If P(48) = 49 Then txtBall(49).BackColor = QBColor(11)
If P(47) = 1 Then txtBall(1).BackColor = QBColor(11)
If P(47) = 2 Then txtBall(2).BackColor = QBColor(11)
If P(47) = 3 Then txtBall(3).BackColor = QBColor(11)
If P(47) = 4 Then txtBall(4).BackColor = QBColor(11)
If P(47) = 5 Then txtBall(5).BackColor = QBColor(11)
If P(47) = 6 Then txtBall(6).BackColor = QBColor(11)
If P(47) = 7 Then txtBall(7).BackColor = QBColor(11)
If P(47) = 8 Then txtBall(8).BackColor = QBColor(11)
If P(47) = 9 Then txtBall(9).BackColor = QBColor(11)
If P(47) = 10 Then txtBall(10).BackColor = QBColor(11)
If P(47) = 11 Then txtBall(11).BackColor = QBColor(11)
If P(47) = 12 Then txtBall(12).BackColor = QBColor(11)
If P(47) = 13 Then txtBall(13).BackColor = QBColor(11)
If P(47) = 14 Then txtBall(14).BackColor = QBColor(11)
If P(47) = 15 Then txtBall(15).BackColor = QBColor(11)
If P(47) = 16 Then txtBall(16).BackColor = QBColor(11)
If P(47) = 17 Then txtBall(17).BackColor = QBColor(11)
If P(47) = 18 Then txtBall(18).BackColor = QBColor(11)
If P(47) = 19 Then txtBall(19).BackColor = QBColor(11)
If P(47) = 20 Then txtBall(20).BackColor = QBColor(11)
If P(47) = 21 Then txtBall(21).BackColor = QBColor(11)
If P(47) = 22 Then txtBall(22).BackColor = QBColor(11)
If P(47) = 23 Then txtBall(23).BackColor = QBColor(11)
If P(47) = 24 Then txtBall(24).BackColor = QBColor(11)
If P(47) = 25 Then txtBall(25).BackColor = QBColor(11)
If P(47) = 26 Then txtBall(26).BackColor = QBColor(11)
If P(47) = 27 Then txtBall(27).BackColor = QBColor(11)
If P(47) = 28 Then txtBall(28).BackColor = QBColor(11)
If P(47) = 29 Then txtBall(29).BackColor = QBColor(11)
If P(47) = 30 Then txtBall(30).BackColor = QBColor(11)
If P(47) = 31 Then txtBall(31).BackColor = QBColor(11)
If P(47) = 32 Then txtBall(32).BackColor = QBColor(11)
If P(47) = 33 Then txtBall(33).BackColor = QBColor(11)
If P(47) = 34 Then txtBall(34).BackColor = QBColor(11)
If P(47) = 35 Then txtBall(35).BackColor = QBColor(11)
If P(47) = 36 Then txtBall(36).BackColor = QBColor(11)
If P(47) = 37 Then txtBall(37).BackColor = QBColor(11)
If P(47) = 38 Then txtBall(38).BackColor = QBColor(11)
If P(47) = 39 Then txtBall(39).BackColor = QBColor(11)
If P(47) = 40 Then txtBall(40).BackColor = QBColor(11)
If P(47) = 41 Then txtBall(41).BackColor = QBColor(11)
If P(47) = 42 Then txtBall(42).BackColor = QBColor(11)
If P(47) = 43 Then txtBall(43).BackColor = QBColor(11)
If P(47) = 44 Then txtBall(44).BackColor = QBColor(11)
If P(47) = 45 Then txtBall(45).BackColor = QBColor(11)
If P(47) = 46 Then txtBall(46).BackColor = QBColor(11)
If P(47) = 47 Then txtBall(47).BackColor = QBColor(11)
If P(47) = 48 Then txtBall(48).BackColor = QBColor(11)
If P(47) = 49 Then txtBall(49).BackColor = QBColor(11)
If P(46) = 1 Then txtBall(1).BackColor = QBColor(11)
If P(46) = 2 Then txtBall(2).BackColor = QBColor(11)
If P(46) = 3 Then txtBall(3).BackColor = QBColor(11)
If P(46) = 4 Then txtBall(4).BackColor = QBColor(11)
If P(46) = 5 Then txtBall(5).BackColor = QBColor(11)
If P(46) = 6 Then txtBall(6).BackColor = QBColor(11)
If P(46) = 7 Then txtBall(7).BackColor = QBColor(11)
If P(46) = 8 Then txtBall(8).BackColor = QBColor(11)
If P(46) = 9 Then txtBall(9).BackColor = QBColor(11)
If P(46) = 10 Then txtBall(10).BackColor = QBColor(11)
If P(46) = 11 Then txtBall(11).BackColor = QBColor(11)
If P(46) = 12 Then txtBall(12).BackColor = QBColor(11)
If P(46) = 13 Then txtBall(13).BackColor = QBColor(11)
If P(46) = 14 Then txtBall(14).BackColor = QBColor(11)
If P(46) = 15 Then txtBall(15).BackColor = QBColor(11)
If P(46) = 16 Then txtBall(16).BackColor = QBColor(11)
If P(46) = 17 Then txtBall(17).BackColor = QBColor(11)
If P(46) = 18 Then txtBall(18).BackColor = QBColor(11)
If P(46) = 19 Then txtBall(19).BackColor = QBColor(11)
If P(46) = 20 Then txtBall(20).BackColor = QBColor(11)
If P(46) = 21 Then txtBall(21).BackColor = QBColor(11)
If P(46) = 22 Then txtBall(22).BackColor = QBColor(11)
If P(46) = 23 Then txtBall(23).BackColor = QBColor(11)
If P(46) = 24 Then txtBall(24).BackColor = QBColor(11)
If P(46) = 25 Then txtBall(25).BackColor = QBColor(11)
If P(46) = 26 Then txtBall(26).BackColor = QBColor(11)
If P(46) = 27 Then txtBall(27).BackColor = QBColor(11)
If P(46) = 28 Then txtBall(28).BackColor = QBColor(11)
If P(46) = 29 Then txtBall(29).BackColor = QBColor(11)
If P(46) = 30 Then txtBall(30).BackColor = QBColor(11)
If P(46) = 31 Then txtBall(31).BackColor = QBColor(11)
If P(46) = 32 Then txtBall(32).BackColor = QBColor(11)
If P(46) = 33 Then txtBall(33).BackColor = QBColor(11)
If P(46) = 34 Then txtBall(34).BackColor = QBColor(11)
If P(46) = 35 Then txtBall(35).BackColor = QBColor(11)
If P(46) = 36 Then txtBall(36).BackColor = QBColor(11)
If P(46) = 37 Then txtBall(37).BackColor = QBColor(11)
If P(46) = 38 Then txtBall(38).BackColor = QBColor(11)
If P(46) = 39 Then txtBall(39).BackColor = QBColor(11)
If P(46) = 40 Then txtBall(40).BackColor = QBColor(11)
If P(46) = 41 Then txtBall(41).BackColor = QBColor(11)
If P(46) = 42 Then txtBall(42).BackColor = QBColor(11)
If P(46) = 43 Then txtBall(43).BackColor = QBColor(11)
If P(46) = 44 Then txtBall(44).BackColor = QBColor(11)
If P(46) = 45 Then txtBall(45).BackColor = QBColor(11)
If P(46) = 46 Then txtBall(46).BackColor = QBColor(11)
If P(46) = 47 Then txtBall(47).BackColor = QBColor(11)
If P(46) = 48 Then txtBall(48).BackColor = QBColor(11)
If P(46) = 49 Then txtBall(49).BackColor = QBColor(11)
If P(45) = 1 Then txtBall(1).BackColor = QBColor(11)
If P(45) = 2 Then txtBall(2).BackColor = QBColor(11)
If P(45) = 3 Then txtBall(3).BackColor = QBColor(11)
If P(45) = 4 Then txtBall(4).BackColor = QBColor(11)
If P(45) = 5 Then txtBall(5).BackColor = QBColor(11)
If P(45) = 6 Then txtBall(6).BackColor = QBColor(11)
If P(45) = 7 Then txtBall(7).BackColor = QBColor(11)
If P(45) = 8 Then txtBall(8).BackColor = QBColor(11)
If P(45) = 9 Then txtBall(9).BackColor = QBColor(11)
If P(45) = 10 Then txtBall(10).BackColor = QBColor(11)
If P(45) = 11 Then txtBall(11).BackColor = QBColor(11)
If P(45) = 12 Then txtBall(12).BackColor = QBColor(11)
If P(45) = 13 Then txtBall(13).BackColor = QBColor(11)
If P(45) = 14 Then txtBall(14).BackColor = QBColor(11)
If P(45) = 15 Then txtBall(15).BackColor = QBColor(11)
If P(45) = 16 Then txtBall(16).BackColor = QBColor(11)
If P(45) = 17 Then txtBall(17).BackColor = QBColor(11)
If P(45) = 18 Then txtBall(18).BackColor = QBColor(11)
If P(45) = 19 Then txtBall(19).BackColor = QBColor(11)
If P(45) = 20 Then txtBall(20).BackColor = QBColor(11)
If P(45) = 21 Then txtBall(21).BackColor = QBColor(11)
If P(45) = 22 Then txtBall(22).BackColor = QBColor(11)
If P(45) = 23 Then txtBall(23).BackColor = QBColor(11)
If P(45) = 24 Then txtBall(24).BackColor = QBColor(11)
If P(45) = 25 Then txtBall(25).BackColor = QBColor(11)
If P(45) = 26 Then txtBall(26).BackColor = QBColor(11)
If P(45) = 27 Then txtBall(27).BackColor = QBColor(11)
If P(45) = 28 Then txtBall(28).BackColor = QBColor(11)
If P(45) = 29 Then txtBall(29).BackColor = QBColor(11)
If P(45) = 30 Then txtBall(30).BackColor = QBColor(11)
If P(45) = 31 Then txtBall(31).BackColor = QBColor(11)
If P(45) = 32 Then txtBall(32).BackColor = QBColor(11)
If P(45) = 33 Then txtBall(33).BackColor = QBColor(11)
If P(45) = 34 Then txtBall(34).BackColor = QBColor(11)
If P(45) = 35 Then txtBall(35).BackColor = QBColor(11)
If P(45) = 36 Then txtBall(36).BackColor = QBColor(11)
If P(45) = 37 Then txtBall(37).BackColor = QBColor(11)
If P(45) = 38 Then txtBall(38).BackColor = QBColor(11)
If P(45) = 39 Then txtBall(39).BackColor = QBColor(11)
If P(45) = 40 Then txtBall(40).BackColor = QBColor(11)
If P(45) = 41 Then txtBall(41).BackColor = QBColor(11)
If P(45) = 42 Then txtBall(42).BackColor = QBColor(11)
If P(45) = 43 Then txtBall(43).BackColor = QBColor(11)
If P(45) = 44 Then txtBall(44).BackColor = QBColor(11)
If P(45) = 45 Then txtBall(45).BackColor = QBColor(11)
If P(45) = 46 Then txtBall(46).BackColor = QBColor(11)
If P(45) = 47 Then txtBall(47).BackColor = QBColor(11)
If P(45) = 48 Then txtBall(48).BackColor = QBColor(11)
If P(45) = 49 Then txtBall(49).BackColor = QBColor(11)
If P(44) = 1 Then txtBall(1).BackColor = QBColor(11)
If P(44) = 2 Then txtBall(2).BackColor = QBColor(11)
If P(44) = 3 Then txtBall(3).BackColor = QBColor(11)
If P(44) = 4 Then txtBall(4).BackColor = QBColor(11)
If P(44) = 5 Then txtBall(5).BackColor = QBColor(11)
If P(44) = 6 Then txtBall(6).BackColor = QBColor(11)
If P(44) = 7 Then txtBall(7).BackColor = QBColor(11)
If P(44) = 8 Then txtBall(8).BackColor = QBColor(11)
If P(44) = 9 Then txtBall(9).BackColor = QBColor(11)
If P(44) = 10 Then txtBall(10).BackColor = QBColor(11)
If P(44) = 11 Then txtBall(11).BackColor = QBColor(11)
If P(44) = 12 Then txtBall(12).BackColor = QBColor(11)
If P(44) = 13 Then txtBall(13).BackColor = QBColor(11)
If P(44) = 14 Then txtBall(14).BackColor = QBColor(11)
If P(44) = 15 Then txtBall(15).BackColor = QBColor(11)
If P(44) = 16 Then txtBall(16).BackColor = QBColor(11)
If P(44) = 17 Then txtBall(17).BackColor = QBColor(11)
If P(44) = 18 Then txtBall(18).BackColor = QBColor(11)
If P(44) = 19 Then txtBall(19).BackColor = QBColor(11)
If P(44) = 20 Then txtBall(20).BackColor = QBColor(11)
If P(44) = 21 Then txtBall(21).BackColor = QBColor(11)
If P(44) = 22 Then txtBall(22).BackColor = QBColor(11)
If P(44) = 23 Then txtBall(23).BackColor = QBColor(11)
If P(44) = 24 Then txtBall(24).BackColor = QBColor(11)
If P(44) = 25 Then txtBall(25).BackColor = QBColor(11)
If P(44) = 26 Then txtBall(26).BackColor = QBColor(11)
If P(44) = 27 Then txtBall(27).BackColor = QBColor(11)
If P(44) = 28 Then txtBall(28).BackColor = QBColor(11)
If P(44) = 29 Then txtBall(29).BackColor = QBColor(11)
If P(44) = 30 Then txtBall(30).BackColor = QBColor(11)
If P(44) = 31 Then txtBall(31).BackColor = QBColor(11)
If P(44) = 32 Then txtBall(32).BackColor = QBColor(11)
If P(44) = 33 Then txtBall(33).BackColor = QBColor(11)
If P(44) = 34 Then txtBall(34).BackColor = QBColor(11)
If P(44) = 35 Then txtBall(35).BackColor = QBColor(11)
If P(44) = 36 Then txtBall(36).BackColor = QBColor(11)
If P(44) = 37 Then txtBall(37).BackColor = QBColor(11)
If P(44) = 38 Then txtBall(38).BackColor = QBColor(11)
If P(44) = 39 Then txtBall(39).BackColor = QBColor(11)
If P(44) = 40 Then txtBall(40).BackColor = QBColor(11)
If P(44) = 41 Then txtBall(41).BackColor = QBColor(11)
If P(44) = 42 Then txtBall(42).BackColor = QBColor(11)
If P(44) = 43 Then txtBall(43).BackColor = QBColor(11)
If P(44) = 44 Then txtBall(44).BackColor = QBColor(11)
If P(44) = 45 Then txtBall(45).BackColor = QBColor(11)
If P(44) = 46 Then txtBall(46).BackColor = QBColor(11)
If P(44) = 47 Then txtBall(47).BackColor = QBColor(11)
If P(44) = 48 Then txtBall(48).BackColor = QBColor(11)
If P(44) = 49 Then txtBall(49).BackColor = QBColor(11)
End If
If HD = 1 Then
If P(1) = 1 Then txtBall(1).BackColor = QBColor(12)
If P(1) = 2 Then txtBall(2).BackColor = QBColor(12)
If P(1) = 3 Then txtBall(3).BackColor = QBColor(12)
If P(1) = 4 Then txtBall(4).BackColor = QBColor(12)
If P(1) = 5 Then txtBall(5).BackColor = QBColor(12)
If P(1) = 6 Then txtBall(6).BackColor = QBColor(12)
If P(1) = 7 Then txtBall(7).BackColor = QBColor(12)
If P(1) = 8 Then txtBall(8).BackColor = QBColor(12)
If P(1) = 9 Then txtBall(9).BackColor = QBColor(12)
If P(1) = 10 Then txtBall(10).BackColor = QBColor(12)
If P(1) = 11 Then txtBall(11).BackColor = QBColor(12)
If P(1) = 12 Then txtBall(12).BackColor = QBColor(12)
If P(1) = 13 Then txtBall(13).BackColor = QBColor(12)
If P(1) = 14 Then txtBall(14).BackColor = QBColor(12)
If P(1) = 15 Then txtBall(15).BackColor = QBColor(12)
If P(1) = 16 Then txtBall(16).BackColor = QBColor(12)
If P(1) = 17 Then txtBall(17).BackColor = QBColor(12)
If P(1) = 18 Then txtBall(18).BackColor = QBColor(12)
If P(1) = 19 Then txtBall(19).BackColor = QBColor(12)
If P(1) = 20 Then txtBall(20).BackColor = QBColor(12)
If P(1) = 21 Then txtBall(21).BackColor = QBColor(12)
If P(1) = 22 Then txtBall(22).BackColor = QBColor(12)
If P(1) = 23 Then txtBall(23).BackColor = QBColor(12)
If P(1) = 24 Then txtBall(24).BackColor = QBColor(12)
If P(1) = 25 Then txtBall(25).BackColor = QBColor(12)
If P(1) = 26 Then txtBall(26).BackColor = QBColor(12)
If P(1) = 27 Then txtBall(27).BackColor = QBColor(12)
If P(1) = 28 Then txtBall(28).BackColor = QBColor(12)
If P(1) = 29 Then txtBall(29).BackColor = QBColor(12)
If P(1) = 30 Then txtBall(30).BackColor = QBColor(12)
If P(1) = 31 Then txtBall(31).BackColor = QBColor(12)
If P(1) = 32 Then txtBall(32).BackColor = QBColor(12)
If P(1) = 33 Then txtBall(33).BackColor = QBColor(12)
If P(1) = 34 Then txtBall(34).BackColor = QBColor(12)
If P(1) = 35 Then txtBall(35).BackColor = QBColor(12)
If P(1) = 36 Then txtBall(36).BackColor = QBColor(12)
If P(1) = 37 Then txtBall(37).BackColor = QBColor(12)
If P(1) = 38 Then txtBall(38).BackColor = QBColor(12)
If P(1) = 39 Then txtBall(39).BackColor = QBColor(12)
If P(1) = 40 Then txtBall(40).BackColor = QBColor(12)
If P(1) = 41 Then txtBall(41).BackColor = QBColor(12)
If P(1) = 42 Then txtBall(42).BackColor = QBColor(12)
If P(1) = 43 Then txtBall(43).BackColor = QBColor(12)
If P(1) = 44 Then txtBall(44).BackColor = QBColor(12)
If P(1) = 45 Then txtBall(45).BackColor = QBColor(12)
If P(1) = 46 Then txtBall(46).BackColor = QBColor(12)
If P(1) = 47 Then txtBall(47).BackColor = QBColor(12)
If P(1) = 48 Then txtBall(48).BackColor = QBColor(12)
If P(1) = 49 Then txtBall(49).BackColor = QBColor(12)
If P(2) = 1 Then txtBall(1).BackColor = QBColor(12)
If P(2) = 2 Then txtBall(2).BackColor = QBColor(12)
If P(2) = 3 Then txtBall(3).BackColor = QBColor(12)
If P(2) = 4 Then txtBall(4).BackColor = QBColor(12)
If P(2) = 5 Then txtBall(5).BackColor = QBColor(12)
If P(2) = 6 Then txtBall(6).BackColor = QBColor(12)
If P(2) = 7 Then txtBall(7).BackColor = QBColor(12)
If P(2) = 8 Then txtBall(8).BackColor = QBColor(12)
If P(2) = 9 Then txtBall(9).BackColor = QBColor(12)
If P(2) = 10 Then txtBall(10).BackColor = QBColor(12)
If P(2) = 11 Then txtBall(11).BackColor = QBColor(12)
If P(2) = 12 Then txtBall(12).BackColor = QBColor(12)
If P(2) = 13 Then txtBall(13).BackColor = QBColor(12)
If P(2) = 14 Then txtBall(14).BackColor = QBColor(12)
If P(2) = 15 Then txtBall(15).BackColor = QBColor(12)
If P(2) = 16 Then txtBall(16).BackColor = QBColor(12)
If P(2) = 17 Then txtBall(17).BackColor = QBColor(12)
If P(2) = 18 Then txtBall(18).BackColor = QBColor(12)
If P(2) = 19 Then txtBall(19).BackColor = QBColor(12)
If P(2) = 20 Then txtBall(20).BackColor = QBColor(12)
If P(2) = 21 Then txtBall(21).BackColor = QBColor(12)
If P(2) = 22 Then txtBall(22).BackColor = QBColor(12)
If P(2) = 23 Then txtBall(23).BackColor = QBColor(12)
If P(2) = 24 Then txtBall(24).BackColor = QBColor(12)
If P(2) = 25 Then txtBall(25).BackColor = QBColor(12)
If P(2) = 26 Then txtBall(26).BackColor = QBColor(12)
If P(2) = 27 Then txtBall(27).BackColor = QBColor(12)
If P(2) = 28 Then txtBall(28).BackColor = QBColor(12)
If P(2) = 29 Then txtBall(29).BackColor = QBColor(12)
If P(2) = 30 Then txtBall(30).BackColor = QBColor(12)
If P(2) = 31 Then txtBall(31).BackColor = QBColor(12)
If P(2) = 32 Then txtBall(32).BackColor = QBColor(12)
If P(2) = 33 Then txtBall(33).BackColor = QBColor(12)
If P(2) = 34 Then txtBall(34).BackColor = QBColor(12)
If P(2) = 35 Then txtBall(35).BackColor = QBColor(12)
If P(2) = 36 Then txtBall(36).BackColor = QBColor(12)
If P(2) = 37 Then txtBall(37).BackColor = QBColor(12)
If P(2) = 38 Then txtBall(38).BackColor = QBColor(12)
If P(2) = 39 Then txtBall(39).BackColor = QBColor(12)
If P(2) = 40 Then txtBall(40).BackColor = QBColor(12)
If P(2) = 41 Then txtBall(41).BackColor = QBColor(12)
If P(2) = 42 Then txtBall(42).BackColor = QBColor(12)
If P(2) = 43 Then txtBall(43).BackColor = QBColor(12)
If P(2) = 44 Then txtBall(44).BackColor = QBColor(12)
If P(2) = 45 Then txtBall(45).BackColor = QBColor(12)
If P(2) = 46 Then txtBall(46).BackColor = QBColor(12)
If P(2) = 47 Then txtBall(47).BackColor = QBColor(12)
If P(2) = 48 Then txtBall(48).BackColor = QBColor(12)
If P(2) = 49 Then txtBall(49).BackColor = QBColor(12)
If P(3) = 1 Then txtBall(1).BackColor = QBColor(12)
If P(3) = 2 Then txtBall(2).BackColor = QBColor(12)
If P(3) = 3 Then txtBall(3).BackColor = QBColor(12)
If P(3) = 4 Then txtBall(4).BackColor = QBColor(12)
If P(3) = 5 Then txtBall(5).BackColor = QBColor(12)
If P(3) = 6 Then txtBall(6).BackColor = QBColor(12)
If P(3) = 7 Then txtBall(7).BackColor = QBColor(12)
If P(3) = 8 Then txtBall(8).BackColor = QBColor(12)
If P(3) = 9 Then txtBall(9).BackColor = QBColor(12)
If P(3) = 10 Then txtBall(10).BackColor = QBColor(12)
If P(3) = 11 Then txtBall(11).BackColor = QBColor(12)
If P(3) = 12 Then txtBall(12).BackColor = QBColor(12)
If P(3) = 13 Then txtBall(13).BackColor = QBColor(12)
If P(3) = 14 Then txtBall(14).BackColor = QBColor(12)
If P(3) = 15 Then txtBall(15).BackColor = QBColor(12)
If P(3) = 16 Then txtBall(16).BackColor = QBColor(12)
If P(3) = 17 Then txtBall(17).BackColor = QBColor(12)
If P(3) = 18 Then txtBall(18).BackColor = QBColor(12)
If P(3) = 19 Then txtBall(19).BackColor = QBColor(12)
If P(3) = 20 Then txtBall(20).BackColor = QBColor(12)
If P(3) = 21 Then txtBall(21).BackColor = QBColor(12)
If P(3) = 22 Then txtBall(22).BackColor = QBColor(12)
If P(3) = 23 Then txtBall(23).BackColor = QBColor(12)
If P(3) = 24 Then txtBall(24).BackColor = QBColor(12)
If P(3) = 25 Then txtBall(25).BackColor = QBColor(12)
If P(3) = 26 Then txtBall(26).BackColor = QBColor(12)
If P(3) = 27 Then txtBall(27).BackColor = QBColor(12)
If P(3) = 28 Then txtBall(28).BackColor = QBColor(12)
If P(3) = 29 Then txtBall(29).BackColor = QBColor(12)
If P(3) = 30 Then txtBall(30).BackColor = QBColor(12)
If P(3) = 31 Then txtBall(31).BackColor = QBColor(12)
If P(3) = 32 Then txtBall(32).BackColor = QBColor(12)
If P(3) = 33 Then txtBall(33).BackColor = QBColor(12)
If P(3) = 34 Then txtBall(34).BackColor = QBColor(12)
If P(3) = 35 Then txtBall(35).BackColor = QBColor(12)
If P(3) = 36 Then txtBall(36).BackColor = QBColor(12)
If P(3) = 37 Then txtBall(37).BackColor = QBColor(12)
If P(3) = 38 Then txtBall(38).BackColor = QBColor(12)
If P(3) = 39 Then txtBall(39).BackColor = QBColor(12)
If P(3) = 40 Then txtBall(40).BackColor = QBColor(12)
If P(3) = 41 Then txtBall(41).BackColor = QBColor(12)
If P(3) = 42 Then txtBall(42).BackColor = QBColor(12)
If P(3) = 43 Then txtBall(43).BackColor = QBColor(12)
If P(3) = 44 Then txtBall(44).BackColor = QBColor(12)
If P(3) = 45 Then txtBall(45).BackColor = QBColor(12)
If P(3) = 46 Then txtBall(46).BackColor = QBColor(12)
If P(3) = 47 Then txtBall(47).BackColor = QBColor(12)
If P(3) = 48 Then txtBall(48).BackColor = QBColor(12)
If P(3) = 49 Then txtBall(49).BackColor = QBColor(12)
If P(4) = 1 Then txtBall(1).BackColor = QBColor(12)
If P(4) = 2 Then txtBall(2).BackColor = QBColor(12)
If P(4) = 3 Then txtBall(3).BackColor = QBColor(12)
If P(4) = 4 Then txtBall(4).BackColor = QBColor(12)
If P(4) = 5 Then txtBall(5).BackColor = QBColor(12)
If P(4) = 6 Then txtBall(6).BackColor = QBColor(12)
If P(4) = 7 Then txtBall(7).BackColor = QBColor(12)
If P(4) = 8 Then txtBall(8).BackColor = QBColor(12)
If P(4) = 9 Then txtBall(9).BackColor = QBColor(12)
If P(4) = 10 Then txtBall(10).BackColor = QBColor(12)
If P(4) = 11 Then txtBall(11).BackColor = QBColor(12)
If P(4) = 12 Then txtBall(12).BackColor = QBColor(12)
If P(4) = 13 Then txtBall(13).BackColor = QBColor(12)
If P(4) = 14 Then txtBall(14).BackColor = QBColor(12)
If P(4) = 15 Then txtBall(15).BackColor = QBColor(12)
If P(4) = 16 Then txtBall(16).BackColor = QBColor(12)
If P(4) = 17 Then txtBall(17).BackColor = QBColor(12)
If P(4) = 18 Then txtBall(18).BackColor = QBColor(12)
If P(4) = 19 Then txtBall(19).BackColor = QBColor(12)
If P(4) = 20 Then txtBall(20).BackColor = QBColor(12)
If P(4) = 21 Then txtBall(21).BackColor = QBColor(12)
If P(4) = 22 Then txtBall(22).BackColor = QBColor(12)
If P(4) = 23 Then txtBall(23).BackColor = QBColor(12)
If P(4) = 24 Then txtBall(24).BackColor = QBColor(12)
If P(4) = 25 Then txtBall(25).BackColor = QBColor(12)
If P(4) = 26 Then txtBall(26).BackColor = QBColor(12)
If P(4) = 27 Then txtBall(27).BackColor = QBColor(12)
If P(4) = 28 Then txtBall(28).BackColor = QBColor(12)
If P(4) = 29 Then txtBall(29).BackColor = QBColor(12)
If P(4) = 30 Then txtBall(30).BackColor = QBColor(12)
If P(4) = 31 Then txtBall(31).BackColor = QBColor(12)
If P(4) = 32 Then txtBall(32).BackColor = QBColor(12)
If P(4) = 33 Then txtBall(33).BackColor = QBColor(12)
If P(4) = 34 Then txtBall(34).BackColor = QBColor(12)
If P(4) = 35 Then txtBall(35).BackColor = QBColor(12)
If P(4) = 36 Then txtBall(36).BackColor = QBColor(12)
If P(4) = 37 Then txtBall(37).BackColor = QBColor(12)
If P(4) = 38 Then txtBall(38).BackColor = QBColor(12)
If P(4) = 39 Then txtBall(39).BackColor = QBColor(12)
If P(4) = 40 Then txtBall(40).BackColor = QBColor(12)
If P(4) = 41 Then txtBall(41).BackColor = QBColor(12)
If P(4) = 42 Then txtBall(42).BackColor = QBColor(12)
If P(4) = 43 Then txtBall(43).BackColor = QBColor(12)
If P(4) = 44 Then txtBall(44).BackColor = QBColor(12)
If P(4) = 45 Then txtBall(45).BackColor = QBColor(12)
If P(4) = 46 Then txtBall(46).BackColor = QBColor(12)
If P(4) = 47 Then txtBall(47).BackColor = QBColor(12)
If P(4) = 48 Then txtBall(48).BackColor = QBColor(12)
If P(4) = 49 Then txtBall(49).BackColor = QBColor(12)
If P(5) = 1 Then txtBall(1).BackColor = QBColor(12)
If P(5) = 2 Then txtBall(2).BackColor = QBColor(12)
If P(5) = 3 Then txtBall(3).BackColor = QBColor(12)
If P(5) = 4 Then txtBall(4).BackColor = QBColor(12)
If P(5) = 5 Then txtBall(5).BackColor = QBColor(12)
If P(5) = 6 Then txtBall(6).BackColor = QBColor(12)
If P(5) = 7 Then txtBall(7).BackColor = QBColor(12)
If P(5) = 8 Then txtBall(8).BackColor = QBColor(12)
If P(5) = 9 Then txtBall(9).BackColor = QBColor(12)
If P(5) = 10 Then txtBall(10).BackColor = QBColor(12)
If P(5) = 11 Then txtBall(11).BackColor = QBColor(12)
If P(5) = 12 Then txtBall(12).BackColor = QBColor(12)
If P(5) = 13 Then txtBall(13).BackColor = QBColor(12)
If P(5) = 14 Then txtBall(14).BackColor = QBColor(12)
If P(5) = 15 Then txtBall(15).BackColor = QBColor(12)
If P(5) = 16 Then txtBall(16).BackColor = QBColor(12)
If P(5) = 17 Then txtBall(17).BackColor = QBColor(12)
If P(5) = 18 Then txtBall(18).BackColor = QBColor(12)
If P(5) = 19 Then txtBall(19).BackColor = QBColor(12)
If P(5) = 20 Then txtBall(20).BackColor = QBColor(12)
If P(5) = 21 Then txtBall(21).BackColor = QBColor(12)
If P(5) = 22 Then txtBall(22).BackColor = QBColor(12)
If P(5) = 23 Then txtBall(23).BackColor = QBColor(12)
If P(5) = 24 Then txtBall(24).BackColor = QBColor(12)
If P(5) = 25 Then txtBall(25).BackColor = QBColor(12)
If P(5) = 26 Then txtBall(26).BackColor = QBColor(12)
If P(5) = 27 Then txtBall(27).BackColor = QBColor(12)
If P(5) = 28 Then txtBall(28).BackColor = QBColor(12)
If P(5) = 29 Then txtBall(29).BackColor = QBColor(12)
If P(5) = 30 Then txtBall(30).BackColor = QBColor(12)
If P(5) = 31 Then txtBall(31).BackColor = QBColor(12)
If P(5) = 32 Then txtBall(32).BackColor = QBColor(12)
If P(5) = 33 Then txtBall(33).BackColor = QBColor(12)
If P(5) = 34 Then txtBall(34).BackColor = QBColor(12)
If P(5) = 35 Then txtBall(35).BackColor = QBColor(12)
If P(5) = 36 Then txtBall(36).BackColor = QBColor(12)
If P(5) = 37 Then txtBall(37).BackColor = QBColor(12)
If P(5) = 38 Then txtBall(38).BackColor = QBColor(12)
If P(5) = 39 Then txtBall(39).BackColor = QBColor(12)
If P(5) = 40 Then txtBall(40).BackColor = QBColor(12)
If P(5) = 41 Then txtBall(41).BackColor = QBColor(12)
If P(5) = 42 Then txtBall(42).BackColor = QBColor(12)
If P(5) = 43 Then txtBall(43).BackColor = QBColor(12)
If P(5) = 44 Then txtBall(44).BackColor = QBColor(12)
If P(5) = 45 Then txtBall(45).BackColor = QBColor(12)
If P(5) = 46 Then txtBall(46).BackColor = QBColor(12)
If P(5) = 47 Then txtBall(47).BackColor = QBColor(12)
If P(5) = 48 Then txtBall(48).BackColor = QBColor(12)
If P(5) = 49 Then txtBall(49).BackColor = QBColor(12)
If P(6) = 1 Then txtBall(1).BackColor = QBColor(12)
If P(6) = 2 Then txtBall(2).BackColor = QBColor(12)
If P(6) = 3 Then txtBall(3).BackColor = QBColor(12)
If P(6) = 4 Then txtBall(4).BackColor = QBColor(12)
If P(6) = 5 Then txtBall(5).BackColor = QBColor(12)
If P(6) = 6 Then txtBall(6).BackColor = QBColor(12)
If P(6) = 7 Then txtBall(7).BackColor = QBColor(12)
If P(6) = 8 Then txtBall(8).BackColor = QBColor(12)
If P(6) = 9 Then txtBall(9).BackColor = QBColor(12)
If P(6) = 10 Then txtBall(10).BackColor = QBColor(12)
If P(6) = 11 Then txtBall(11).BackColor = QBColor(12)
If P(6) = 12 Then txtBall(12).BackColor = QBColor(12)
If P(6) = 13 Then txtBall(13).BackColor = QBColor(12)
If P(6) = 14 Then txtBall(14).BackColor = QBColor(12)
If P(6) = 15 Then txtBall(15).BackColor = QBColor(12)
If P(6) = 16 Then txtBall(16).BackColor = QBColor(12)
If P(6) = 17 Then txtBall(17).BackColor = QBColor(12)
If P(6) = 18 Then txtBall(18).BackColor = QBColor(12)
If P(6) = 19 Then txtBall(19).BackColor = QBColor(12)
If P(6) = 20 Then txtBall(20).BackColor = QBColor(12)
If P(6) = 21 Then txtBall(21).BackColor = QBColor(12)
If P(6) = 22 Then txtBall(22).BackColor = QBColor(12)
If P(6) = 23 Then txtBall(23).BackColor = QBColor(12)
If P(6) = 24 Then txtBall(24).BackColor = QBColor(12)
If P(6) = 25 Then txtBall(25).BackColor = QBColor(12)
If P(6) = 26 Then txtBall(26).BackColor = QBColor(12)
If P(6) = 27 Then txtBall(27).BackColor = QBColor(12)
If P(6) = 28 Then txtBall(28).BackColor = QBColor(12)
If P(6) = 29 Then txtBall(29).BackColor = QBColor(12)
If P(6) = 30 Then txtBall(30).BackColor = QBColor(12)
If P(6) = 31 Then txtBall(31).BackColor = QBColor(12)
If P(6) = 32 Then txtBall(32).BackColor = QBColor(12)
If P(6) = 33 Then txtBall(33).BackColor = QBColor(12)
If P(6) = 34 Then txtBall(34).BackColor = QBColor(12)
If P(6) = 35 Then txtBall(35).BackColor = QBColor(12)
If P(6) = 36 Then txtBall(36).BackColor = QBColor(12)
If P(6) = 37 Then txtBall(37).BackColor = QBColor(12)
If P(6) = 38 Then txtBall(38).BackColor = QBColor(12)
If P(6) = 39 Then txtBall(39).BackColor = QBColor(12)
If P(6) = 40 Then txtBall(40).BackColor = QBColor(12)
If P(6) = 41 Then txtBall(41).BackColor = QBColor(12)
If P(6) = 42 Then txtBall(42).BackColor = QBColor(12)
If P(6) = 43 Then txtBall(43).BackColor = QBColor(12)
If P(6) = 44 Then txtBall(44).BackColor = QBColor(12)
If P(6) = 45 Then txtBall(45).BackColor = QBColor(12)
If P(6) = 46 Then txtBall(46).BackColor = QBColor(12)
If P(6) = 47 Then txtBall(47).BackColor = QBColor(12)
If P(6) = 48 Then txtBall(48).BackColor = QBColor(12)
If P(6) = 49 Then txtBall(49).BackColor = QBColor(12)
End If
NumDisplay2
End Sub







Public Function CLC2()
txtBall(1).BackColor = QBColor(10)
txtBall(2).BackColor = QBColor(10)
txtBall(3).BackColor = QBColor(10)
txtBall(4).BackColor = QBColor(10)
txtBall(5).BackColor = QBColor(10)
txtBall(6).BackColor = QBColor(10)
txtBall(7).BackColor = QBColor(10)
txtBall(8).BackColor = QBColor(10)
txtBall(9).BackColor = QBColor(10)
txtBall(10).BackColor = QBColor(10)
txtBall(11).BackColor = QBColor(10)
txtBall(12).BackColor = QBColor(10)
txtBall(13).BackColor = QBColor(10)
txtBall(14).BackColor = QBColor(10)
txtBall(15).BackColor = QBColor(10)
txtBall(16).BackColor = QBColor(10)
txtBall(17).BackColor = QBColor(10)
txtBall(18).BackColor = QBColor(10)
txtBall(19).BackColor = QBColor(10)
txtBall(20).BackColor = QBColor(10)
txtBall(21).BackColor = QBColor(10)
txtBall(22).BackColor = QBColor(10)
txtBall(23).BackColor = QBColor(10)
txtBall(24).BackColor = QBColor(10)
txtBall(25).BackColor = QBColor(10)
txtBall(26).BackColor = QBColor(10)
txtBall(27).BackColor = QBColor(10)
txtBall(28).BackColor = QBColor(10)
txtBall(29).BackColor = QBColor(10)
txtBall(30).BackColor = QBColor(10)
txtBall(31).BackColor = QBColor(10)
txtBall(32).BackColor = QBColor(10)
txtBall(33).BackColor = QBColor(10)
txtBall(34).BackColor = QBColor(10)
txtBall(35).BackColor = QBColor(10)
txtBall(36).BackColor = QBColor(10)
txtBall(37).BackColor = QBColor(10)
txtBall(38).BackColor = QBColor(10)
txtBall(39).BackColor = QBColor(10)
txtBall(40).BackColor = QBColor(10)
txtBall(41).BackColor = QBColor(10)
txtBall(42).BackColor = QBColor(10)
txtBall(43).BackColor = QBColor(10)
txtBall(44).BackColor = QBColor(10)
txtBall(45).BackColor = QBColor(10)
txtBall(46).BackColor = QBColor(10)
txtBall(47).BackColor = QBColor(10)
txtBall(48).BackColor = QBColor(10)
txtBall(49).BackColor = QBColor(10)
End Function
Public Function HotBalls()
PB(1) = P(1)
PB(2) = P(2)
PB(3) = P(3)
PB(4) = P(4)
PB(5) = P(5)
PB(6) = P(6)
End Function

Public Function ColdBalls()
PB(1) = P(49)
PB(2) = P(48)
PB(3) = P(47)
PB(4) = P(46)
PB(5) = P(45)
PB(6) = P(44)
End Function


Public Function hotballdisp()
lblPb(1).Caption = Str(PB(1))
lblPb(2).Caption = Str(PB(2))
lblPb(3).Caption = Str(PB(3))
lblPb(4).Caption = Str(PB(4))
lblPb(5).Caption = Str(PB(5))
lblPb(6).Caption = Str(PB(6))
End Function

Public Function Coldballdisp()
lblPb(1).Caption = Str(PB(1))
lblPb(2).Caption = Str(PB(2))
lblPb(3).Caption = Str(PB(3))
lblPb(4).Caption = Str(PB(4))
lblPb(5).Caption = Str(PB(5))
lblPb(6).Caption = Str(PB(6))
End Function

Public Function FastMatch()
MA1 = 0
If PB(1) = LB(1) Then
MA1 = MA1 + 1
End If
If PB(2) = LB(1) Then
MA1 = MA1 + 1
End If
If PB(3) = LB(1) Then
MA1 = MA1 + 1
End If
If PB(4) = LB(1) Then
MA1 = MA1 + 1
End If
If PB(5) = LB(1) Then
MA1 = MA1 + 1
End If
If PB(6) = LB(1) Then
MA1 = MA1 + 1
End If
If PB(1) = LB(2) Then
MA1 = MA1 + 1
End If
If PB(2) = LB(2) Then
MA1 = MA1 + 1
End If
If PB(3) = LB(2) Then
MA1 = MA1 + 1
End If
If PB(4) = LB(2) Then
MA1 = MA1 + 1
End If
If PB(5) = LB(2) Then
MA1 = MA1 + 1
End If
If PB(6) = LB(2) Then
MA1 = MA1 + 1
End If
If PB(1) = LB(3) Then
MA1 = MA1 + 1
End If
If PB(2) = LB(3) Then
MA1 = MA1 + 1
End If
If PB(3) = LB(3) Then
MA1 = MA1 + 1
End If
If PB(4) = LB(3) Then
MA1 = MA1 + 1
End If
If PB(5) = LB(3) Then
MA1 = MA1 + 1
End If
If PB(6) = LB(3) Then
MA1 = MA1 + 1
End If
If PB(1) = LB(4) Then
MA1 = MA1 + 1
End If
If PB(2) = LB(4) Then
MA1 = MA1 + 1
End If
If PB(3) = LB(4) Then
MA1 = MA1 + 1
End If
If PB(4) = LB(4) Then
MA1 = MA1 + 1
End If
If PB(5) = LB(4) Then
MA1 = MA1 + 1
End If
If PB(6) = LB(4) Then
MA1 = MA1 + 1
End If
If PB(1) = LB(5) Then
MA1 = MA1 + 1
End If
If PB(2) = LB(5) Then
MA1 = MA1 + 1
End If
If PB(3) = LB(5) Then
MA1 = MA1 + 1
End If
If PB(4) = LB(5) Then
MA1 = MA1 + 1
End If
If PB(5) = LB(5) Then
MA1 = MA1 + 1
End If
If PB(6) = LB(5) Then
MA1 = MA1 + 1
End If
If PB(1) = LB(6) Then
MA1 = MA1 + 1
End If
If PB(2) = LB(6) Then
MA1 = MA1 + 1
End If
If PB(3) = LB(6) Then
MA1 = MA1 + 1
End If
If PB(4) = LB(6) Then
MA1 = MA1 + 1
End If
If PB(5) = LB(6) Then
MA1 = MA1 + 1
End If
If PB(6) = LB(6) Then
MA1 = MA1 + 1
End If
Select Case MA1
Case 1
TB(1) = TB(1) + 1
Case 2
TB(2) = TB(2) + 1
Case 3
TB(3) = TB(3) + 1
Case 4
TB(4) = TB(4) + 1
Case 5
TB(5) = TB(5) + 1
Bonus
Case 6
TB(6) = TB(6) + 1
End Select
If TB(6) > 300 Then
MsgBox "Too Many Jackpots!"
End
End If
PR(0) = TB(0) * PRIZE(0)
PR(1) = TB(1) * PRIZE(1)
PR(2) = TB(2) * PRIZE(2)
PR(3) = TB(3) * PRIZE(3)
PR(4) = TB(4) * PRIZE(4)
PR(5) = TB(5) * PRIZE(5)
PR(6) = TB(6) * PRIZE(6)
WKK = PR(1) + PR(2) + PR(3) + PR(4) + PR(5) + PR(0) + PR(6)
NUMENT = NUMENT + Stake
Stack = Stack + 1
CL = NUMENT - WKK
PLT = CL / NUMENT * 100
PWN = WKK / NUMENT * 100
If PWN > 100 Then PWN = 100
If PRIZE(1) >= 1 Then
ODD = TB(1) + TB(2) + TB(3) + TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(2) >= 1 And PRIZE(1) = 0 Then
ODD = TB(2) + TB(3) + TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(3) >= 1 And PRIZE(1) = 0 And PRIZE(2) = 0 Then
ODD = TB(3) + TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(4) >= 1 And PRIZE(1) = 0 And PRIZE(2) = 0 And PRIZE(3) = 0 Then
ODD = TB(4) + TB(5) + TB(0) + TB(6)
End If
If PRIZE(5) >= 1 And PRIZE(1) = 0 And PRIZE(2) = 0 And PRIZE(3) = 0 And PRIZE(4) = 0 Then
ODD = TB(5) + TB(0) + TB(6)
End If
If ODD = 0 Then ODD = 1
ODDS = Stack / ODD
AGTA = AGTA + AGT
AGTO = AGTA / Stack
RangeCalc
If Hsum < AGT Then
Hsum = AGT
End If
If Lsum > AGT Then
Lsum = AGT
End If
End Function
Public Function FASTLOTBALLS()
Randomize Timer
LB(1) = Int((49 * Rnd) + 1)
Do
Randomize Timer
LB(2) = Int((49 * Rnd) + 1)
Loop Until LB(2) <> LB(1)
Do
Randomize Timer
LB(3) = Int((49 * Rnd) + 1)
Loop Until LB(3) <> LB(2) And LB(3) <> LB(1)
Do
Randomize Timer
LB(4) = Int((49 * Rnd) + 1)
Loop Until LB(4) <> LB(3) And LB(4) <> LB(2) And LB(4) <> LB(1)
Do
Randomize Timer
LB(5) = Int((49 * Rnd) + 1)
Loop Until LB(5) <> LB(4) And LB(5) <> LB(3) And LB(5) <> LB(2) And LB(5) <> LB(1)
Do
Randomize Timer
LB(6) = Int((49 * Rnd) + 1)
Loop Until LB(6) <> LB(5) And LB(6) <> LB(4) And LB(6) <> LB(3) And LB(6) <> LB(2) And LB(6) <> LB(1)
Do
Randomize Timer
LB(0) = Int((49 * Rnd) + 1)
Loop Until LB(0) <> LB(6) And LB(0) <> LB(5) And LB(0) <> LB(4) And LB(0) <> LB(3) And LB(0) <> LB(2) And LB(0) <> LB(1)
SortLottery
If PC = 1 Then
NumCounter
End If
AGT = LB(1) + LB(2) + LB(3) + LB(4) + LB(5) + LB(6)
End Function

Public Sub FastRandPlayerBalls()
Randomize Timer
PB(1) = Int((49 * Rnd) + 1)
Do
Randomize Timer
PB(2) = Int((49 * Rnd) + 1)
Loop Until PB(2) <> PB(1)
Do
Randomize Timer
PB(3) = Int((49 * Rnd) + 1)
Loop Until PB(3) <> PB(2) And PB(3) <> PB(1)
Do
Randomize Timer
PB(4) = Int((49 * Rnd) + 1)
Loop Until PB(4) <> PB(3) And PB(4) <> PB(2) And PB(4) <> PB(1)
Do
Randomize Timer
PB(5) = Int((49 * Rnd) + 1)
Loop Until PB(5) <> PB(4) And PB(5) <> PB(3) And PB(5) <> PB(2) And PB(5) <> PB(1)
Do
Randomize Timer
PB(6) = Int((49 * Rnd) + 1)
Loop Until PB(6) <> PB(5) And PB(6) <> PB(4) And PB(6) <> PB(3) And PB(6) <> PB(2) And PB(6) <> PB(1)
PAGT = 0
PAGT = PB(1) + PB(2) + PB(3) + PB(4) + PB(5) + PB(6)
PrangeCalc
End Sub
Public Function RangeCalc()
RP(1) = 0
RP(2) = 0
RP(3) = 0
RP(4) = 0
RP(5) = 0
RP(6) = 0
RP(0) = 0
If LB(1) > LB(2) And LB(1) > LB(3) And LB(1) > LB(4) And LB(1) > LB(5) And LB(1) > LB(6) And LB(1) > LB(0) Then
RP(1) = LB(1)
End If
If LB(2) > LB(1) And LB(2) > LB(3) And LB(2) > LB(4) And LB(2) > LB(5) And LB(2) > LB(6) And LB(2) > LB(0) Then
RP(1) = LB(2)
End If
If LB(3) > LB(1) And LB(3) > LB(2) And LB(3) > LB(4) And LB(3) > LB(5) And LB(3) > LB(6) And LB(3) > LB(0) Then
RP(1) = LB(3)
End If
If LB(4) > LB(1) And LB(4) > LB(2) And LB(4) > LB(3) And LB(4) > LB(5) And LB(4) > LB(6) And LB(4) > LB(0) Then
RP(1) = LB(4)
End If
If LB(5) > LB(1) And LB(5) > LB(2) And LB(5) > LB(3) And LB(5) > LB(4) And LB(5) > LB(6) And LB(5) > LB(0) Then
RP(1) = LB(5)
End If
If LB(6) > LB(1) And LB(6) > LB(2) And LB(6) > LB(3) And LB(6) > LB(4) And LB(6) > LB(5) And LB(6) > LB(0) Then
RP(1) = LB(6)
End If
If LB(0) > LB(1) And LB(0) > LB(2) And LB(0) > LB(3) And LB(0) > LB(4) And LB(0) > LB(5) And LB(0) > LB(6) Then
RP(1) = LB(0)
End If
If LB(1) < LB(2) And LB(1) < LB(3) And LB(1) < LB(4) And LB(1) < LB(5) And LB(1) < LB(6) And LB(1) < LB(0) Then
RP(0) = LB(1)
End If
If LB(2) < LB(1) And LB(2) < LB(3) And LB(2) < LB(4) And LB(2) < LB(5) And LB(2) < LB(6) And LB(2) < LB(0) Then
RP(0) = LB(2)
End If
If LB(3) < LB(1) And LB(3) < LB(2) And LB(3) < LB(4) And LB(3) < LB(5) And LB(3) < LB(6) And LB(3) < LB(0) Then
RP(0) = LB(3)
End If
If LB(4) < LB(1) And LB(4) < LB(2) And LB(4) < LB(3) And LB(4) < LB(5) And LB(4) < LB(6) And LB(4) < LB(0) Then
RP(0) = LB(4)
End If
If LB(5) < LB(1) And LB(5) < LB(2) And LB(5) < LB(3) And LB(5) < LB(4) And LB(5) < LB(6) And LB(5) < LB(0) Then
RP(0) = LB(5)
End If
If LB(6) < LB(1) And LB(6) < LB(2) And LB(6) < LB(3) And LB(6) < LB(4) And LB(6) < LB(5) And LB(6) < LB(0) Then
RP(0) = LB(6)
End If
If LB(0) < LB(1) And LB(0) < LB(2) And LB(0) < LB(3) And LB(0) < LB(4) And LB(0) < LB(5) And LB(0) < LB(6) Then
RP(0) = LB(0)
End If
RANGE = RP(1) - RP(0)
If Hrange <= RANGE Then
Hrange = RANGE
End If
If LRange >= RANGE Then
LRange = RANGE
End If
AVRANGEA = AVRANGEA + RANGE
AVRANGE = AVRANGEA / Stack
End Function

Private Sub txtBlock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
If LsetAct = 1 Then
MsgBox "You are running ACTUAL Lottery Draws, there are " + Str(IPUT) + " of them"
GoTo 100
Else
TOOL19.Picture = ImageList1.ListImages(56).Picture
End If
FrmRuns.Show 1
txtBlock.Caption = Str(IPUT)
100
TOOL19.Picture = ImageList1.ListImages(55).Picture
End If
If Button = vbRightButton Then
HPP = 19
FrmQuickHelp.Show
End If

End Sub

Private Sub txtC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 44
FrmQuickHelp.Show
End If
End Sub


Private Sub txtH_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 43
FrmQuickHelp.Show
End If
End Sub


Private Sub txtNument_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 24
FrmQuickHelp.Show
End If
End Sub


Private Sub txtOdds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 31
FrmQuickHelp.Show
End If
End Sub

Public Function RunRandom()
If IPUT <= 0 Then IPUT = 1
BP = IPUT
ProgressBar2.Max = Str(IPUT)
199
ProgressBar2.Value = Str(COUNTER)
Screen.MousePointer = 11
If EHN = 1 Then
HotBalls
GoTo 202
End If
If ECN = 1 Then
ColdBalls
GoTo 202
End If
If SetFix = 1 Then
FixedPlayerBalls
GoTo 202
End If
If ECN = 0 And SetFix = 0 And EHN = 0 And AV = 0 Then
200
FastRandPlayerBalls
If Sumfil = 0 Then
GoTo 202
End If
End If
If Sumfil = 1 Then
If PAGT > USL Then
GoTo 200
End If
End If
If Sumfil = 1 Then
If PAGT < LSL Then
GoTo 200
End If
End If
202
If RangeFil = 0 Then
GoTo 203
End If
If RangeFil = 1 Then
PrangeCalc
If Prange > URL Then
GoTo 200
End If
End If
If RangeFil = 1 Then
If Prange < LRL Then
GoTo 200
End If
End If
203
If AV = 1 Then
PB(1) = RB(1)
PB(2) = RB(2)
PB(3) = RB(3)
PB(4) = RB(4)
PB(5) = RB(5)
PB(6) = RB(6)
End If
FASTLOTBALLS
Order
FastMatch
If PWN < 0 Then PWN = 0
ProgressBar1.Value = Str(PWN)
COUNTER = COUNTER + 1
If COUNTER > IPUT + 1 Then GoTo 500
If COUNTER < IPUT + 1 Then GoTo 199
500
If ECN = 0 And EHN = 0 Then
SortPlayer
End If
clc
Order
'SortLottery
EndSeq
MA1 = 0
lblran.Caption = Str(Prange)
lblSum.Caption = Str(PAGT)
For J = 0 To 6
Ball(J).Caption = Str(TB(J))
Next
txtwn.Caption = "£" + Str(WKK)
txtNument.Caption = "£" + Str(NUMENT)
If PLT <= 0 Then PLT = 0
txtPerl.Caption = Str(PLT) + "%"
txtPWn.Caption = Str(PWN) + "%"
If PWN >= 100 Then
txtPWn.Caption = "< " + Str(PWN) + "%"
End If
ProgressBar1.Value = Str(PWN)
txtBlock.Caption = Str(IPUT)
txtOdds.Caption = Str(ODDS)
lblTotalF.Caption = Str(AGT)
lblAvtotalF.Caption = Str(AGTO)
lblRangeF.Caption = Str(RANGE)
lblAvRangeF.Caption = Str(AVRANGE)
lblHighRange.Caption = Str(Hrange)
lblLowRange.Caption = Str(LRange)
lblHighSum.Caption = Str(Hsum)
lblLowSum.Caption = Str(Lsum)
lblNumDraws.Caption = Str(Stack)
For J = 0 To 6
lbllb(J).Caption = Str(LB(J))
Next
For J = 1 To 6
lblPb(J).Caption = Str(PB(J))
Next
COUNTER = 1
ProgressBar2.Value = 0
Screen.MousePointer = 0
End Function
Public Function RunFix()
BP = IPUT
clc
ECN = 0
EHN = 0
ProgressBar2.Max = Str(IPUT)
199
ProgressBar2.Value = Str(COUNTER&)
If SetFix = 1 Then
FixedPlayerBalls
End If
If SetFix = 0 Then
FastRandPlayerBalls
End If
FastMatch
ProgressBar1.Value = Str(PWN)
COUNTER& = COUNTER& + 1
If COUNTER& > IPUT + 1 Then GoTo 500
If COUNTER& < IPUT + 1 Then GoTo 199
500
'SortLottery
SortPlayer
Order
NumDisplay
Select Case MA1
Case 1
TB(1) = TB(1) + 1
lblw(1).ForeColor = QBColor(D)
Case 2
TB(2) = TB(2) + 1
lblw(2).ForeColor = QBColor(D)
Case 3
lblw(3).ForeColor = QBColor(D)
TB(3) = TB(3) + 1
Case 4
lblw(4).ForeColor = QBColor(D)
TB(4) = TB(4) + 1
Case 5
lblw(5).ForeColor = QBColor(D)
TB(5) = TB(5) + 1
Bonus
Case 6
lblw(6).ForeColor = QBColor(D)
TB(6) = TB(6) + 1
End Select
MA1 = 0
For J = 0 To 6
Ball(J).Caption = Str(TB(J))
lbllb(J).Caption = Str(LB(J))
Next
txtwn.Caption = "£" + Str(WKK)
txtNument.Caption = "£" + Str(NUMENT)
If PLT <= 0 Then PLT = 0
txtPerl.Caption = Str(PLT) + "%"
txtPWn.Caption = Str(PWN) + "%"
If PWN >= 100 Then
txtPWn.Caption = "< " + Str(PWN) + "%"
End If
ProgressBar1.Value = Str(PWN)
txtBlock.Caption = Str(IPUT)
txtOdds.Caption = Str(ODDS)
lblTotalF.Caption = Str(AGT)
lblAvtotalF.Caption = Str(AGTO)
lblRangeF.Caption = Str(RANGE)
lblAvRangeF.Caption = Str(AVRANGE)
lblHighRange.Caption = Str(Hrange)
lblLowRange.Caption = Str(LRange)
lblHighSum.Caption = Str(Hsum)
lblLowSum.Caption = Str(Lsum)
lblPb(1).Caption = Str(PB(1))
lblPb(2).Caption = Str(PB(2))
lblPb(3).Caption = Str(PB(3))
lblPb(4).Caption = Str(PB(4))
lblPb(5).Caption = Str(PB(5))
lblPb(6).Caption = Str(PB(6))
COUNTER = 1
ProgressBar2.Value = 0
End Function
Public Function RunActual()
IPUT = FD
BP = IPUT
ProgressBar2.Max = Str(IPUT)
199
Screen.MousePointer = 11
If COUNTER <= 0 Then COUNTER = 1
ProgressBar2.Value = Str(COUNTER)
StartKey = StartKey + 1
If StartKey < 2 And LsetAct = 1 Then
NUMENT = 0
End If
If EHN = 1 Then
HotBalls
End If
If ECN = 1 Then
ColdBalls
End If
If SetFix = 1 Then
FixedPlayerBalls
End If
If ECN = 0 And SetFix = 0 And EHN = 0 And AV = 0 Then
200
FastRandPlayerBalls
If Sumfil = 0 Then
GoTo 202
End If
End If
If Sumfil = 1 Then
If PAGT > USL Then
GoTo 200
End If
End If
If Sumfil = 1 Then
If PAGT < LSL Then
GoTo 200
End If
End If
202
If RangeFil = 0 Then
GoTo 203
End If
If RangeFil = 1 Then
PrangeCalc
If Prange > URL Then
GoTo 200
End If
End If
If RangeFil = 1 Then
If Prange < LRL Then
GoTo 200
End If
End If
203
If AV = 1 Then
PB(1) = RB(1)
PB(2) = RB(2)
PB(3) = RB(3)
PB(4) = RB(4)
PB(5) = RB(5)
PB(6) = RB(6)
End If
ActBalls
Order
FastMatch
ProgressBar1.Value = Str(PWN)
COUNTER = COUNTER + 1
If COUNTER >= IPUT Then GoTo 500
If COUNTER < IPUT Then GoTo 199
500
SortPlayer
clc
NUMENT = NUMENT + Stake
NumDisplay
displayNums
MA1 = 0
For J = 1 To 6
lblPb(J).Caption = Str(PB(J))
lbllb(J).Caption = Str(LB(J))
Next
txtwn.Caption = "£" + Str(WKK)
txtNument.Caption = "£" + Str(NUMENT)
If PLT <= 0 Then PLT = 0
txtPerl.Caption = Str(PLT) + "%"
txtPWn.Caption = Str(PWN) + "%"
If PWN >= 100 Then
txtPWn.Caption = "< " + Str(PWN) + "%"
End If
ProgressBar1.Value = Str(PWN)
txtBlock.Caption = Str(IPUT)
txtOdds.Caption = Str(ODDS)
lblTotalF.Caption = Str(AGT)
lblAvtotalF.Caption = Str(AGTO)
lblRangeF.Caption = Str(RANGE)
lblAvRangeF.Caption = Str(AVRANGE)
lblHighRange.Caption = Str(Hrange)
lblLowRange.Caption = Str(LRange)
lblHighSum.Caption = Str(Hsum)
lblLowSum.Caption = Str(Lsum)
lblNumDraws.Caption = Str(Stack)
For J = 0 To 6
Ball(J).Caption = Str(TB(J))
Next
lblran.Caption = Str(Prange)
lblSum.Caption = Str(PAGT)
COUNTER = 1
ProgressBar2.Value = 0
Screen.MousePointer = 0
End Function
Public Function StepRan()
If EHN = 1 Then
HotBalls
hotballdisp
End If
If ECN = 1 Then
ColdBalls
Coldballdisp
End If
If SetFix = 1 Then
EHN = 0
ECN = 0
FixedPlayerBalls
End If
If ECN = 0 And EHN = 0 And SetFix = 0 And AV = 0 Then
1009
playerballs
If Sumfil = 1 And PAGT > USL Then
GoTo 1009
End If
If Sumfil = 1 And PAGT < LSL Then
GoTo 1009
End If
PrangeCalc
If RangeFil = 1 And Prange > URL Then
GoTo 1009
End If
If RangeFil = 1 And Prange < LRL Then
GoTo 1009
End If
End If
If AV = 1 Then
For J = 1 To 6
PB(J) = RB(J)
Next
End If
SortPlayer
lblPb(1).Caption = Str(PB(1))
lblPb(2).Caption = Str(PB(2))
lblPb(3).Caption = Str(PB(3))
lblPb(4).Caption = Str(PB(4))
lblPb(5).Caption = Str(PB(5))
lblPb(6).Caption = Str(PB(6))
lblSum.Caption = Str(PAGT)
lblran.Caption = Str(Prange)
clc
LotBalls
Order
Matcher
NumDisplay
If HSW = 1 Then
For J = 1 To 49
frmPositions.Field(J).Caption = Str(P(J))
Next
End If
End Function
Public Function StepActual()
clc
If EHN = 1 Then
HotBalls
hotballdisp
End If
If ECN = 1 Then
ColdBalls
Coldballdisp
End If
If SetFix = 1 Then
FixedPlayerBalls
End If
If ECN = 0 And EHN = 0 And SetFix = 0 And AV = 0 Then
1009
playerballs
If Sumfil = 1 And PAGT > USL Then
GoTo 1009
End If
If Sumfil = 1 And PAGT < LSL Then
GoTo 1009
End If
PrangeCalc
If RangeFil = 1 And Prange > URL Then
GoTo 1009
End If
If RangeFil = 1 And Prange < LRL Then
GoTo 1009
End If
End If
If AV = 1 Then
For J = 1 To 6
PB(J) = RB(J)
Next
End If
lblPb(1).Caption = Str(PB(1))
lblPb(2).Caption = Str(PB(2))
lblPb(3).Caption = Str(PB(3))
lblPb(4).Caption = Str(PB(4))
lblPb(5).Caption = Str(PB(5))
lblPb(6).Caption = Str(PB(6))
ActBalls
Order
Matcher
NumDisplay
If FB = FD Then
Timer1.Enabled = False
Pshow.Picture = ImageList1.ListImages(53).Picture
Vspeed.Visible = False
Beep
MsgBox "End of Draws, The results are now displayed on screen"
End If
If HSW = 1 Then
For J = 1 To 49
frmPositions.Field(J).Caption = Str(P(J))
Next
End If
End Function
Public Function StepFixed()
'Screen.MousePointer = 11
clc
If SetFix = 1 Then
FixedPlayerBalls
End If
If SetFix = 0 Then
playerballs
End If
Matcher
AGT = LB(1) + LB(2) + LB(3) + LB(4) + LB(5) + LB(6)
AGTA = AGTA + AGT
AGTO = AGTA / NUMENT
RangeCalc
lblTotalF.Caption = Str(AGT)
lblAvtotalF.Caption = Str(AGTO)
lblRangeF.Caption = Str(RANGE)
lblAvRangeF.Caption = Str(AVRANGE)
lblHighRange.Caption = Str(Hrange)
lblLowRange.Caption = Str(LRange)
lblHighSum.Caption = Str(Hsum)
lblLowSum.Caption = Str(Lsum)
'Screen.MousePointer = 0
End Function
Public Function ActBalls()
FB = FB + 1
If FB >= FD + 1 Then
GoTo 30
Else
Get #FA, FB, FC
LB(1) = Str(FC.Ball1)
LB(2) = Str(FC.Ball2)
LB(3) = Str(FC.Ball3)
LB(4) = Str(FC.Ball4)
LB(5) = Str(FC.ball5)
LB(6) = Str(FC.ball6)
LB(0) = Str(FC.BallBon)
lbllb(1).Caption = Str(LB(1))
lbllb(2).Caption = Str(LB(2))
lbllb(3).Caption = Str(LB(3))
lbllb(4).Caption = Str(LB(4))
lbllb(5).Caption = Str(LB(5))
lbllb(6).Caption = Str(LB(6))
lbllb(0).Caption = Str(LB(0))
lblDrawDate.Caption = Trim(FC.Date)
AGT = LB(1) + LB(2) + LB(3) + LB(4) + LB(5) + LB(6)
Random6
End If
If FD <= 1 Then
If LB(1) = 0 And LB(2) = 0 And LB(3) = 0 And LB(4) = 0 Then
MsgBox "There are no Lottery Draw Results stored. Please see Ordering the Draw Results in the Help Contents Menu"
LsetAct = 0
LsetRan = 1
TOOL3.Picture = ImageList1.ListImages(33).Picture
TOOL2.Picture = ImageList1.ListImages(20).Picture
mnuActualNumbers.Checked = False
mnuRandomNumbers.Checked = True
lblLotteryStat.Caption = "(Random Numbers)"
IPUT = 200
GoTo 30
End If
End If
NumCounter
30
End Function
Public Function ConvertData()
If LS(1) = 1 Then LB(1) = 1
If LS(1) = 2 Then LB(1) = 2
If LS(1) = 3 Then LB(1) = 3
If LS(1) = 4 Then LB(1) = 4
If LS(1) = 5 Then LB(1) = 5
If LS(1) = 6 Then LB(1) = 6
If LS(1) = 7 Then LB(1) = 7
If LS(1) = 8 Then LB(1) = 8
If LS(1) = 9 Then LB(1) = 9
If LS(1) = 10 Then LB(1) = 10
If LS(1) = 11 Then LB(1) = 11
If LS(1) = 12 Then LB(1) = 12
If LS(1) = 13 Then LB(1) = 13
If LS(1) = 14 Then LB(1) = 14
If LS(1) = 15 Then LB(1) = 15
If LS(1) = 16 Then LB(1) = 16
If LS(1) = 17 Then LB(1) = 17
If LS(1) = 18 Then LB(1) = 18
If LS(1) = 19 Then LB(1) = 19
If LS(1) = 20 Then LB(1) = 20
If LS(1) = 21 Then LB(1) = 21
If LS(1) = 22 Then LB(1) = 22
If LS(1) = 23 Then LB(1) = 23
If LS(1) = 24 Then LB(1) = 24
If LS(1) = 25 Then LB(1) = 25
If LS(1) = 26 Then LB(1) = 26
If LS(1) = 27 Then LB(1) = 27
If LS(1) = 28 Then LB(1) = 28
If LS(1) = 29 Then LB(1) = 29
If LS(1) = 30 Then LB(1) = 30
If LS(1) = 31 Then LB(1) = 31
If LS(1) = 32 Then LB(1) = 32
If LS(1) = 33 Then LB(1) = 33
If LS(1) = 34 Then LB(1) = 34
If LS(1) = 35 Then LB(1) = 35
If LS(1) = 36 Then LB(1) = 36
If LS(1) = 37 Then LB(1) = 37
If LS(1) = 38 Then LB(1) = 38
If LS(1) = 39 Then LB(1) = 39
If LS(1) = 40 Then LB(1) = 40
If LS(1) = 41 Then LB(1) = 41
If LS(1) = 42 Then LB(1) = 42
If LS(1) = 43 Then LB(1) = 43
If LS(1) = 44 Then LB(1) = 44
If LS(1) = 45 Then LB(1) = 45
If LS(1) = 46 Then LB(1) = 46
If LS(1) = 47 Then LB(1) = 47
If LS(1) = 48 Then LB(1) = 48
If LS(1) = 49 Then LB(1) = 49
If LS(2) = 1 Then LB(2) = 1
If LS(2) = 2 Then LB(2) = 2
If LS(2) = 3 Then LB(2) = 3
If LS(2) = 4 Then LB(2) = 4
If LS(2) = 5 Then LB(2) = 5
If LS(2) = 6 Then LB(2) = 6
If LS(2) = 7 Then LB(2) = 7
If LS(2) = 8 Then LB(2) = 8
If LS(2) = 9 Then LB(2) = 9
If LS(2) = 10 Then LB(2) = 10
If LS(2) = 11 Then LB(2) = 11
If LS(2) = 12 Then LB(2) = 12
If LS(2) = 13 Then LB(2) = 13
If LS(2) = 14 Then LB(2) = 14
If LS(2) = 15 Then LB(2) = 15
If LS(2) = 16 Then LB(2) = 16
If LS(2) = 17 Then LB(2) = 17
If LS(2) = 18 Then LB(2) = 18
If LS(2) = 19 Then LB(2) = 19
If LS(2) = 20 Then LB(2) = 20
If LS(2) = 21 Then LB(2) = 21
If LS(2) = 22 Then LB(2) = 22
If LS(2) = 23 Then LB(2) = 23
If LS(2) = 24 Then LB(2) = 24
If LS(2) = 25 Then LB(2) = 25
If LS(2) = 26 Then LB(2) = 26
If LS(2) = 27 Then LB(2) = 27
If LS(2) = 28 Then LB(2) = 28
If LS(2) = 29 Then LB(2) = 29
If LS(2) = 30 Then LB(2) = 30
If LS(2) = 31 Then LB(2) = 31
If LS(2) = 32 Then LB(2) = 32
If LS(2) = 33 Then LB(2) = 33
If LS(2) = 34 Then LB(2) = 34
If LS(2) = 35 Then LB(2) = 35
If LS(2) = 36 Then LB(2) = 36
If LS(2) = 37 Then LB(2) = 37
If LS(2) = 38 Then LB(2) = 38
If LS(2) = 39 Then LB(2) = 39
If LS(2) = 40 Then LB(2) = 40
If LS(2) = 41 Then LB(2) = 41
If LS(2) = 42 Then LB(2) = 42
If LS(2) = 43 Then LB(2) = 43
If LS(2) = 44 Then LB(2) = 44
If LS(2) = 45 Then LB(2) = 45
If LS(2) = 46 Then LB(2) = 46
If LS(2) = 47 Then LB(2) = 47
If LS(2) = 48 Then LB(2) = 48
If LS(2) = 49 Then LB(2) = 49
If LS(3) = 1 Then LB(3) = 1
If LS(3) = 2 Then LB(3) = 2
If LS(3) = 3 Then LB(3) = 3
If LS(3) = 4 Then LB(3) = 4
If LS(3) = 5 Then LB(3) = 5
If LS(3) = 6 Then LB(3) = 6
If LS(3) = 7 Then LB(3) = 7
If LS(3) = 8 Then LB(3) = 8
If LS(3) = 9 Then LB(3) = 9
If LS(3) = 10 Then LB(3) = 10
If LS(3) = 11 Then LB(3) = 11
If LS(3) = 12 Then LB(3) = 12
If LS(3) = 13 Then LB(3) = 13
If LS(3) = 14 Then LB(3) = 14
If LS(3) = 15 Then LB(3) = 15
If LS(3) = 16 Then LB(3) = 16
If LS(3) = 17 Then LB(3) = 17
If LS(3) = 18 Then LB(3) = 18
If LS(3) = 19 Then LB(3) = 19
If LS(3) = 20 Then LB(3) = 20
If LS(3) = 21 Then LB(3) = 21
If LS(3) = 22 Then LB(3) = 22
If LS(3) = 23 Then LB(3) = 23
If LS(3) = 24 Then LB(3) = 24
If LS(3) = 25 Then LB(3) = 25
If LS(3) = 26 Then LB(3) = 26
If LS(3) = 27 Then LB(3) = 27
If LS(3) = 28 Then LB(3) = 28
If LS(3) = 29 Then LB(3) = 29
If LS(3) = 30 Then LB(3) = 30
If LS(3) = 31 Then LB(3) = 31
If LS(3) = 32 Then LB(3) = 32
If LS(3) = 33 Then LB(3) = 33
If LS(3) = 34 Then LB(3) = 34
If LS(3) = 35 Then LB(3) = 35
If LS(3) = 36 Then LB(3) = 36
If LS(3) = 37 Then LB(3) = 37
If LS(3) = 38 Then LB(3) = 38
If LS(3) = 39 Then LB(3) = 39
If LS(3) = 40 Then LB(3) = 40
If LS(3) = 41 Then LB(3) = 41
If LS(3) = 42 Then LB(3) = 42
If LS(3) = 43 Then LB(3) = 43
If LS(3) = 44 Then LB(3) = 44
If LS(3) = 45 Then LB(3) = 45
If LS(3) = 46 Then LB(3) = 46
If LS(3) = 47 Then LB(3) = 47
If LS(3) = 48 Then LB(3) = 48
If LS(3) = 49 Then LB(3) = 49
If LS(4) = 1 Then LB(4) = 1
If LS(4) = 2 Then LB(4) = 2
If LS(4) = 3 Then LB(4) = 3
If LS(4) = 4 Then LB(4) = 4
If LS(4) = 5 Then LB(4) = 5
If LS(4) = 6 Then LB(4) = 6
If LS(4) = 7 Then LB(4) = 7
If LS(4) = 8 Then LB(4) = 8
If LS(4) = 9 Then LB(4) = 9
If LS(4) = 10 Then LB(4) = 10
If LS(4) = 11 Then LB(4) = 11
If LS(4) = 12 Then LB(4) = 12
If LS(4) = 13 Then LB(4) = 13
If LS(4) = 14 Then LB(4) = 14
If LS(4) = 15 Then LB(4) = 15
If LS(4) = 16 Then LB(4) = 16
If LS(4) = 17 Then LB(4) = 17
If LS(4) = 18 Then LB(4) = 18
If LS(4) = 19 Then LB(4) = 19
If LS(4) = 20 Then LB(4) = 20
If LS(4) = 21 Then LB(4) = 21
If LS(4) = 22 Then LB(4) = 22
If LS(4) = 23 Then LB(4) = 23
If LS(4) = 24 Then LB(4) = 24
If LS(4) = 25 Then LB(4) = 25
If LS(4) = 26 Then LB(4) = 26
If LS(4) = 27 Then LB(4) = 27
If LS(4) = 28 Then LB(4) = 28
If LS(4) = 29 Then LB(4) = 29
If LS(4) = 30 Then LB(4) = 30
If LS(4) = 31 Then LB(4) = 31
If LS(4) = 32 Then LB(4) = 32
If LS(4) = 33 Then LB(4) = 33
If LS(4) = 34 Then LB(4) = 34
If LS(4) = 35 Then LB(4) = 35
If LS(4) = 36 Then LB(4) = 36
If LS(4) = 37 Then LB(4) = 37
If LS(4) = 38 Then LB(4) = 38
If LS(4) = 39 Then LB(4) = 39
If LS(4) = 40 Then LB(4) = 40
If LS(4) = 41 Then LB(4) = 41
If LS(4) = 42 Then LB(4) = 42
If LS(4) = 43 Then LB(4) = 43
If LS(4) = 44 Then LB(4) = 44
If LS(4) = 45 Then LB(4) = 45
If LS(4) = 46 Then LB(4) = 46
If LS(4) = 47 Then LB(4) = 47
If LS(4) = 48 Then LB(4) = 48
If LS(4) = 49 Then LB(4) = 49
If LS(5) = 1 Then LB(5) = 1
If LS(5) = 2 Then LB(5) = 2
If LS(5) = 3 Then LB(5) = 3
If LS(5) = 4 Then LB(5) = 4
If LS(5) = 5 Then LB(5) = 5
If LS(5) = 6 Then LB(5) = 6
If LS(5) = 7 Then LB(5) = 7
If LS(5) = 8 Then LB(5) = 8
If LS(5) = 9 Then LB(5) = 9
If LS(5) = 10 Then LB(5) = 10
If LS(5) = 11 Then LB(5) = 11
If LS(5) = 12 Then LB(5) = 12
If LS(5) = 13 Then LB(5) = 13
If LS(5) = 14 Then LB(5) = 14
If LS(5) = 15 Then LB(5) = 15
If LS(5) = 16 Then LB(5) = 16
If LS(5) = 17 Then LB(5) = 17
If LS(5) = 18 Then LB(5) = 18
If LS(5) = 19 Then LB(5) = 19
If LS(5) = 20 Then LB(5) = 20
If LS(5) = 21 Then LB(5) = 21
If LS(5) = 22 Then LB(5) = 22
If LS(5) = 23 Then LB(5) = 23
If LS(5) = 24 Then LB(5) = 24
If LS(5) = 25 Then LB(5) = 25
If LS(5) = 26 Then LB(5) = 26
If LS(5) = 27 Then LB(5) = 27
If LS(5) = 28 Then LB(5) = 28
If LS(5) = 29 Then LB(5) = 29
If LS(5) = 30 Then LB(5) = 30
If LS(5) = 31 Then LB(5) = 31
If LS(5) = 32 Then LB(5) = 32
If LS(5) = 33 Then LB(5) = 33
If LS(5) = 34 Then LB(5) = 34
If LS(5) = 35 Then LB(5) = 35
If LS(5) = 36 Then LB(5) = 36
If LS(5) = 37 Then LB(5) = 37
If LS(5) = 38 Then LB(5) = 38
If LS(5) = 39 Then LB(5) = 39
If LS(5) = 40 Then LB(5) = 40
If LS(5) = 41 Then LB(5) = 41
If LS(5) = 42 Then LB(5) = 42
If LS(5) = 43 Then LB(5) = 43
If LS(5) = 44 Then LB(5) = 44
If LS(5) = 45 Then LB(5) = 45
If LS(5) = 46 Then LB(5) = 46
If LS(5) = 47 Then LB(5) = 47
If LS(5) = 48 Then LB(5) = 48
If LS(5) = 49 Then LB(5) = 49
If LS(6) = 1 Then LB(6) = 1
If LS(6) = 2 Then LB(6) = 2
If LS(6) = 3 Then LB(6) = 3
If LS(6) = 4 Then LB(6) = 4
If LS(6) = 5 Then LB(6) = 5
If LS(6) = 6 Then LB(6) = 6
If LS(6) = 7 Then LB(6) = 7
If LS(6) = 8 Then LB(6) = 8
If LS(6) = 9 Then LB(6) = 9
If LS(6) = 10 Then LB(6) = 10
If LS(6) = 11 Then LB(6) = 11
If LS(6) = 12 Then LB(6) = 12
If LS(6) = 13 Then LB(6) = 13
If LS(6) = 14 Then LB(6) = 14
If LS(6) = 15 Then LB(6) = 15
If LS(6) = 16 Then LB(6) = 16
If LS(6) = 17 Then LB(6) = 17
If LS(6) = 18 Then LB(6) = 18
If LS(6) = 19 Then LB(6) = 19
If LS(6) = 20 Then LB(6) = 20
If LS(6) = 21 Then LB(6) = 21
If LS(6) = 22 Then LB(6) = 22
If LS(6) = 23 Then LB(6) = 23
If LS(6) = 24 Then LB(6) = 24
If LS(6) = 25 Then LB(6) = 25
If LS(6) = 26 Then LB(6) = 26
If LS(6) = 27 Then LB(6) = 27
If LS(6) = 28 Then LB(6) = 28
If LS(6) = 29 Then LB(6) = 29
If LS(6) = 30 Then LB(6) = 30
If LS(6) = 31 Then LB(6) = 31
If LS(6) = 32 Then LB(6) = 32
If LS(6) = 33 Then LB(6) = 33
If LS(6) = 34 Then LB(6) = 34
If LS(6) = 35 Then LB(6) = 35
If LS(6) = 36 Then LB(6) = 36
If LS(6) = 37 Then LB(6) = 37
If LS(6) = 38 Then LB(6) = 38
If LS(6) = 39 Then LB(6) = 39
If LS(6) = 40 Then LB(6) = 40
If LS(6) = 41 Then LB(6) = 41
If LS(6) = 42 Then LB(6) = 42
If LS(6) = 43 Then LB(6) = 43
If LS(6) = 44 Then LB(6) = 44
If LS(6) = 45 Then LB(6) = 45
If LS(6) = 46 Then LB(6) = 46
If LS(6) = 47 Then LB(6) = 47
If LS(6) = 48 Then LB(6) = 48
If LS(6) = 49 Then LB(6) = 49
If LS(0) = 1 Then LB(0) = 1
If LS(0) = 2 Then LB(0) = 2
If LS(0) = 3 Then LB(0) = 3
If LS(0) = 4 Then LB(0) = 4
If LS(0) = 5 Then LB(0) = 5
If LS(0) = 6 Then LB(0) = 6
If LS(0) = 7 Then LB(0) = 7
If LS(0) = 8 Then LB(0) = 8
If LS(0) = 9 Then LB(0) = 9
If LS(0) = 10 Then LB(0) = 10
If LS(0) = 11 Then LB(0) = 11
If LS(0) = 12 Then LB(0) = 12
If LS(0) = 13 Then LB(0) = 13
If LS(0) = 14 Then LB(0) = 14
If LS(0) = 15 Then LB(0) = 15
If LS(0) = 16 Then LB(0) = 16
If LS(0) = 17 Then LB(0) = 17
If LS(0) = 18 Then LB(0) = 18
If LS(0) = 19 Then LB(0) = 19
If LS(0) = 20 Then LB(0) = 20
If LS(0) = 21 Then LB(0) = 21
If LS(0) = 22 Then LB(0) = 22
If LS(0) = 23 Then LB(0) = 23
If LS(0) = 24 Then LB(0) = 24
If LS(0) = 25 Then LB(0) = 25
If LS(0) = 26 Then LB(0) = 26
If LS(0) = 27 Then LB(0) = 27
If LS(0) = 28 Then LB(0) = 28
If LS(0) = 29 Then LB(0) = 29
If LS(0) = 30 Then LB(0) = 30
If LS(0) = 31 Then LB(0) = 31
If LS(0) = 32 Then LB(0) = 32
If LS(0) = 33 Then LB(0) = 33
If LS(0) = 34 Then LB(0) = 34
If LS(0) = 35 Then LB(0) = 35
If LS(0) = 36 Then LB(0) = 36
If LS(0) = 37 Then LB(0) = 37
If LS(0) = 38 Then LB(0) = 38
If LS(0) = 39 Then LB(0) = 39
If LS(0) = 40 Then LB(0) = 40
If LS(0) = 41 Then LB(0) = 41
If LS(0) = 42 Then LB(0) = 42
If LS(0) = 43 Then LB(0) = 43
If LS(0) = 44 Then LB(0) = 44
If LS(0) = 45 Then LB(0) = 45
If LS(0) = 46 Then LB(0) = 46
If LS(0) = 47 Then LB(0) = 47
If LS(0) = 48 Then LB(0) = 48
If LS(0) = 49 Then LB(0) = 49
End Function

Public Function ConvertAggTotal()
AggTotal = LB(1) + LB(2) + LB(3) + LB(4) + LB(5) + LB(6)
MsgBox "HIT!"
If AggTotal = 21 Then AGT = 21
If AggTotal = 22 Then AGT = 22
If AggTotal = 23 Then AGT = 23
If AggTotal = 24 Then AGT = 24
If AggTotal = 25 Then AGT = 25
If AggTotal = 26 Then AGT = 26
If AggTotal = 27 Then AGT = 27
If AggTotal = 28 Then AGT = 28
If AggTotal = 29 Then AGT = 29
If AggTotal = 30 Then AGT = 30
If AggTotal = 31 Then AGT = 31
If AggTotal = 32 Then AGT = 32
If AggTotal = 33 Then AGT = 33
If AggTotal = 34 Then AGT = 34
If AggTotal = 35 Then AGT = 35
If AggTotal = 36 Then AGT = 36
If AggTotal = 37 Then AGT = 37
If AggTotal = 38 Then AGT = 38
If AggTotal = 39 Then AGT = 39
If AggTotal = 40 Then AGT = 40
If AggTotal = 41 Then AGT = 41
If AggTotal = 42 Then AGT = 42
If AggTotal = 43 Then AGT = 43
If AggTotal = 44 Then AGT = 44
If AggTotal = 45 Then AGT = 45
If AggTotal = 46 Then AGT = 46
If AggTotal = 47 Then AGT = 47
If AggTotal = 48 Then AGT = 48
If AggTotal = 49 Then AGT = 49
If AggTotal = 50 Then AGT = 50
If AggTotal = 51 Then AGT = 51
If AggTotal = 52 Then AGT = 52
If AggTotal = 53 Then AGT = 53
If AggTotal = 54 Then AGT = 54
If AggTotal = 55 Then AGT = 55
If AggTotal = 56 Then AGT = 56
If AggTotal = 57 Then AGT = 57
If AggTotal = 58 Then AGT = 58
If AggTotal = 59 Then AGT = 59
If AggTotal = 60 Then AGT = 60
If AggTotal = 61 Then AGT = 61
If AggTotal = 62 Then AGT = 62
If AggTotal = 63 Then AGT = 63
If AggTotal = 64 Then AGT = 64
If AggTotal = 65 Then AGT = 65
If AggTotal = 66 Then AGT = 66
If AggTotal = 67 Then AGT = 67
If AggTotal = 68 Then AGT = 68
If AggTotal = 69 Then AGT = 69
If AggTotal = 70 Then AGT = 70
If AggTotal = 71 Then AGT = 71
If AggTotal = 72 Then AGT = 72
If AggTotal = 73 Then AGT = 73
If AggTotal = 74 Then AGT = 74
If AggTotal = 75 Then AGT = 75
If AggTotal = 76 Then AGT = 76
If AggTotal = 77 Then AGT = 77
If AggTotal = 78 Then AGT = 78
If AggTotal = 79 Then AGT = 79
If AggTotal = 80 Then AGT = 80
If AggTotal = 81 Then AGT = 81
If AggTotal = 82 Then AGT = 82
If AggTotal = 83 Then AGT = 83
If AggTotal = 84 Then AGT = 84
If AggTotal = 85 Then AGT = 85
If AggTotal = 86 Then AGT = 86
If AggTotal = 87 Then AGT = 87
If AggTotal = 88 Then AGT = 88
If AggTotal = 89 Then AGT = 89
If AggTotal = 90 Then AGT = 90
If AggTotal = 91 Then AGT = 91
If AggTotal = 92 Then AGT = 92
If AggTotal = 93 Then AGT = 93
If AggTotal = 94 Then AGT = 94
If AggTotal = 95 Then AGT = 95
If AggTotal = 96 Then AGT = 96
If AggTotal = 97 Then AGT = 97
If AggTotal = 98 Then AGT = 98
If AggTotal = 99 Then AGT = 99
If AggTotal = 100 Then AGT = 100
If AggTotal = 101 Then AGT = 101
If AggTotal = 102 Then AGT = 102
If AggTotal = 103 Then AGT = 103
If AggTotal = 104 Then AGT = 104
If AggTotal = 105 Then AGT = 105
If AggTotal = 106 Then AGT = 106
If AggTotal = 107 Then AGT = 107
If AggTotal = 108 Then AGT = 108
If AggTotal = 109 Then AGT = 109
If AggTotal = 110 Then AGT = 110
If AggTotal = 111 Then AGT = 111
If AggTotal = 112 Then AGT = 112
If AggTotal = 113 Then AGT = 113
If AggTotal = 114 Then AGT = 114
If AggTotal = 115 Then AGT = 115
If AggTotal = 116 Then AGT = 116
If AggTotal = 117 Then AGT = 117
If AggTotal = 118 Then AGT = 118
If AggTotal = 119 Then AGT = 119
If AggTotal = 120 Then AGT = 120
If AggTotal = 121 Then AGT = 121
If AggTotal = 122 Then AGT = 122
If AggTotal = 123 Then AGT = 123
If AggTotal = 124 Then AGT = 124
If AggTotal = 125 Then AGT = 125
If AggTotal = 126 Then AGT = 126
If AggTotal = 127 Then AGT = 127
If AggTotal = 128 Then AGT = 128
If AggTotal = 129 Then AGT = 129
If AggTotal = 130 Then AGT = 130
If AggTotal = 131 Then AGT = 131
If AggTotal = 132 Then AGT = 132
If AggTotal = 133 Then AGT = 133
If AggTotal = 134 Then AGT = 134
If AggTotal = 135 Then AGT = 135
If AggTotal = 136 Then AGT = 136
If AggTotal = 137 Then AGT = 137
If AggTotal = 138 Then AGT = 138
If AggTotal = 139 Then AGT = 139
If AggTotal = 140 Then AGT = 140
If AggTotal = 141 Then AGT = 141
If AggTotal = 142 Then AGT = 142
If AggTotal = 143 Then AGT = 143
If AggTotal = 144 Then AGT = 144
If AggTotal = 145 Then AGT = 145
If AggTotal = 146 Then AGT = 146
If AggTotal = 147 Then AGT = 147
If AggTotal = 148 Then AGT = 148
If AggTotal = 149 Then AGT = 149
If AggTotal = 150 Then AGT = 150
If AggTotal = 151 Then AGT = 151
If AggTotal = 152 Then AGT = 152
If AggTotal = 153 Then AGT = 153
If AggTotal = 154 Then AGT = 154
If AggTotal = 155 Then AGT = 155
If AggTotal = 156 Then AGT = 156
If AggTotal = 157 Then AGT = 157
If AggTotal = 158 Then AGT = 158
If AggTotal = 159 Then AGT = 159
If AggTotal = 160 Then AGT = 160
If AggTotal = 161 Then AGT = 161
If AggTotal = 162 Then AGT = 162
If AggTotal = 163 Then AGT = 163
If AggTotal = 164 Then AGT = 164
If AggTotal = 165 Then AGT = 165
If AggTotal = 166 Then AGT = 166
If AggTotal = 167 Then AGT = 167
If AggTotal = 168 Then AGT = 168
If AggTotal = 169 Then AGT = 169
If AggTotal = 170 Then AGT = 170
If AggTotal = 171 Then AGT = 171
If AggTotal = 172 Then AGT = 172
If AggTotal = 173 Then AGT = 173
If AggTotal = 174 Then AGT = 174
If AggTotal = 175 Then AGT = 175
If AggTotal = 176 Then AGT = 176
If AggTotal = 177 Then AGT = 177
If AggTotal = 178 Then AGT = 178
If AggTotal = 179 Then AGT = 179
If AggTotal = 180 Then AGT = 180
If AggTotal = 181 Then AGT = 181
If AggTotal = 182 Then AGT = 182
If AggTotal = 183 Then AGT = 183
If AggTotal = 184 Then AGT = 184
If AggTotal = 185 Then AGT = 185
If AggTotal = 186 Then AGT = 186
If AggTotal = 187 Then AGT = 187
If AggTotal = 188 Then AGT = 188
If AggTotal = 189 Then AGT = 189
If AggTotal = 190 Then AGT = 190
If AggTotal = 191 Then AGT = 191
If AggTotal = 192 Then AGT = 192
If AggTotal = 193 Then AGT = 193
If AggTotal = 194 Then AGT = 194
If AggTotal = 195 Then AGT = 195
If AggTotal = 196 Then AGT = 196
If AggTotal = 197 Then AGT = 197
If AggTotal = 198 Then AGT = 198
If AggTotal = 199 Then AGT = 199
If AggTotal = 200 Then AGT = 200
If AggTotal = 201 Then AGT = 201
If AggTotal = 202 Then AGT = 202
If AggTotal = 203 Then AGT = 203
If AggTotal = 204 Then AGT = 204
If AggTotal = 205 Then AGT = 205
If AggTotal = 206 Then AGT = 206
If AggTotal = 207 Then AGT = 207
If AggTotal = 208 Then AGT = 208
If AggTotal = 209 Then AGT = 209
If AggTotal = 210 Then AGT = 210
If AggTotal = 211 Then AGT = 211
If AggTotal = 212 Then AGT = 212
If AggTotal = 213 Then AGT = 213
If AggTotal = 214 Then AGT = 214
If AggTotal = 215 Then AGT = 215
If AggTotal = 216 Then AGT = 216
If AggTotal = 217 Then AGT = 217
If AggTotal = 218 Then AGT = 218
If AggTotal = 219 Then AGT = 219
If AggTotal = 220 Then AGT = 220
If AggTotal = 221 Then AGT = 221
If AggTotal = 222 Then AGT = 222
If AggTotal = 223 Then AGT = 223
If AggTotal = 224 Then AGT = 224
If AggTotal = 225 Then AGT = 225
If AggTotal = 226 Then AGT = 226
If AggTotal = 227 Then AGT = 227
If AggTotal = 228 Then AGT = 228
If AggTotal = 229 Then AGT = 229
If AggTotal = 230 Then AGT = 230
If AggTotal = 231 Then AGT = 231
If AggTotal = 232 Then AGT = 232
If AggTotal = 233 Then AGT = 233
If AggTotal = 234 Then AGT = 234
If AggTotal = 235 Then AGT = 235
If AggTotal = 236 Then AGT = 236
If AggTotal = 237 Then AGT = 237
If AggTotal = 238 Then AGT = 238
If AggTotal = 239 Then AGT = 239
If AggTotal = 240 Then AGT = 240
If AggTotal = 241 Then AGT = 241
If AggTotal = 242 Then AGT = 242
If AggTotal = 243 Then AGT = 243
If AggTotal = 244 Then AGT = 244
If AggTotal = 245 Then AGT = 245
If AggTotal = 246 Then AGT = 246
If AggTotal = 247 Then AGT = 247
If AggTotal = 248 Then AGT = 248
If AggTotal = 249 Then AGT = 249
If AggTotal = 250 Then AGT = 250
If AggTotal = 251 Then AGT = 251
If AggTotal = 252 Then AGT = 252
If AggTotal = 252 Then AGT = 253
If AggTotal = 254 Then AGT = 254
If AggTotal = 255 Then AGT = 255
If AggTotal = 256 Then AGT = 256
If AggTotal = 257 Then AGT = 257
If AggTotal = 258 Then AGT = 258
If AggTotal = 259 Then AGT = 259
If AggTotal = 260 Then AGT = 260
If AggTotal = 261 Then AGT = 261
If AggTotal = 262 Then AGT = 262
If AggTotal = 263 Then AGT = 263
If AggTotal = 264 Then AGT = 264
If AggTotal = 265 Then AGT = 265
If AggTotal = 266 Then AGT = 266
If AggTotal = 267 Then AGT = 267
If AggTotal = 268 Then AGT = 268
If AggTotal = 269 Then AGT = 269
If AggTotal = 270 Then AGT = 270
If AggTotal = 271 Then AGT = 271
If AggTotal = 272 Then AGT = 272
If AggTotal = 273 Then AGT = 273
If AggTotal = 274 Then AGT = 274
If AggTotal = 275 Then AGT = 275
If AggTotal = 276 Then AGT = 276
If AggTotal = 277 Then AGT = 277
If AggTotal = 278 Then AGT = 278
If AggTotal = 279 Then AGT = 279
End Function

Public Function SystemRestart()
StartKey = 0
FB = 1
clc
Stack = 0
NUMENT = 0
RANGE = 0
Prange = 0
AVRANGE = 0
AVRANGEA = 0
Hrange = 0
LRange = 49
Lsum = 279
Hsum = 0
PAGT = 0
CL = 0
AGTA = 0
AGTO = 0
WKK = 0
TRB(1) = 0
TRB(2) = 0
TRB(3) = 0
TRB(4) = 0
TRB(5) = 0
TRB(6) = 0
RB(1) = 0
RB(2) = 0
RB(3) = 0
RB(4) = 0
RB(5) = 0
RB(6) = 0
TB(1) = 0
TB(2) = 0
TB(3) = 0
TB(4) = 0
TB(5) = 0
TB(0) = 0
TB(6) = 0
For J = 1 To 49
B(J) = 0
P(J) = 0
Per(J) = 0
txtBall(J).Caption = Str(B(J))
Next
If LsetRan = 1 Then
LotBalls
End If
If LsetAct = 1 Then
FB = 0
ActBalls
End If
If SetRan = 1 Then
playerballs
PrangeCalc
End If
If SetFix = 1 Then
FixedPlayerBalls
End If
Order
If ECN = 0 And EHN = 0 Then GoTo 113
If LsetAct = 1 And NUMENT < 2 Then
For J = 1 To 6
lblPb(J).Caption = Str(PB(J))
Next
End If
113
Matcher
displayNums
NumDisplay
lblSum.Caption = Str(PAGT)
lblran.Caption = Str(Prange)
If Stack <= 1 Then
For J = 1 To 6
RB(J) = LB(J)
txtH(J).Caption = Str(LB(J))
Next
End If
End Function
Public Function SortLottery()
L = 1
If LB(1) > LB(2) Then L = L + 1
If LB(1) > LB(3) Then L = L + 1
If LB(1) > LB(4) Then L = L + 1
If LB(1) > LB(5) Then L = L + 1
If LB(1) > LB(6) Then L = L + 1
If L = 1 Then H(1) = LB(1)
If L = 2 Then H(2) = LB(1)
If L = 3 Then H(3) = LB(1)
If L = 4 Then H(4) = LB(1)
If L = 5 Then H(5) = LB(1)
If L = 6 Then H(6) = LB(1)
L = 1
If LB(2) > LB(1) Then L = L + 1
If LB(2) > LB(3) Then L = L + 1
If LB(2) > LB(4) Then L = L + 1
If LB(2) > LB(5) Then L = L + 1
If LB(2) > LB(6) Then L = L + 1
If L = 1 Then H(1) = LB(2)
If L = 2 Then H(2) = LB(2)
If L = 3 Then H(3) = LB(2)
If L = 4 Then H(4) = LB(2)
If L = 5 Then H(5) = LB(2)
If L = 6 Then H(6) = LB(2)
L = 1
If LB(3) > LB(1) Then L = L + 1
If LB(3) > LB(2) Then L = L + 1
If LB(3) > LB(4) Then L = L + 1
If LB(3) > LB(5) Then L = L + 1
If LB(3) > LB(6) Then L = L + 1
If L = 1 Then H(1) = LB(3)
If L = 2 Then H(2) = LB(3)
If L = 3 Then H(3) = LB(3)
If L = 4 Then H(4) = LB(3)
If L = 5 Then H(5) = LB(3)
If L = 6 Then H(6) = LB(3)
L = 1
If LB(4) > LB(1) Then L = L + 1
If LB(4) > LB(2) Then L = L + 1
If LB(4) > LB(3) Then L = L + 1
If LB(4) > LB(5) Then L = L + 1
If LB(4) > LB(6) Then L = L + 1
If L = 1 Then H(1) = LB(4)
If L = 2 Then H(2) = LB(4)
If L = 3 Then H(3) = LB(4)
If L = 4 Then H(4) = LB(4)
If L = 5 Then H(5) = LB(4)
If L = 6 Then H(6) = LB(4)
L = 1
If LB(5) > LB(1) Then L = L + 1
If LB(5) > LB(2) Then L = L + 1
If LB(5) > LB(3) Then L = L + 1
If LB(5) > LB(4) Then L = L + 1
If LB(5) > LB(6) Then L = L + 1
If L = 1 Then H(1) = LB(5)
If L = 2 Then H(2) = LB(5)
If L = 3 Then H(3) = LB(5)
If L = 4 Then H(4) = LB(5)
If L = 5 Then H(5) = LB(5)
If L = 6 Then H(6) = LB(5)
L = 1
If LB(6) > LB(1) Then L = L + 1
If LB(6) > LB(2) Then L = L + 1
If LB(6) > LB(3) Then L = L + 1
If LB(6) > LB(4) Then L = L + 1
If LB(6) > LB(5) Then L = L + 1
If L = 1 Then H(1) = LB(6)
If L = 2 Then H(2) = LB(6)
If L = 3 Then H(3) = LB(6)
If L = 4 Then H(4) = LB(6)
If L = 5 Then H(5) = LB(6)
If L = 6 Then H(6) = LB(6)
LB(1) = H(1)
LB(2) = H(2)
LB(3) = H(3)
LB(4) = H(4)
LB(5) = H(5)
LB(6) = H(6)
Random6
End Function

Public Function SortPlayer()
Q = 1
If PB(1) > PB(2) Then Q = Q + 1
If PB(1) > PB(3) Then Q = Q + 1
If PB(1) > PB(4) Then Q = Q + 1
If PB(1) > PB(5) Then Q = Q + 1
If PB(1) > PB(6) Then Q = Q + 1
If Q = 1 Then H(1) = PB(1)
If Q = 2 Then H(2) = PB(1)
If Q = 3 Then H(3) = PB(1)
If Q = 4 Then H(4) = PB(1)
If Q = 5 Then H(5) = PB(1)
If Q = 6 Then H(6) = PB(1)
Q = 1
If PB(2) > PB(1) Then Q = Q + 1
If PB(2) > PB(3) Then Q = Q + 1
If PB(2) > PB(4) Then Q = Q + 1
If PB(2) > PB(5) Then Q = Q + 1
If PB(2) > PB(6) Then Q = Q + 1
If Q = 1 Then H(1) = PB(2)
If Q = 2 Then H(2) = PB(2)
If Q = 3 Then H(3) = PB(2)
If Q = 4 Then H(4) = PB(2)
If Q = 5 Then H(5) = PB(2)
If Q = 6 Then H(6) = PB(2)
Q = 1
If PB(3) > PB(1) Then Q = Q + 1
If PB(3) > PB(2) Then Q = Q + 1
If PB(3) > PB(4) Then Q = Q + 1
If PB(3) > PB(5) Then Q = Q + 1
If PB(3) > PB(6) Then Q = Q + 1
If Q = 1 Then H(1) = PB(3)
If Q = 2 Then H(2) = PB(3)
If Q = 3 Then H(3) = PB(3)
If Q = 4 Then H(4) = PB(3)
If Q = 5 Then H(5) = PB(3)
If Q = 6 Then H(6) = PB(3)
Q = 1
If PB(4) > PB(1) Then Q = Q + 1
If PB(4) > PB(2) Then Q = Q + 1
If PB(4) > PB(3) Then Q = Q + 1
If PB(4) > PB(5) Then Q = Q + 1
If PB(4) > PB(6) Then Q = Q + 1
If Q = 1 Then H(1) = PB(4)
If Q = 2 Then H(2) = PB(4)
If Q = 3 Then H(3) = PB(4)
If Q = 4 Then H(4) = PB(4)
If Q = 5 Then H(5) = PB(4)
If Q = 6 Then H(6) = PB(4)
Q = 1
If PB(5) > PB(1) Then Q = Q + 1
If PB(5) > PB(2) Then Q = Q + 1
If PB(5) > PB(3) Then Q = Q + 1
If PB(5) > PB(4) Then Q = Q + 1
If PB(5) > PB(6) Then Q = Q + 1
If Q = 1 Then H(1) = PB(5)
If Q = 2 Then H(2) = PB(5)
If Q = 3 Then H(3) = PB(5)
If Q = 4 Then H(4) = PB(5)
If Q = 5 Then H(5) = PB(5)
If Q = 6 Then H(6) = PB(5)
Q = 1
If PB(6) > PB(1) Then Q = Q + 1
If PB(6) > PB(2) Then Q = Q + 1
If PB(6) > PB(3) Then Q = Q + 1
If PB(6) > PB(4) Then Q = Q + 1
If PB(6) > PB(5) Then Q = Q + 1
If Q = 1 Then H(1) = PB(6)
If Q = 2 Then H(2) = PB(6)
If Q = 3 Then H(3) = PB(6)
If Q = 4 Then H(4) = PB(6)
If Q = 5 Then H(5) = PB(6)
If Q = 6 Then H(6) = PB(6)
PB(1) = H(1)
PB(2) = H(2)
PB(3) = H(3)
PB(4) = H(4)
PB(5) = H(5)
PB(6) = H(6)
End Function
Public Function SetUpHotNumbers()
SetRan = 0
SetFix = 0
EHN = 1
ECN = 0
lblPlayerStat.Caption = "(Hottest Numbers)"
End Function

Public Function SetUpColdNumbers()
ECN = 1
EHN = 0
SetRan = 0
SetFix = 0
lblPlayerStat.Caption = "(Coldest Numbers)"
If NUMENT < (Stake * 16) And ECN = 1 Then
MsgBox "Recomendation!,  If you have not carried out many draws yet, there can be up to 44 Cold Numbers to choose from.  Suggest running a few more draws - just to get some more result data. Otherwise there can be some discrepency in which cold numbers (Numbers that have not been drawn yet) to choose from."
End If

End Function


Public Function FixedPlayerBalls()
PB(1) = SF(1)
PB(2) = SF(2)
PB(3) = SF(3)
PB(4) = SF(4)
PB(5) = SF(5)
PB(6) = SF(6)
End Function

Private Sub txtOddsL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 30
FrmQuickHelp.Show
End If
End Sub


Private Sub txtOddsLa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 31
FrmQuickHelp.Show
End If
End Sub

Private Sub txtPerl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 26
FrmQuickHelp.Show
End If
End Sub


Private Sub txtPWn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 27
FrmQuickHelp.Show
End If
End Sub


Public Function RangeFilter()
RP(1) = 0
RP(2) = 0
RP(3) = 0
RP(4) = 0
RP(5) = 0
RP(6) = 0
RP(0) = 0
If LB(1) > LB(2) And LB(1) > LB(3) And LB(1) > LB(4) And LB(1) > LB(5) And LB(1) > LB(6) And LB(1) > LB(0) Then
RP(1) = LB(1)
End If
If LB(2) > LB(1) And LB(2) > LB(3) And LB(2) > LB(4) And LB(2) > LB(5) And LB(2) > LB(6) And LB(2) > LB(0) Then
RP(1) = LB(2)
End If
If LB(3) > LB(1) And LB(3) > LB(2) And LB(3) > LB(4) And LB(3) > LB(5) And LB(3) > LB(6) And LB(3) > LB(0) Then
RP(1) = LB(3)
End If
If LB(4) > LB(1) And LB(4) > LB(2) And LB(4) > LB(3) And LB(4) > LB(5) And LB(4) > LB(6) And LB(4) > LB(0) Then
RP(1) = LB(4)
End If
If LB(5) > LB(1) And LB(5) > LB(2) And LB(5) > LB(3) And LB(5) > LB(4) And LB(5) > LB(6) And LB(5) > LB(0) Then
RP(1) = LB(5)
End If
If LB(6) > LB(1) And LB(6) > LB(2) And LB(6) > LB(3) And LB(6) > LB(4) And LB(6) > LB(5) And LB(6) > LB(0) Then
RP(1) = LB(6)
End If
If LB(0) > LB(1) And LB(0) > LB(2) And LB(0) > LB(3) And LB(0) > LB(4) And LB(0) > LB(5) And LB(0) > LB(6) Then
RP(1) = LB(0)
End If
If LB(1) < LB(2) And LB(1) < LB(3) And LB(1) < LB(4) And LB(1) < LB(5) And LB(1) < LB(6) And LB(1) < LB(0) Then
RP(0) = LB(1)
End If
If LB(2) < LB(1) And LB(2) < LB(3) And LB(2) < LB(4) And LB(2) < LB(5) And LB(2) < LB(6) And LB(2) < LB(0) Then
RP(0) = LB(2)
End If
If LB(3) < LB(1) And LB(3) < LB(2) And LB(3) < LB(4) And LB(3) < LB(5) And LB(3) < LB(6) And LB(3) < LB(0) Then
RP(0) = LB(3)
End If
If LB(4) < LB(1) And LB(4) < LB(2) And LB(4) < LB(3) And LB(4) < LB(5) And LB(4) < LB(6) And LB(4) < LB(0) Then
RP(0) = LB(4)
End If
If LB(5) < LB(1) And LB(5) < LB(2) And LB(5) < LB(3) And LB(5) < LB(4) And LB(5) < LB(6) And LB(5) < LB(0) Then
RP(0) = LB(5)
End If
If LB(6) < LB(1) And LB(6) < LB(2) And LB(6) < LB(3) And LB(6) < LB(4) And LB(6) < LB(5) And LB(6) < LB(0) Then
RP(0) = LB(6)
End If
If LB(0) < LB(1) And LB(0) < LB(2) And LB(0) < LB(3) And LB(0) < LB(4) And LB(0) < LB(5) And LB(0) < LB(6) Then
RP(0) = LB(0)
End If
RANGE = RP(1) - RP(0)
If Hrange <= RANGE Then
Hrange = RANGE
End If
If LRange >= RANGE Then
LRange = RANGE
End If
AVRANGEA = AVRANGEA + RANGE
AVRANGE = AVRANGEA / Stack
End Function

Public Function PrangeCalc()
Prange = 0
RP(1) = 0
RP(2) = 0
RP(3) = 0
RP(4) = 0
RP(5) = 0
RP(6) = 0
RP(0) = 0
If PB(1) > PB(2) And PB(1) > PB(3) And PB(1) > PB(4) And PB(1) > PB(5) And PB(1) > PB(6) Then
RP(1) = PB(1)
End If
If PB(2) > PB(1) And PB(2) > PB(3) And PB(2) > PB(4) And PB(2) > PB(5) And PB(2) > PB(6) Then
RP(1) = PB(2)
End If
If PB(3) > PB(1) And PB(3) > PB(2) And PB(3) > PB(4) And PB(3) > PB(5) And PB(3) > PB(6) Then
RP(1) = PB(3)
End If
If PB(4) > PB(1) And PB(4) > PB(2) And PB(4) > PB(3) And PB(4) > PB(5) And PB(4) > PB(6) Then
RP(1) = PB(4)
End If
If PB(5) > PB(1) And PB(5) > PB(2) And PB(5) > PB(3) And PB(5) > PB(4) And PB(5) > PB(6) Then
RP(1) = PB(5)
End If
If PB(6) > PB(1) And PB(6) > PB(2) And PB(6) > PB(3) And PB(6) > PB(4) And PB(6) > PB(5) Then
RP(1) = PB(6)
End If
If PB(1) < PB(2) And PB(1) < PB(3) And PB(1) < PB(4) And PB(1) < PB(5) And PB(1) < PB(6) Then
RP(6) = PB(1)
End If
If PB(2) < PB(1) And PB(2) < PB(3) And PB(2) < PB(4) And PB(2) < PB(5) And PB(2) < PB(6) Then
RP(6) = PB(2)
End If
If PB(3) < PB(1) And PB(3) < PB(2) And PB(3) < PB(4) And PB(3) < PB(5) And PB(3) < PB(6) Then
RP(6) = PB(3)
End If
If PB(4) < PB(1) And PB(4) < PB(2) And PB(4) < PB(3) And PB(4) < PB(5) And PB(4) < PB(6) Then
RP(6) = PB(4)
End If
If PB(5) < PB(1) And PB(5) < PB(2) And PB(5) < PB(3) And PB(5) < PB(4) And PB(5) < PB(6) Then
RP(6) = PB(5)
End If
If PB(6) < PB(1) And PB(6) < PB(2) And PB(6) < PB(3) And PB(6) < PB(4) And PB(6) < PB(5) Then
RP(6) = PB(6)
End If
Prange = RP(1) - RP(6)
End Function


Public Function EndSeq()
A = 12
D = 13
If PB(1) = LB(1) Then
lbllb(1).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(1) Then
lbllb(1).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(1) Then
lbllb(1).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(1) Then
lbllb(1).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(1) Then
lbllb(1).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(1) Then
lbllb(1).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(2) Then
lbllb(2).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(2) Then
lbllb(2).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(2) Then
lbllb(2).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(2) Then
lbllb(2).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(2) Then
lbllb(2).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(2) Then
lbllb(2).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(3) Then
lbllb(3).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(3) Then
lbllb(3).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(3) Then
lbllb(3).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(3) Then
lbllb(3).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(3) Then
lbllb(3).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(3) Then
lbllb(3).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(4) Then
lbllb(4).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(4) Then
lbllb(4).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(4) Then
lbllb(4).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(4) Then
lbllb(4).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(4) Then
lbllb(4).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(4) Then
lbllb(4).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(5) Then
lbllb(5).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(5) Then
lbllb(5).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(5) Then
lbllb(5).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(5) Then
lbllb(5).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(5) Then
lbllb(5).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(5) Then
lbllb(5).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
If PB(1) = LB(6) Then
lbllb(6).BackColor = QBColor(A)
lblPb(1).BackColor = QBColor(A)
End If
If PB(2) = LB(6) Then
lbllb(6).BackColor = QBColor(A)
lblPb(2).BackColor = QBColor(A)
End If
If PB(3) = LB(6) Then
lbllb(6).BackColor = QBColor(A)
lblPb(3).BackColor = QBColor(A)
End If
If PB(4) = LB(6) Then
lbllb(6).BackColor = QBColor(A)
lblPb(4).BackColor = QBColor(A)
End If
If PB(5) = LB(6) Then
lbllb(6).BackColor = QBColor(A)
lblPb(5).BackColor = QBColor(A)
End If
If PB(6) = LB(6) Then
lbllb(6).BackColor = QBColor(A)
lblPb(6).BackColor = QBColor(A)
End If
NumDisplay
displayNums
End Function
Public Function Order2()
M = 21: L = 1
If B(21) < B(1) Then L = L + 1
If B(21) < B(2) Then L = L + 1
If B(21) < B(3) Then L = L + 1
If B(21) < B(4) Then L = L + 1
If B(21) < B(5) Then L = L + 1
If B(21) < B(6) Then L = L + 1
If B(21) < B(7) Then L = L + 1
If B(21) < B(8) Then L = L + 1
If B(21) < B(9) Then L = L + 1
If B(21) < B(10) Then L = L + 1
If B(21) < B(11) Then L = L + 1
If B(21) < B(12) Then L = L + 1
If B(21) < B(13) Then L = L + 1
If B(21) < B(14) Then L = L + 1
If B(21) < B(15) Then L = L + 1
If B(21) < B(16) Then L = L + 1
If B(21) < B(17) Then L = L + 1
If B(21) < B(18) Then L = L + 1
If B(21) < B(19) Then L = L + 1
If B(21) < B(20) Then L = L + 1
If B(21) <= B(22) Then L = L + 1
If B(21) <= B(23) Then L = L + 1
If B(21) <= B(24) Then L = L + 1
If B(21) <= B(25) Then L = L + 1
If B(21) <= B(26) Then L = L + 1
If B(21) <= B(27) Then L = L + 1
If B(21) <= B(28) Then L = L + 1
If B(21) <= B(29) Then L = L + 1
If B(21) <= B(30) Then L = L + 1
If B(21) <= B(31) Then L = L + 1
If B(21) <= B(32) Then L = L + 1
If B(21) <= B(33) Then L = L + 1
If B(21) <= B(34) Then L = L + 1
If B(21) <= B(35) Then L = L + 1
If B(21) <= B(36) Then L = L + 1
If B(21) <= B(37) Then L = L + 1
If B(21) <= B(38) Then L = L + 1
If B(21) <= B(39) Then L = L + 1
If B(21) <= B(40) Then L = L + 1
If B(21) <= B(41) Then L = L + 1
If B(21) <= B(42) Then L = L + 1
If B(21) <= B(43) Then L = L + 1
If B(21) <= B(44) Then L = L + 1
If B(21) <= B(45) Then L = L + 1
If B(21) <= B(46) Then L = L + 1
If B(21) <= B(47) Then L = L + 1
If B(21) <= B(48) Then L = L + 1
If B(21) <= B(49) Then L = L + 1
 orderdisp
M = 22: L = 1
If B(22) < B(1) Then L = L + 1
If B(22) < B(2) Then L = L + 1
If B(22) < B(3) Then L = L + 1
If B(22) < B(4) Then L = L + 1
If B(22) < B(5) Then L = L + 1
If B(22) < B(6) Then L = L + 1
If B(22) < B(7) Then L = L + 1
If B(22) < B(8) Then L = L + 1
If B(22) < B(9) Then L = L + 1
If B(22) < B(10) Then L = L + 1
If B(22) < B(11) Then L = L + 1
If B(22) < B(12) Then L = L + 1
If B(22) < B(13) Then L = L + 1
If B(22) < B(14) Then L = L + 1
If B(22) < B(15) Then L = L + 1
If B(22) < B(16) Then L = L + 1
If B(22) < B(17) Then L = L + 1
If B(22) < B(18) Then L = L + 1
If B(22) < B(19) Then L = L + 1
If B(22) < B(20) Then L = L + 1
If B(22) < B(21) Then L = L + 1
If B(22) <= B(23) Then L = L + 1
If B(22) <= B(24) Then L = L + 1
If B(22) <= B(25) Then L = L + 1
If B(22) <= B(26) Then L = L + 1
If B(22) <= B(27) Then L = L + 1
If B(22) <= B(28) Then L = L + 1
If B(22) <= B(29) Then L = L + 1
If B(22) <= B(30) Then L = L + 1
If B(22) <= B(31) Then L = L + 1
If B(22) <= B(32) Then L = L + 1
If B(22) <= B(33) Then L = L + 1
If B(22) <= B(34) Then L = L + 1
If B(22) <= B(35) Then L = L + 1
If B(22) <= B(36) Then L = L + 1
If B(22) <= B(37) Then L = L + 1
If B(22) <= B(38) Then L = L + 1
If B(22) <= B(39) Then L = L + 1
If B(22) <= B(40) Then L = L + 1
If B(22) <= B(41) Then L = L + 1
If B(22) <= B(42) Then L = L + 1
If B(22) <= B(43) Then L = L + 1
If B(22) <= B(44) Then L = L + 1
If B(22) <= B(45) Then L = L + 1
If B(22) <= B(46) Then L = L + 1
If B(22) <= B(47) Then L = L + 1
If B(22) <= B(48) Then L = L + 1
If B(22) <= B(49) Then L = L + 1
 orderdisp
M = 23: L = 1
If B(23) < B(1) Then L = L + 1
If B(23) < B(2) Then L = L + 1
If B(23) < B(3) Then L = L + 1
If B(23) < B(4) Then L = L + 1
If B(23) < B(5) Then L = L + 1
If B(23) < B(6) Then L = L + 1
If B(23) < B(7) Then L = L + 1
If B(23) < B(8) Then L = L + 1
If B(23) < B(9) Then L = L + 1
If B(23) < B(10) Then L = L + 1
If B(23) < B(11) Then L = L + 1
If B(23) < B(12) Then L = L + 1
If B(23) < B(13) Then L = L + 1
If B(23) < B(14) Then L = L + 1
If B(23) < B(15) Then L = L + 1
If B(23) < B(16) Then L = L + 1
If B(23) < B(17) Then L = L + 1
If B(23) < B(18) Then L = L + 1
If B(23) < B(19) Then L = L + 1
If B(23) < B(20) Then L = L + 1
If B(23) < B(21) Then L = L + 1
If B(23) < B(22) Then L = L + 1
If B(23) <= B(24) Then L = L + 1
If B(23) <= B(25) Then L = L + 1
If B(23) <= B(26) Then L = L + 1
If B(23) <= B(27) Then L = L + 1
If B(23) <= B(28) Then L = L + 1
If B(23) <= B(29) Then L = L + 1
If B(23) <= B(30) Then L = L + 1
If B(23) <= B(31) Then L = L + 1
If B(23) <= B(32) Then L = L + 1
If B(23) <= B(33) Then L = L + 1
If B(23) <= B(34) Then L = L + 1
If B(23) <= B(35) Then L = L + 1
If B(23) <= B(36) Then L = L + 1
If B(23) <= B(37) Then L = L + 1
If B(23) <= B(38) Then L = L + 1
If B(23) <= B(39) Then L = L + 1
If B(23) <= B(40) Then L = L + 1
If B(23) <= B(41) Then L = L + 1
If B(23) <= B(42) Then L = L + 1
If B(23) <= B(43) Then L = L + 1
If B(23) <= B(44) Then L = L + 1
If B(23) <= B(45) Then L = L + 1
If B(23) <= B(46) Then L = L + 1
If B(23) <= B(47) Then L = L + 1
If B(23) <= B(48) Then L = L + 1
If B(23) <= B(49) Then L = L + 1
 orderdisp
M = 24: L = 1
If B(24) < B(1) Then L = L + 1
If B(24) < B(2) Then L = L + 1
If B(24) < B(3) Then L = L + 1
If B(24) < B(4) Then L = L + 1
If B(24) < B(5) Then L = L + 1
If B(24) < B(6) Then L = L + 1
If B(24) < B(7) Then L = L + 1
If B(24) < B(8) Then L = L + 1
If B(24) < B(9) Then L = L + 1
If B(24) < B(10) Then L = L + 1
If B(24) < B(11) Then L = L + 1
If B(24) < B(12) Then L = L + 1
If B(24) < B(13) Then L = L + 1
If B(24) < B(14) Then L = L + 1
If B(24) < B(15) Then L = L + 1
If B(24) < B(16) Then L = L + 1
If B(24) < B(17) Then L = L + 1
If B(24) < B(18) Then L = L + 1
If B(24) < B(19) Then L = L + 1
If B(24) < B(20) Then L = L + 1
If B(24) < B(21) Then L = L + 1
If B(24) < B(22) Then L = L + 1
If B(24) < B(23) Then L = L + 1
If B(24) <= B(25) Then L = L + 1
If B(24) <= B(26) Then L = L + 1
If B(24) <= B(27) Then L = L + 1
If B(24) <= B(28) Then L = L + 1
If B(24) <= B(29) Then L = L + 1
If B(24) <= B(30) Then L = L + 1
If B(24) <= B(31) Then L = L + 1
If B(24) <= B(32) Then L = L + 1
If B(24) <= B(33) Then L = L + 1
If B(24) <= B(34) Then L = L + 1
If B(24) <= B(35) Then L = L + 1
If B(24) <= B(36) Then L = L + 1
If B(24) <= B(37) Then L = L + 1
If B(24) <= B(38) Then L = L + 1
If B(24) <= B(39) Then L = L + 1
If B(24) <= B(40) Then L = L + 1
If B(24) <= B(41) Then L = L + 1
If B(24) <= B(42) Then L = L + 1
If B(24) <= B(43) Then L = L + 1
If B(24) <= B(44) Then L = L + 1
If B(24) <= B(45) Then L = L + 1
If B(24) <= B(46) Then L = L + 1
If B(24) <= B(47) Then L = L + 1
If B(24) <= B(48) Then L = L + 1
If B(24) <= B(49) Then L = L + 1
 orderdisp
M = 25: L = 1
If B(25) < B(1) Then L = L + 1
If B(25) < B(2) Then L = L + 1
If B(25) < B(3) Then L = L + 1
If B(25) < B(4) Then L = L + 1
If B(25) < B(5) Then L = L + 1
If B(25) < B(6) Then L = L + 1
If B(25) < B(7) Then L = L + 1
If B(25) < B(8) Then L = L + 1
If B(25) < B(9) Then L = L + 1
If B(25) < B(10) Then L = L + 1
If B(25) < B(11) Then L = L + 1
If B(25) < B(12) Then L = L + 1
If B(25) < B(13) Then L = L + 1
If B(25) < B(14) Then L = L + 1
If B(25) < B(15) Then L = L + 1
If B(25) < B(16) Then L = L + 1
If B(25) < B(17) Then L = L + 1
If B(25) < B(18) Then L = L + 1
If B(25) < B(19) Then L = L + 1
If B(25) < B(20) Then L = L + 1
If B(25) < B(21) Then L = L + 1
If B(25) < B(22) Then L = L + 1
If B(25) < B(23) Then L = L + 1
If B(25) < B(24) Then L = L + 1
If B(25) <= B(26) Then L = L + 1
If B(25) <= B(27) Then L = L + 1
If B(25) <= B(28) Then L = L + 1
If B(25) <= B(29) Then L = L + 1
If B(25) <= B(30) Then L = L + 1
If B(25) <= B(31) Then L = L + 1
If B(25) <= B(32) Then L = L + 1
If B(25) <= B(33) Then L = L + 1
If B(25) <= B(34) Then L = L + 1
If B(25) <= B(35) Then L = L + 1
If B(25) <= B(36) Then L = L + 1
If B(25) <= B(37) Then L = L + 1
If B(25) <= B(38) Then L = L + 1
If B(25) <= B(39) Then L = L + 1
If B(25) <= B(40) Then L = L + 1
If B(25) <= B(41) Then L = L + 1
If B(25) <= B(42) Then L = L + 1
If B(25) <= B(43) Then L = L + 1
If B(25) <= B(44) Then L = L + 1
If B(25) <= B(45) Then L = L + 1
If B(25) <= B(46) Then L = L + 1
If B(25) <= B(47) Then L = L + 1
If B(25) <= B(48) Then L = L + 1
If B(25) <= B(49) Then L = L + 1
orderdisp

M = 26: L = 1
If B(26) < B(1) Then L = L + 1
If B(26) < B(2) Then L = L + 1
If B(26) < B(3) Then L = L + 1
If B(26) < B(4) Then L = L + 1
If B(26) < B(5) Then L = L + 1
If B(26) < B(6) Then L = L + 1
If B(26) < B(7) Then L = L + 1
If B(26) < B(8) Then L = L + 1
If B(26) < B(9) Then L = L + 1
If B(26) < B(10) Then L = L + 1
If B(26) < B(11) Then L = L + 1
If B(26) < B(12) Then L = L + 1
If B(26) < B(13) Then L = L + 1
If B(26) < B(14) Then L = L + 1
If B(26) < B(15) Then L = L + 1
If B(26) < B(16) Then L = L + 1
If B(26) < B(17) Then L = L + 1
If B(26) < B(18) Then L = L + 1
If B(26) < B(19) Then L = L + 1
If B(26) < B(20) Then L = L + 1
If B(26) < B(21) Then L = L + 1
If B(26) < B(22) Then L = L + 1
If B(26) < B(23) Then L = L + 1
If B(26) < B(24) Then L = L + 1
If B(26) < B(25) Then L = L + 1
If B(26) <= B(27) Then L = L + 1
If B(26) <= B(28) Then L = L + 1
If B(26) <= B(29) Then L = L + 1
If B(26) <= B(30) Then L = L + 1
If B(26) <= B(31) Then L = L + 1
If B(26) <= B(32) Then L = L + 1
If B(26) <= B(33) Then L = L + 1
If B(26) <= B(34) Then L = L + 1
If B(26) <= B(35) Then L = L + 1
If B(26) <= B(36) Then L = L + 1
If B(26) <= B(37) Then L = L + 1
If B(26) <= B(38) Then L = L + 1
If B(26) <= B(39) Then L = L + 1
If B(26) <= B(40) Then L = L + 1
If B(26) <= B(41) Then L = L + 1
If B(26) <= B(42) Then L = L + 1
If B(26) <= B(43) Then L = L + 1
If B(26) <= B(44) Then L = L + 1
If B(26) <= B(45) Then L = L + 1
If B(26) <= B(46) Then L = L + 1
If B(26) <= B(47) Then L = L + 1
If B(26) <= B(48) Then L = L + 1
If B(26) <= B(49) Then L = L + 1
orderdisp
M = 27: L = 1
If B(27) < B(1) Then L = L + 1
If B(27) < B(2) Then L = L + 1
If B(27) < B(3) Then L = L + 1
If B(27) < B(4) Then L = L + 1
If B(27) < B(5) Then L = L + 1
If B(27) < B(6) Then L = L + 1
If B(27) < B(7) Then L = L + 1
If B(27) < B(8) Then L = L + 1
If B(27) < B(9) Then L = L + 1
If B(27) < B(10) Then L = L + 1
If B(27) < B(11) Then L = L + 1
If B(27) < B(12) Then L = L + 1
If B(27) < B(13) Then L = L + 1
If B(27) < B(14) Then L = L + 1
If B(27) < B(15) Then L = L + 1
If B(27) < B(16) Then L = L + 1
If B(27) < B(17) Then L = L + 1
If B(27) < B(18) Then L = L + 1
If B(27) < B(19) Then L = L + 1
If B(27) < B(20) Then L = L + 1
If B(27) < B(21) Then L = L + 1
If B(27) < B(22) Then L = L + 1
If B(27) < B(23) Then L = L + 1
If B(27) < B(24) Then L = L + 1
If B(27) < B(25) Then L = L + 1
If B(27) < B(26) Then L = L + 1
If B(27) <= B(28) Then L = L + 1
If B(27) <= B(29) Then L = L + 1
If B(27) <= B(30) Then L = L + 1
If B(27) <= B(31) Then L = L + 1
If B(27) <= B(32) Then L = L + 1
If B(27) <= B(33) Then L = L + 1
If B(27) <= B(34) Then L = L + 1
If B(27) <= B(35) Then L = L + 1
If B(27) <= B(36) Then L = L + 1
If B(27) <= B(37) Then L = L + 1
If B(27) <= B(38) Then L = L + 1
If B(27) <= B(39) Then L = L + 1
If B(27) <= B(40) Then L = L + 1
If B(27) <= B(41) Then L = L + 1
If B(27) <= B(42) Then L = L + 1
If B(27) <= B(43) Then L = L + 1
If B(27) <= B(44) Then L = L + 1
If B(27) <= B(45) Then L = L + 1
If B(27) <= B(46) Then L = L + 1
If B(27) <= B(47) Then L = L + 1
If B(27) <= B(48) Then L = L + 1
If B(27) <= B(49) Then L = L + 1
orderdisp
M = 28: L = 1
If B(28) < B(1) Then L = L + 1
If B(28) < B(2) Then L = L + 1
If B(28) < B(3) Then L = L + 1
If B(28) < B(4) Then L = L + 1
If B(28) < B(5) Then L = L + 1
If B(28) < B(6) Then L = L + 1
If B(28) < B(7) Then L = L + 1
If B(28) < B(8) Then L = L + 1
If B(28) < B(9) Then L = L + 1
If B(28) < B(10) Then L = L + 1
If B(28) < B(11) Then L = L + 1
If B(28) < B(12) Then L = L + 1
If B(28) < B(13) Then L = L + 1
If B(28) < B(14) Then L = L + 1
If B(28) < B(15) Then L = L + 1
If B(28) < B(16) Then L = L + 1
If B(28) < B(17) Then L = L + 1
If B(28) < B(18) Then L = L + 1
If B(28) < B(19) Then L = L + 1
If B(28) < B(20) Then L = L + 1
If B(28) < B(21) Then L = L + 1
If B(28) < B(22) Then L = L + 1
If B(28) < B(23) Then L = L + 1
If B(28) < B(24) Then L = L + 1
If B(28) < B(25) Then L = L + 1
If B(28) < B(26) Then L = L + 1
If B(28) < B(27) Then L = L + 1
If B(28) <= B(29) Then L = L + 1
If B(28) <= B(30) Then L = L + 1
If B(28) <= B(31) Then L = L + 1
If B(28) <= B(32) Then L = L + 1
If B(28) <= B(33) Then L = L + 1
If B(28) <= B(34) Then L = L + 1
If B(28) <= B(35) Then L = L + 1
If B(28) <= B(36) Then L = L + 1
If B(28) <= B(37) Then L = L + 1
If B(28) <= B(38) Then L = L + 1
If B(28) <= B(39) Then L = L + 1
If B(28) <= B(40) Then L = L + 1
If B(28) <= B(41) Then L = L + 1
If B(28) <= B(42) Then L = L + 1
If B(28) <= B(43) Then L = L + 1
If B(28) <= B(44) Then L = L + 1
If B(28) <= B(45) Then L = L + 1
If B(28) <= B(46) Then L = L + 1
If B(28) <= B(47) Then L = L + 1
If B(28) <= B(48) Then L = L + 1
If B(28) <= B(49) Then L = L + 1
orderdisp
M = 29: L = 1
If B(29) < B(1) Then L = L + 1
If B(29) < B(2) Then L = L + 1
If B(29) < B(3) Then L = L + 1
If B(29) < B(4) Then L = L + 1
If B(29) < B(5) Then L = L + 1
If B(29) < B(6) Then L = L + 1
If B(29) < B(7) Then L = L + 1
If B(29) < B(8) Then L = L + 1
If B(29) < B(9) Then L = L + 1
If B(29) < B(10) Then L = L + 1
If B(29) < B(11) Then L = L + 1
If B(29) < B(12) Then L = L + 1
If B(29) < B(13) Then L = L + 1
If B(29) < B(14) Then L = L + 1
If B(29) < B(15) Then L = L + 1
If B(29) < B(16) Then L = L + 1
If B(29) < B(17) Then L = L + 1
If B(29) < B(18) Then L = L + 1
If B(29) < B(19) Then L = L + 1
If B(29) < B(20) Then L = L + 1
If B(29) < B(21) Then L = L + 1
If B(29) < B(22) Then L = L + 1
If B(29) < B(23) Then L = L + 1
If B(29) < B(24) Then L = L + 1
If B(29) < B(25) Then L = L + 1
If B(29) < B(26) Then L = L + 1
If B(29) < B(27) Then L = L + 1
If B(29) < B(28) Then L = L + 1
If B(29) <= B(30) Then L = L + 1
If B(29) <= B(31) Then L = L + 1
If B(29) <= B(32) Then L = L + 1
If B(29) <= B(33) Then L = L + 1
If B(29) <= B(34) Then L = L + 1
If B(29) <= B(35) Then L = L + 1
If B(29) <= B(36) Then L = L + 1
If B(29) <= B(37) Then L = L + 1
If B(29) <= B(38) Then L = L + 1
If B(29) <= B(39) Then L = L + 1
If B(29) <= B(40) Then L = L + 1
If B(29) <= B(41) Then L = L + 1
If B(29) <= B(42) Then L = L + 1
If B(29) <= B(43) Then L = L + 1
If B(29) <= B(44) Then L = L + 1
If B(29) <= B(45) Then L = L + 1
If B(29) <= B(46) Then L = L + 1
If B(29) <= B(47) Then L = L + 1
If B(29) <= B(48) Then L = L + 1
If B(29) <= B(49) Then L = L + 1
orderdisp
M = 30: L = 1
If B(30) < B(1) Then L = L + 1
If B(30) < B(2) Then L = L + 1
If B(30) < B(3) Then L = L + 1
If B(30) < B(4) Then L = L + 1
If B(30) < B(5) Then L = L + 1
If B(30) < B(6) Then L = L + 1
If B(30) < B(7) Then L = L + 1
If B(30) < B(8) Then L = L + 1
If B(30) < B(9) Then L = L + 1
If B(30) < B(10) Then L = L + 1
If B(30) < B(11) Then L = L + 1
If B(30) < B(12) Then L = L + 1
If B(30) < B(13) Then L = L + 1
If B(30) < B(14) Then L = L + 1
If B(30) < B(15) Then L = L + 1
If B(30) < B(16) Then L = L + 1
If B(30) < B(17) Then L = L + 1
If B(30) < B(18) Then L = L + 1
If B(30) < B(19) Then L = L + 1
If B(30) < B(20) Then L = L + 1
If B(30) < B(21) Then L = L + 1
If B(30) < B(22) Then L = L + 1
If B(30) < B(23) Then L = L + 1
If B(30) < B(24) Then L = L + 1
If B(30) < B(25) Then L = L + 1
If B(30) < B(26) Then L = L + 1
If B(30) < B(27) Then L = L + 1
If B(30) < B(28) Then L = L + 1
If B(30) < B(29) Then L = L + 1
If B(30) <= B(31) Then L = L + 1
If B(30) <= B(32) Then L = L + 1
If B(30) <= B(33) Then L = L + 1
If B(30) <= B(34) Then L = L + 1
If B(30) <= B(35) Then L = L + 1
If B(30) <= B(36) Then L = L + 1
If B(30) <= B(37) Then L = L + 1
If B(30) <= B(38) Then L = L + 1
If B(30) <= B(39) Then L = L + 1
If B(30) <= B(40) Then L = L + 1
If B(30) <= B(41) Then L = L + 1
If B(30) <= B(42) Then L = L + 1
If B(30) <= B(43) Then L = L + 1
If B(30) <= B(44) Then L = L + 1
If B(30) <= B(45) Then L = L + 1
If B(30) <= B(46) Then L = L + 1
If B(30) <= B(47) Then L = L + 1
If B(30) <= B(48) Then L = L + 1
If B(30) <= B(49) Then L = L + 1
orderdisp
M = 31: L = 1
If B(31) < B(1) Then L = L + 1
If B(31) < B(2) Then L = L + 1
If B(31) < B(3) Then L = L + 1
If B(31) < B(4) Then L = L + 1
If B(31) < B(5) Then L = L + 1
If B(31) < B(6) Then L = L + 1
If B(31) < B(7) Then L = L + 1
If B(31) < B(8) Then L = L + 1
If B(31) < B(9) Then L = L + 1
If B(31) < B(10) Then L = L + 1
If B(31) < B(11) Then L = L + 1
If B(31) < B(12) Then L = L + 1
If B(31) < B(13) Then L = L + 1
If B(31) < B(14) Then L = L + 1
If B(31) < B(15) Then L = L + 1
If B(31) < B(16) Then L = L + 1
If B(31) < B(17) Then L = L + 1
If B(31) < B(18) Then L = L + 1
If B(31) < B(19) Then L = L + 1
If B(31) < B(20) Then L = L + 1
If B(31) < B(21) Then L = L + 1
If B(31) < B(22) Then L = L + 1
If B(31) < B(23) Then L = L + 1
If B(31) < B(24) Then L = L + 1
If B(31) < B(25) Then L = L + 1
If B(31) < B(26) Then L = L + 1
If B(31) < B(27) Then L = L + 1
If B(31) < B(28) Then L = L + 1
If B(31) < B(29) Then L = L + 1
If B(31) < B(30) Then L = L + 1
If B(31) <= B(32) Then L = L + 1
If B(31) <= B(33) Then L = L + 1
If B(31) <= B(34) Then L = L + 1
If B(31) <= B(35) Then L = L + 1
If B(31) <= B(36) Then L = L + 1
If B(31) <= B(37) Then L = L + 1
If B(31) <= B(38) Then L = L + 1
If B(31) <= B(39) Then L = L + 1
If B(31) <= B(40) Then L = L + 1
If B(31) <= B(41) Then L = L + 1
If B(31) <= B(42) Then L = L + 1
If B(31) <= B(43) Then L = L + 1
If B(31) <= B(44) Then L = L + 1
If B(31) <= B(45) Then L = L + 1
If B(31) <= B(46) Then L = L + 1
If B(31) <= B(47) Then L = L + 1
If B(31) <= B(48) Then L = L + 1
If B(31) <= B(49) Then L = L + 1
 orderdisp
M = 32: L = 1
If B(32) < B(1) Then L = L + 1
If B(32) < B(2) Then L = L + 1
If B(32) < B(3) Then L = L + 1
If B(32) < B(4) Then L = L + 1
If B(32) < B(5) Then L = L + 1
If B(32) < B(6) Then L = L + 1
If B(32) < B(7) Then L = L + 1
If B(32) < B(8) Then L = L + 1
If B(32) < B(9) Then L = L + 1
If B(32) < B(10) Then L = L + 1
If B(32) < B(11) Then L = L + 1
If B(32) < B(12) Then L = L + 1
If B(32) < B(13) Then L = L + 1
If B(32) < B(14) Then L = L + 1
If B(32) < B(15) Then L = L + 1
If B(32) < B(16) Then L = L + 1
If B(32) < B(17) Then L = L + 1
If B(32) < B(18) Then L = L + 1
If B(32) < B(19) Then L = L + 1
If B(32) < B(20) Then L = L + 1
If B(32) < B(21) Then L = L + 1
If B(32) < B(22) Then L = L + 1
If B(32) < B(23) Then L = L + 1
If B(32) < B(24) Then L = L + 1
If B(32) < B(25) Then L = L + 1
If B(32) < B(26) Then L = L + 1
If B(32) < B(27) Then L = L + 1
If B(32) < B(28) Then L = L + 1
If B(32) < B(29) Then L = L + 1
If B(32) < B(30) Then L = L + 1
If B(32) < B(31) Then L = L + 1
If B(32) <= B(33) Then L = L + 1
If B(32) <= B(34) Then L = L + 1
If B(32) <= B(35) Then L = L + 1
If B(32) <= B(36) Then L = L + 1
If B(32) <= B(37) Then L = L + 1
If B(32) <= B(38) Then L = L + 1
If B(32) <= B(39) Then L = L + 1
If B(32) <= B(40) Then L = L + 1
If B(32) <= B(41) Then L = L + 1
If B(32) <= B(42) Then L = L + 1
If B(32) <= B(43) Then L = L + 1
If B(32) <= B(44) Then L = L + 1
If B(32) <= B(45) Then L = L + 1
If B(32) <= B(46) Then L = L + 1
If B(32) <= B(47) Then L = L + 1
If B(32) <= B(48) Then L = L + 1
If B(32) <= B(49) Then L = L + 1
 orderdisp
M = 33: L = 1
If B(33) < B(1) Then L = L + 1
If B(33) < B(2) Then L = L + 1
If B(33) < B(3) Then L = L + 1
If B(33) < B(4) Then L = L + 1
If B(33) < B(5) Then L = L + 1
If B(33) < B(6) Then L = L + 1
If B(33) < B(7) Then L = L + 1
If B(33) < B(8) Then L = L + 1
If B(33) < B(9) Then L = L + 1
If B(33) < B(10) Then L = L + 1
If B(33) < B(11) Then L = L + 1
If B(33) < B(12) Then L = L + 1
If B(33) < B(13) Then L = L + 1
If B(33) < B(14) Then L = L + 1
If B(33) < B(15) Then L = L + 1
If B(33) < B(16) Then L = L + 1
If B(33) < B(17) Then L = L + 1
If B(33) < B(18) Then L = L + 1
If B(33) < B(19) Then L = L + 1
If B(33) < B(20) Then L = L + 1
If B(33) < B(21) Then L = L + 1
If B(33) < B(22) Then L = L + 1
If B(33) < B(23) Then L = L + 1
If B(33) < B(24) Then L = L + 1
If B(33) < B(25) Then L = L + 1
If B(33) < B(26) Then L = L + 1
If B(33) < B(27) Then L = L + 1
If B(33) < B(28) Then L = L + 1
If B(33) < B(29) Then L = L + 1
If B(33) < B(30) Then L = L + 1
If B(33) < B(31) Then L = L + 1
If B(33) < B(32) Then L = L + 1
If B(33) <= B(34) Then L = L + 1
If B(33) <= B(35) Then L = L + 1
If B(33) <= B(36) Then L = L + 1
If B(33) <= B(37) Then L = L + 1
If B(33) <= B(38) Then L = L + 1
If B(33) <= B(39) Then L = L + 1
If B(33) <= B(40) Then L = L + 1
If B(33) <= B(41) Then L = L + 1
If B(33) <= B(42) Then L = L + 1
If B(33) <= B(43) Then L = L + 1
If B(33) <= B(44) Then L = L + 1
If B(33) <= B(45) Then L = L + 1
If B(33) <= B(46) Then L = L + 1
If B(33) <= B(47) Then L = L + 1
If B(33) <= B(48) Then L = L + 1
If B(33) <= B(49) Then L = L + 1
 orderdisp
M = 34: L = 1
If B(34) < B(1) Then L = L + 1
If B(34) < B(2) Then L = L + 1
If B(34) < B(3) Then L = L + 1
If B(34) < B(4) Then L = L + 1
If B(34) < B(5) Then L = L + 1
If B(34) < B(6) Then L = L + 1
If B(34) < B(7) Then L = L + 1
If B(34) < B(8) Then L = L + 1
If B(34) < B(9) Then L = L + 1
If B(34) < B(10) Then L = L + 1
If B(34) < B(11) Then L = L + 1
If B(34) < B(12) Then L = L + 1
If B(34) < B(13) Then L = L + 1
If B(34) < B(14) Then L = L + 1
If B(34) < B(15) Then L = L + 1
If B(34) < B(16) Then L = L + 1
If B(34) < B(17) Then L = L + 1
If B(34) < B(18) Then L = L + 1
If B(34) < B(19) Then L = L + 1
If B(34) < B(20) Then L = L + 1
If B(34) < B(21) Then L = L + 1
If B(34) < B(22) Then L = L + 1
If B(34) < B(23) Then L = L + 1
If B(34) < B(24) Then L = L + 1
If B(34) < B(25) Then L = L + 1
If B(34) < B(26) Then L = L + 1
If B(34) < B(27) Then L = L + 1
If B(34) < B(28) Then L = L + 1
If B(34) < B(29) Then L = L + 1
If B(34) < B(30) Then L = L + 1
If B(34) < B(31) Then L = L + 1
If B(34) < B(32) Then L = L + 1
If B(34) < B(33) Then L = L + 1
If B(34) <= B(35) Then L = L + 1
If B(34) <= B(36) Then L = L + 1
If B(34) <= B(37) Then L = L + 1
If B(34) <= B(38) Then L = L + 1
If B(34) <= B(39) Then L = L + 1
If B(34) <= B(40) Then L = L + 1
If B(34) <= B(41) Then L = L + 1
If B(34) <= B(42) Then L = L + 1
If B(34) <= B(43) Then L = L + 1
If B(34) <= B(44) Then L = L + 1
If B(34) <= B(45) Then L = L + 1
If B(34) <= B(46) Then L = L + 1
If B(34) <= B(47) Then L = L + 1
If B(34) <= B(48) Then L = L + 1
If B(34) <= B(49) Then L = L + 1
 orderdisp
M = 35: L = 1
If B(35) < B(1) Then L = L + 1
If B(35) < B(2) Then L = L + 1
If B(35) < B(3) Then L = L + 1
If B(35) < B(4) Then L = L + 1
If B(35) < B(5) Then L = L + 1
If B(35) < B(6) Then L = L + 1
If B(35) < B(7) Then L = L + 1
If B(35) < B(8) Then L = L + 1
If B(35) < B(9) Then L = L + 1
If B(35) < B(10) Then L = L + 1
If B(35) < B(11) Then L = L + 1
If B(35) < B(12) Then L = L + 1
If B(35) < B(13) Then L = L + 1
If B(35) < B(14) Then L = L + 1
If B(35) < B(15) Then L = L + 1
If B(35) < B(16) Then L = L + 1
If B(35) < B(17) Then L = L + 1
If B(35) < B(18) Then L = L + 1
If B(35) < B(19) Then L = L + 1
If B(35) < B(20) Then L = L + 1
If B(35) < B(21) Then L = L + 1
If B(35) < B(22) Then L = L + 1
If B(35) < B(23) Then L = L + 1
If B(35) < B(24) Then L = L + 1
If B(35) < B(25) Then L = L + 1
If B(35) < B(26) Then L = L + 1
If B(35) < B(27) Then L = L + 1
If B(35) < B(28) Then L = L + 1
If B(35) < B(29) Then L = L + 1
If B(35) < B(30) Then L = L + 1
If B(35) < B(31) Then L = L + 1
If B(35) < B(32) Then L = L + 1
If B(35) < B(33) Then L = L + 1
If B(35) < B(34) Then L = L + 1
If B(35) <= B(36) Then L = L + 1
If B(35) <= B(37) Then L = L + 1
If B(35) <= B(38) Then L = L + 1
If B(35) <= B(39) Then L = L + 1
If B(35) <= B(40) Then L = L + 1
If B(35) <= B(41) Then L = L + 1
If B(35) <= B(42) Then L = L + 1
If B(35) <= B(43) Then L = L + 1
If B(35) <= B(44) Then L = L + 1
If B(35) <= B(45) Then L = L + 1
If B(35) <= B(46) Then L = L + 1
If B(35) <= B(47) Then L = L + 1
If B(35) <= B(48) Then L = L + 1
If B(35) <= B(49) Then L = L + 1
 orderdisp
M = 36: L = 1
If B(36) < B(1) Then L = L + 1
If B(36) < B(2) Then L = L + 1
If B(36) < B(3) Then L = L + 1
If B(36) < B(4) Then L = L + 1
If B(36) < B(5) Then L = L + 1
If B(36) < B(6) Then L = L + 1
If B(36) < B(7) Then L = L + 1
If B(36) < B(8) Then L = L + 1
If B(36) < B(9) Then L = L + 1
If B(36) < B(10) Then L = L + 1
If B(36) < B(11) Then L = L + 1
If B(36) < B(12) Then L = L + 1
If B(36) < B(13) Then L = L + 1
If B(36) < B(14) Then L = L + 1
If B(36) < B(15) Then L = L + 1
If B(36) < B(16) Then L = L + 1
If B(36) < B(17) Then L = L + 1
If B(36) < B(18) Then L = L + 1
If B(36) < B(19) Then L = L + 1
If B(36) < B(20) Then L = L + 1
If B(36) < B(21) Then L = L + 1
If B(36) < B(22) Then L = L + 1
If B(36) < B(23) Then L = L + 1
If B(36) < B(24) Then L = L + 1
If B(36) < B(25) Then L = L + 1
If B(36) < B(26) Then L = L + 1
If B(36) < B(27) Then L = L + 1
If B(36) < B(28) Then L = L + 1
If B(36) < B(29) Then L = L + 1
If B(36) < B(30) Then L = L + 1
If B(36) < B(31) Then L = L + 1
If B(36) < B(32) Then L = L + 1
If B(36) < B(33) Then L = L + 1
If B(36) < B(34) Then L = L + 1
If B(36) < B(35) Then L = L + 1
If B(36) <= B(37) Then L = L + 1
If B(36) <= B(38) Then L = L + 1
If B(36) <= B(39) Then L = L + 1
If B(36) <= B(40) Then L = L + 1
If B(36) <= B(41) Then L = L + 1
If B(36) <= B(42) Then L = L + 1
If B(36) <= B(43) Then L = L + 1
If B(36) <= B(44) Then L = L + 1
If B(36) <= B(45) Then L = L + 1
If B(36) <= B(46) Then L = L + 1
If B(36) <= B(47) Then L = L + 1
If B(36) <= B(48) Then L = L + 1
If B(36) <= B(49) Then L = L + 1
 orderdisp
M = 37: L = 1
If B(37) < B(1) Then L = L + 1
If B(37) < B(2) Then L = L + 1
If B(37) < B(3) Then L = L + 1
If B(37) < B(4) Then L = L + 1
If B(37) < B(5) Then L = L + 1
If B(37) < B(6) Then L = L + 1
If B(37) < B(7) Then L = L + 1
If B(37) < B(8) Then L = L + 1
If B(37) < B(9) Then L = L + 1
If B(37) < B(10) Then L = L + 1
If B(37) < B(11) Then L = L + 1
If B(37) < B(12) Then L = L + 1
If B(37) < B(13) Then L = L + 1
If B(37) < B(14) Then L = L + 1
If B(37) < B(15) Then L = L + 1
If B(37) < B(16) Then L = L + 1
If B(37) < B(17) Then L = L + 1
If B(37) < B(18) Then L = L + 1
If B(37) < B(19) Then L = L + 1
If B(37) < B(20) Then L = L + 1
If B(37) < B(21) Then L = L + 1
If B(37) < B(22) Then L = L + 1
If B(37) < B(23) Then L = L + 1
If B(37) < B(24) Then L = L + 1
If B(37) < B(25) Then L = L + 1
If B(37) < B(26) Then L = L + 1
If B(37) < B(27) Then L = L + 1
If B(37) < B(28) Then L = L + 1
If B(37) < B(29) Then L = L + 1
If B(37) < B(30) Then L = L + 1
If B(37) < B(31) Then L = L + 1
If B(37) < B(32) Then L = L + 1
If B(37) < B(33) Then L = L + 1
If B(37) < B(34) Then L = L + 1
If B(37) < B(35) Then L = L + 1
If B(37) < B(36) Then L = L + 1
If B(37) <= B(38) Then L = L + 1
If B(37) <= B(39) Then L = L + 1
If B(37) <= B(40) Then L = L + 1
If B(37) <= B(41) Then L = L + 1
If B(37) <= B(42) Then L = L + 1
If B(37) <= B(43) Then L = L + 1
If B(37) <= B(44) Then L = L + 1
If B(37) <= B(45) Then L = L + 1
If B(37) <= B(46) Then L = L + 1
If B(37) <= B(47) Then L = L + 1
If B(37) <= B(48) Then L = L + 1
If B(37) <= B(49) Then L = L + 1
 orderdisp
M = 38: L = 1
If B(38) < B(1) Then L = L + 1
If B(38) < B(2) Then L = L + 1
If B(38) < B(3) Then L = L + 1
If B(38) < B(4) Then L = L + 1
If B(38) < B(5) Then L = L + 1
If B(38) < B(6) Then L = L + 1
If B(38) < B(7) Then L = L + 1
If B(38) < B(8) Then L = L + 1
If B(38) < B(9) Then L = L + 1
If B(38) < B(10) Then L = L + 1
If B(38) < B(11) Then L = L + 1
If B(38) < B(12) Then L = L + 1
If B(38) < B(13) Then L = L + 1
If B(38) < B(14) Then L = L + 1
If B(38) < B(15) Then L = L + 1
If B(38) < B(16) Then L = L + 1
If B(38) < B(17) Then L = L + 1
If B(38) < B(18) Then L = L + 1
If B(38) < B(19) Then L = L + 1
If B(38) < B(20) Then L = L + 1
If B(38) < B(21) Then L = L + 1
If B(38) < B(22) Then L = L + 1
If B(38) < B(23) Then L = L + 1
If B(38) < B(24) Then L = L + 1
If B(38) < B(25) Then L = L + 1
If B(38) < B(26) Then L = L + 1
If B(38) < B(27) Then L = L + 1
If B(38) < B(28) Then L = L + 1
If B(38) < B(29) Then L = L + 1
If B(38) < B(30) Then L = L + 1
If B(38) < B(31) Then L = L + 1
If B(38) < B(32) Then L = L + 1
If B(38) < B(33) Then L = L + 1
If B(38) < B(34) Then L = L + 1
If B(38) < B(35) Then L = L + 1
If B(38) < B(36) Then L = L + 1
If B(38) < B(37) Then L = L + 1
If B(38) <= B(39) Then L = L + 1
If B(38) <= B(40) Then L = L + 1
If B(38) <= B(41) Then L = L + 1
If B(38) <= B(42) Then L = L + 1
If B(38) <= B(43) Then L = L + 1
If B(38) <= B(44) Then L = L + 1
If B(38) <= B(45) Then L = L + 1
If B(38) <= B(46) Then L = L + 1
If B(38) <= B(47) Then L = L + 1
If B(38) <= B(48) Then L = L + 1
If B(38) <= B(49) Then L = L + 1
 orderdisp
M = 39: L = 1
If B(39) < B(1) Then L = L + 1
If B(39) < B(2) Then L = L + 1
If B(39) < B(3) Then L = L + 1
If B(39) < B(4) Then L = L + 1
If B(39) < B(5) Then L = L + 1
If B(39) < B(6) Then L = L + 1
If B(39) < B(7) Then L = L + 1
If B(39) < B(8) Then L = L + 1
If B(39) < B(9) Then L = L + 1
If B(39) < B(10) Then L = L + 1
If B(39) < B(11) Then L = L + 1
If B(39) < B(12) Then L = L + 1
If B(39) < B(13) Then L = L + 1
If B(39) < B(14) Then L = L + 1
If B(39) < B(15) Then L = L + 1
If B(39) < B(16) Then L = L + 1
If B(39) < B(17) Then L = L + 1
If B(39) < B(18) Then L = L + 1
If B(39) < B(19) Then L = L + 1
If B(39) < B(20) Then L = L + 1
If B(39) < B(21) Then L = L + 1
If B(39) < B(22) Then L = L + 1
If B(39) < B(23) Then L = L + 1
If B(39) < B(24) Then L = L + 1
If B(39) < B(25) Then L = L + 1
If B(39) < B(26) Then L = L + 1
If B(39) < B(27) Then L = L + 1
If B(39) < B(28) Then L = L + 1
If B(39) < B(29) Then L = L + 1
If B(39) < B(30) Then L = L + 1
If B(39) < B(31) Then L = L + 1
If B(39) < B(32) Then L = L + 1
If B(39) < B(33) Then L = L + 1
If B(39) < B(34) Then L = L + 1
If B(39) < B(35) Then L = L + 1
If B(39) < B(36) Then L = L + 1
If B(39) < B(37) Then L = L + 1
If B(39) < B(38) Then L = L + 1
If B(39) <= B(40) Then L = L + 1
If B(39) <= B(41) Then L = L + 1
If B(39) <= B(42) Then L = L + 1
If B(39) <= B(43) Then L = L + 1
If B(39) <= B(44) Then L = L + 1
If B(39) <= B(45) Then L = L + 1
If B(39) <= B(46) Then L = L + 1
If B(39) <= B(47) Then L = L + 1
If B(39) <= B(48) Then L = L + 1
If B(39) <= B(49) Then L = L + 1
 orderdisp
M = 40: L = 1
If B(40) < B(1) Then L = L + 1
If B(40) < B(2) Then L = L + 1
If B(40) < B(3) Then L = L + 1
If B(40) < B(4) Then L = L + 1
If B(40) < B(5) Then L = L + 1
If B(40) < B(6) Then L = L + 1
If B(40) < B(7) Then L = L + 1
If B(40) < B(8) Then L = L + 1
If B(40) < B(9) Then L = L + 1
If B(40) < B(10) Then L = L + 1
If B(40) < B(11) Then L = L + 1
If B(40) < B(12) Then L = L + 1
If B(40) < B(13) Then L = L + 1
If B(40) < B(14) Then L = L + 1
If B(40) < B(15) Then L = L + 1
If B(40) < B(16) Then L = L + 1
If B(40) < B(17) Then L = L + 1
If B(40) < B(18) Then L = L + 1
If B(40) < B(19) Then L = L + 1
If B(40) < B(20) Then L = L + 1
If B(40) < B(21) Then L = L + 1
If B(40) < B(22) Then L = L + 1
If B(40) < B(23) Then L = L + 1
If B(40) < B(24) Then L = L + 1
If B(40) < B(25) Then L = L + 1
If B(40) < B(26) Then L = L + 1
If B(40) < B(27) Then L = L + 1
If B(40) < B(28) Then L = L + 1
If B(40) < B(29) Then L = L + 1
If B(40) < B(30) Then L = L + 1
If B(40) < B(31) Then L = L + 1
If B(40) < B(32) Then L = L + 1
If B(40) < B(33) Then L = L + 1
If B(40) < B(34) Then L = L + 1
If B(40) < B(35) Then L = L + 1
If B(40) < B(36) Then L = L + 1
If B(40) < B(37) Then L = L + 1
If B(40) < B(38) Then L = L + 1
If B(40) < B(39) Then L = L + 1
If B(40) <= B(41) Then L = L + 1
If B(40) <= B(42) Then L = L + 1
If B(40) <= B(43) Then L = L + 1
If B(40) <= B(44) Then L = L + 1
If B(40) <= B(45) Then L = L + 1
If B(40) <= B(46) Then L = L + 1
If B(40) <= B(47) Then L = L + 1
If B(40) <= B(48) Then L = L + 1
If B(40) <= B(49) Then L = L + 1
orderdisp
Order3
End Function

Public Function NumDisplay2()
If LD = 1 Then
If LB(1) = 1 Then txtBall(1).BackColor = QBColor(R)
If LB(1) = 2 Then txtBall(2).BackColor = QBColor(R)
If LB(1) = 3 Then txtBall(3).BackColor = QBColor(R)
If LB(1) = 4 Then txtBall(4).BackColor = QBColor(R)
If LB(1) = 5 Then txtBall(5).BackColor = QBColor(R)
If LB(1) = 6 Then txtBall(6).BackColor = QBColor(R)
If LB(1) = 7 Then txtBall(7).BackColor = QBColor(R)
If LB(1) = 8 Then txtBall(8).BackColor = QBColor(R)
If LB(1) = 9 Then txtBall(9).BackColor = QBColor(R)
If LB(1) = 10 Then txtBall(10).BackColor = QBColor(R)
If LB(1) = 11 Then txtBall(11).BackColor = QBColor(R)
If LB(1) = 12 Then txtBall(12).BackColor = QBColor(R)
If LB(1) = 13 Then txtBall(13).BackColor = QBColor(R)
If LB(1) = 14 Then txtBall(14).BackColor = QBColor(R)
If LB(1) = 15 Then txtBall(15).BackColor = QBColor(R)
If LB(1) = 16 Then txtBall(16).BackColor = QBColor(R)
If LB(1) = 17 Then txtBall(17).BackColor = QBColor(R)
If LB(1) = 18 Then txtBall(18).BackColor = QBColor(R)
If LB(1) = 19 Then txtBall(19).BackColor = QBColor(R)
If LB(1) = 20 Then txtBall(20).BackColor = QBColor(R)
If LB(1) = 21 Then txtBall(21).BackColor = QBColor(R)
If LB(1) = 22 Then txtBall(22).BackColor = QBColor(R)
If LB(1) = 23 Then txtBall(23).BackColor = QBColor(R)
If LB(1) = 24 Then txtBall(24).BackColor = QBColor(R)
If LB(1) = 25 Then txtBall(25).BackColor = QBColor(R)
If LB(1) = 26 Then txtBall(26).BackColor = QBColor(R)
If LB(1) = 27 Then txtBall(27).BackColor = QBColor(R)
If LB(1) = 28 Then txtBall(28).BackColor = QBColor(R)
If LB(1) = 29 Then txtBall(29).BackColor = QBColor(R)
If LB(1) = 30 Then txtBall(30).BackColor = QBColor(R)
If LB(1) = 31 Then txtBall(31).BackColor = QBColor(R)
If LB(1) = 32 Then txtBall(32).BackColor = QBColor(R)
If LB(1) = 33 Then txtBall(33).BackColor = QBColor(R)
If LB(1) = 34 Then txtBall(34).BackColor = QBColor(R)
If LB(1) = 35 Then txtBall(35).BackColor = QBColor(R)
If LB(1) = 36 Then txtBall(36).BackColor = QBColor(R)
If LB(1) = 37 Then txtBall(37).BackColor = QBColor(R)
If LB(1) = 38 Then txtBall(38).BackColor = QBColor(R)
If LB(1) = 39 Then txtBall(39).BackColor = QBColor(R)
If LB(1) = 40 Then txtBall(40).BackColor = QBColor(R)
If LB(1) = 41 Then txtBall(41).BackColor = QBColor(R)
If LB(1) = 42 Then txtBall(42).BackColor = QBColor(R)
If LB(1) = 43 Then txtBall(43).BackColor = QBColor(R)
If LB(1) = 44 Then txtBall(44).BackColor = QBColor(R)
If LB(1) = 45 Then txtBall(45).BackColor = QBColor(R)
If LB(1) = 46 Then txtBall(46).BackColor = QBColor(R)
If LB(1) = 47 Then txtBall(47).BackColor = QBColor(R)
If LB(1) = 48 Then txtBall(48).BackColor = QBColor(R)
If LB(1) = 49 Then txtBall(49).BackColor = QBColor(R)
If LB(2) = 1 Then txtBall(1).BackColor = QBColor(R)
If LB(2) = 2 Then txtBall(2).BackColor = QBColor(R)
If LB(2) = 3 Then txtBall(3).BackColor = QBColor(R)
If LB(2) = 4 Then txtBall(4).BackColor = QBColor(R)
If LB(2) = 5 Then txtBall(5).BackColor = QBColor(R)
If LB(2) = 6 Then txtBall(6).BackColor = QBColor(R)
If LB(2) = 7 Then txtBall(7).BackColor = QBColor(R)
If LB(2) = 8 Then txtBall(8).BackColor = QBColor(R)
If LB(2) = 9 Then txtBall(9).BackColor = QBColor(R)
If LB(2) = 10 Then txtBall(10).BackColor = QBColor(R)
If LB(2) = 11 Then txtBall(11).BackColor = QBColor(R)
If LB(2) = 12 Then txtBall(12).BackColor = QBColor(R)
If LB(2) = 13 Then txtBall(13).BackColor = QBColor(R)
If LB(2) = 14 Then txtBall(14).BackColor = QBColor(R)
If LB(2) = 15 Then txtBall(15).BackColor = QBColor(R)
If LB(2) = 16 Then txtBall(16).BackColor = QBColor(R)
If LB(2) = 17 Then txtBall(17).BackColor = QBColor(R)
If LB(2) = 18 Then txtBall(18).BackColor = QBColor(R)
If LB(2) = 19 Then txtBall(19).BackColor = QBColor(R)
If LB(2) = 20 Then txtBall(20).BackColor = QBColor(R)
If LB(2) = 21 Then txtBall(21).BackColor = QBColor(R)
If LB(2) = 22 Then txtBall(22).BackColor = QBColor(R)
If LB(2) = 23 Then txtBall(23).BackColor = QBColor(R)
If LB(2) = 24 Then txtBall(24).BackColor = QBColor(R)
If LB(2) = 25 Then txtBall(25).BackColor = QBColor(R)
If LB(2) = 26 Then txtBall(26).BackColor = QBColor(R)
If LB(2) = 27 Then txtBall(27).BackColor = QBColor(R)
If LB(2) = 28 Then txtBall(28).BackColor = QBColor(R)
If LB(2) = 29 Then txtBall(29).BackColor = QBColor(R)
If LB(2) = 30 Then txtBall(30).BackColor = QBColor(R)
If LB(2) = 31 Then txtBall(31).BackColor = QBColor(R)
If LB(2) = 32 Then txtBall(32).BackColor = QBColor(R)
If LB(2) = 33 Then txtBall(33).BackColor = QBColor(R)
If LB(2) = 34 Then txtBall(34).BackColor = QBColor(R)
If LB(2) = 35 Then txtBall(35).BackColor = QBColor(R)
If LB(2) = 36 Then txtBall(36).BackColor = QBColor(R)
If LB(2) = 37 Then txtBall(37).BackColor = QBColor(R)
If LB(2) = 38 Then txtBall(38).BackColor = QBColor(R)
If LB(2) = 39 Then txtBall(39).BackColor = QBColor(R)
If LB(2) = 40 Then txtBall(40).BackColor = QBColor(R)
If LB(2) = 41 Then txtBall(41).BackColor = QBColor(R)
If LB(2) = 42 Then txtBall(42).BackColor = QBColor(R)
If LB(2) = 43 Then txtBall(43).BackColor = QBColor(R)
If LB(2) = 44 Then txtBall(44).BackColor = QBColor(R)
If LB(2) = 45 Then txtBall(45).BackColor = QBColor(R)
If LB(2) = 46 Then txtBall(46).BackColor = QBColor(R)
If LB(2) = 47 Then txtBall(47).BackColor = QBColor(R)
If LB(2) = 48 Then txtBall(48).BackColor = QBColor(R)
If LB(2) = 49 Then txtBall(49).BackColor = QBColor(R)
If LB(3) = 1 Then txtBall(1).BackColor = QBColor(R)
If LB(3) = 2 Then txtBall(2).BackColor = QBColor(R)
If LB(3) = 3 Then txtBall(3).BackColor = QBColor(R)
If LB(3) = 4 Then txtBall(4).BackColor = QBColor(R)
If LB(3) = 5 Then txtBall(5).BackColor = QBColor(R)
If LB(3) = 6 Then txtBall(6).BackColor = QBColor(R)
If LB(3) = 7 Then txtBall(7).BackColor = QBColor(R)
If LB(3) = 8 Then txtBall(8).BackColor = QBColor(R)
If LB(3) = 9 Then txtBall(9).BackColor = QBColor(R)
If LB(3) = 10 Then txtBall(10).BackColor = QBColor(R)
If LB(3) = 11 Then txtBall(11).BackColor = QBColor(R)
If LB(3) = 12 Then txtBall(12).BackColor = QBColor(R)
If LB(3) = 13 Then txtBall(13).BackColor = QBColor(R)
If LB(3) = 14 Then txtBall(14).BackColor = QBColor(R)
If LB(3) = 15 Then txtBall(15).BackColor = QBColor(R)
If LB(3) = 16 Then txtBall(16).BackColor = QBColor(R)
If LB(3) = 17 Then txtBall(17).BackColor = QBColor(R)
If LB(3) = 18 Then txtBall(18).BackColor = QBColor(R)
If LB(3) = 19 Then txtBall(19).BackColor = QBColor(R)
If LB(3) = 20 Then txtBall(20).BackColor = QBColor(R)
If LB(3) = 21 Then txtBall(21).BackColor = QBColor(R)
If LB(3) = 22 Then txtBall(22).BackColor = QBColor(R)
If LB(3) = 23 Then txtBall(23).BackColor = QBColor(R)
If LB(3) = 24 Then txtBall(24).BackColor = QBColor(R)
If LB(3) = 25 Then txtBall(25).BackColor = QBColor(R)
If LB(3) = 26 Then txtBall(26).BackColor = QBColor(R)
If LB(3) = 27 Then txtBall(27).BackColor = QBColor(R)
If LB(3) = 28 Then txtBall(28).BackColor = QBColor(R)
If LB(3) = 29 Then txtBall(29).BackColor = QBColor(R)
If LB(3) = 30 Then txtBall(30).BackColor = QBColor(R)
If LB(3) = 31 Then txtBall(31).BackColor = QBColor(R)
If LB(3) = 32 Then txtBall(32).BackColor = QBColor(R)
If LB(3) = 33 Then txtBall(33).BackColor = QBColor(R)
If LB(3) = 34 Then txtBall(34).BackColor = QBColor(R)
If LB(3) = 35 Then txtBall(35).BackColor = QBColor(R)
If LB(3) = 36 Then txtBall(36).BackColor = QBColor(R)
If LB(3) = 37 Then txtBall(37).BackColor = QBColor(R)
If LB(3) = 38 Then txtBall(38).BackColor = QBColor(R)
If LB(3) = 39 Then txtBall(39).BackColor = QBColor(R)
If LB(3) = 40 Then txtBall(40).BackColor = QBColor(R)
If LB(3) = 41 Then txtBall(41).BackColor = QBColor(R)
If LB(3) = 42 Then txtBall(42).BackColor = QBColor(R)
If LB(3) = 43 Then txtBall(43).BackColor = QBColor(R)
If LB(3) = 44 Then txtBall(44).BackColor = QBColor(R)
If LB(3) = 45 Then txtBall(45).BackColor = QBColor(R)
If LB(3) = 46 Then txtBall(46).BackColor = QBColor(R)
If LB(3) = 47 Then txtBall(47).BackColor = QBColor(R)
If LB(3) = 48 Then txtBall(48).BackColor = QBColor(R)
If LB(3) = 49 Then txtBall(49).BackColor = QBColor(R)
If LB(4) = 1 Then txtBall(1).BackColor = QBColor(R)
If LB(4) = 2 Then txtBall(2).BackColor = QBColor(R)
If LB(4) = 3 Then txtBall(3).BackColor = QBColor(R)
If LB(4) = 4 Then txtBall(4).BackColor = QBColor(R)
If LB(4) = 5 Then txtBall(5).BackColor = QBColor(R)
If LB(4) = 6 Then txtBall(6).BackColor = QBColor(R)
If LB(4) = 7 Then txtBall(7).BackColor = QBColor(R)
If LB(4) = 8 Then txtBall(8).BackColor = QBColor(R)
If LB(4) = 9 Then txtBall(9).BackColor = QBColor(R)
If LB(4) = 10 Then txtBall(10).BackColor = QBColor(R)
If LB(4) = 11 Then txtBall(11).BackColor = QBColor(R)
If LB(4) = 12 Then txtBall(12).BackColor = QBColor(R)
If LB(4) = 13 Then txtBall(13).BackColor = QBColor(R)
If LB(4) = 14 Then txtBall(14).BackColor = QBColor(R)
If LB(4) = 15 Then txtBall(15).BackColor = QBColor(R)
If LB(4) = 16 Then txtBall(16).BackColor = QBColor(R)
If LB(4) = 17 Then txtBall(17).BackColor = QBColor(R)
If LB(4) = 18 Then txtBall(18).BackColor = QBColor(R)
If LB(4) = 19 Then txtBall(19).BackColor = QBColor(R)
If LB(4) = 20 Then txtBall(20).BackColor = QBColor(R)
If LB(4) = 21 Then txtBall(21).BackColor = QBColor(R)
If LB(4) = 22 Then txtBall(22).BackColor = QBColor(R)
If LB(4) = 23 Then txtBall(23).BackColor = QBColor(R)
If LB(4) = 24 Then txtBall(24).BackColor = QBColor(R)
If LB(4) = 25 Then txtBall(25).BackColor = QBColor(R)
If LB(4) = 26 Then txtBall(26).BackColor = QBColor(R)
If LB(4) = 27 Then txtBall(27).BackColor = QBColor(R)
If LB(4) = 28 Then txtBall(28).BackColor = QBColor(R)
If LB(4) = 29 Then txtBall(29).BackColor = QBColor(R)
If LB(4) = 30 Then txtBall(30).BackColor = QBColor(R)
If LB(4) = 31 Then txtBall(31).BackColor = QBColor(R)
If LB(4) = 32 Then txtBall(32).BackColor = QBColor(R)
If LB(4) = 33 Then txtBall(33).BackColor = QBColor(R)
If LB(4) = 34 Then txtBall(34).BackColor = QBColor(R)
If LB(4) = 35 Then txtBall(35).BackColor = QBColor(R)
If LB(4) = 36 Then txtBall(36).BackColor = QBColor(R)
If LB(4) = 37 Then txtBall(37).BackColor = QBColor(R)
If LB(4) = 38 Then txtBall(38).BackColor = QBColor(R)
If LB(4) = 39 Then txtBall(39).BackColor = QBColor(R)
If LB(4) = 40 Then txtBall(40).BackColor = QBColor(R)
If LB(4) = 41 Then txtBall(41).BackColor = QBColor(R)
If LB(4) = 42 Then txtBall(42).BackColor = QBColor(R)
If LB(4) = 43 Then txtBall(43).BackColor = QBColor(R)
If LB(4) = 44 Then txtBall(44).BackColor = QBColor(R)
If LB(4) = 45 Then txtBall(45).BackColor = QBColor(R)
If LB(4) = 46 Then txtBall(46).BackColor = QBColor(R)
If LB(4) = 47 Then txtBall(47).BackColor = QBColor(R)
If LB(4) = 48 Then txtBall(48).BackColor = QBColor(R)
If LB(4) = 49 Then txtBall(49).BackColor = QBColor(R)
If LB(5) = 1 Then txtBall(1).BackColor = QBColor(R)
If LB(5) = 2 Then txtBall(2).BackColor = QBColor(R)
If LB(5) = 3 Then txtBall(3).BackColor = QBColor(R)
If LB(5) = 4 Then txtBall(4).BackColor = QBColor(R)
If LB(5) = 5 Then txtBall(5).BackColor = QBColor(R)
If LB(5) = 6 Then txtBall(6).BackColor = QBColor(R)
If LB(5) = 7 Then txtBall(7).BackColor = QBColor(R)
If LB(5) = 8 Then txtBall(8).BackColor = QBColor(R)
If LB(5) = 9 Then txtBall(9).BackColor = QBColor(R)
If LB(5) = 10 Then txtBall(10).BackColor = QBColor(R)
If LB(5) = 11 Then txtBall(11).BackColor = QBColor(R)
If LB(5) = 12 Then txtBall(12).BackColor = QBColor(R)
If LB(5) = 13 Then txtBall(13).BackColor = QBColor(R)
If LB(5) = 14 Then txtBall(14).BackColor = QBColor(R)
If LB(5) = 15 Then txtBall(15).BackColor = QBColor(R)
If LB(5) = 16 Then txtBall(16).BackColor = QBColor(R)
If LB(5) = 17 Then txtBall(17).BackColor = QBColor(R)
If LB(5) = 18 Then txtBall(18).BackColor = QBColor(R)
If LB(5) = 19 Then txtBall(19).BackColor = QBColor(R)
If LB(5) = 20 Then txtBall(20).BackColor = QBColor(R)
If LB(5) = 21 Then txtBall(21).BackColor = QBColor(R)
If LB(5) = 22 Then txtBall(22).BackColor = QBColor(R)
If LB(5) = 23 Then txtBall(23).BackColor = QBColor(R)
If LB(5) = 24 Then txtBall(24).BackColor = QBColor(R)
If LB(5) = 25 Then txtBall(25).BackColor = QBColor(R)
If LB(5) = 26 Then txtBall(26).BackColor = QBColor(R)
If LB(5) = 27 Then txtBall(27).BackColor = QBColor(R)
If LB(5) = 28 Then txtBall(28).BackColor = QBColor(R)
If LB(5) = 29 Then txtBall(29).BackColor = QBColor(R)
If LB(5) = 30 Then txtBall(30).BackColor = QBColor(R)
If LB(5) = 31 Then txtBall(31).BackColor = QBColor(R)
If LB(5) = 32 Then txtBall(32).BackColor = QBColor(R)
If LB(5) = 33 Then txtBall(33).BackColor = QBColor(R)
If LB(5) = 34 Then txtBall(34).BackColor = QBColor(R)
If LB(5) = 35 Then txtBall(35).BackColor = QBColor(R)
If LB(5) = 36 Then txtBall(36).BackColor = QBColor(R)
If LB(5) = 37 Then txtBall(37).BackColor = QBColor(R)
If LB(5) = 38 Then txtBall(38).BackColor = QBColor(R)
If LB(5) = 39 Then txtBall(39).BackColor = QBColor(R)
If LB(5) = 40 Then txtBall(40).BackColor = QBColor(R)
If LB(5) = 41 Then txtBall(41).BackColor = QBColor(R)
If LB(5) = 42 Then txtBall(42).BackColor = QBColor(R)
If LB(5) = 43 Then txtBall(43).BackColor = QBColor(R)
If LB(5) = 44 Then txtBall(44).BackColor = QBColor(R)
If LB(5) = 45 Then txtBall(45).BackColor = QBColor(R)
If LB(5) = 46 Then txtBall(46).BackColor = QBColor(R)
If LB(5) = 47 Then txtBall(47).BackColor = QBColor(R)
If LB(5) = 48 Then txtBall(48).BackColor = QBColor(R)
If LB(5) = 49 Then txtBall(49).BackColor = QBColor(R)
If LB(6) = 1 Then txtBall(1).BackColor = QBColor(R)
If LB(6) = 2 Then txtBall(2).BackColor = QBColor(R)
If LB(6) = 3 Then txtBall(3).BackColor = QBColor(R)
If LB(6) = 4 Then txtBall(4).BackColor = QBColor(R)
If LB(6) = 5 Then txtBall(5).BackColor = QBColor(R)
If LB(6) = 6 Then txtBall(6).BackColor = QBColor(R)
If LB(6) = 7 Then txtBall(7).BackColor = QBColor(R)
If LB(6) = 8 Then txtBall(8).BackColor = QBColor(R)
If LB(6) = 9 Then txtBall(9).BackColor = QBColor(R)
If LB(6) = 10 Then txtBall(10).BackColor = QBColor(R)
If LB(6) = 11 Then txtBall(11).BackColor = QBColor(R)
If LB(6) = 12 Then txtBall(12).BackColor = QBColor(R)
If LB(6) = 13 Then txtBall(13).BackColor = QBColor(R)
If LB(6) = 14 Then txtBall(14).BackColor = QBColor(R)
If LB(6) = 15 Then txtBall(15).BackColor = QBColor(R)
If LB(6) = 16 Then txtBall(16).BackColor = QBColor(R)
If LB(6) = 17 Then txtBall(17).BackColor = QBColor(R)
If LB(6) = 18 Then txtBall(18).BackColor = QBColor(R)
If LB(6) = 19 Then txtBall(19).BackColor = QBColor(R)
If LB(6) = 20 Then txtBall(20).BackColor = QBColor(R)
If LB(6) = 21 Then txtBall(21).BackColor = QBColor(R)
If LB(6) = 22 Then txtBall(22).BackColor = QBColor(R)
If LB(6) = 23 Then txtBall(23).BackColor = QBColor(R)
If LB(6) = 24 Then txtBall(24).BackColor = QBColor(R)
If LB(6) = 25 Then txtBall(25).BackColor = QBColor(R)
If LB(6) = 26 Then txtBall(26).BackColor = QBColor(R)
If LB(6) = 27 Then txtBall(27).BackColor = QBColor(R)
If LB(6) = 28 Then txtBall(28).BackColor = QBColor(R)
If LB(6) = 29 Then txtBall(29).BackColor = QBColor(R)
If LB(6) = 30 Then txtBall(30).BackColor = QBColor(R)
If LB(6) = 31 Then txtBall(31).BackColor = QBColor(R)
If LB(6) = 32 Then txtBall(32).BackColor = QBColor(R)
If LB(6) = 33 Then txtBall(33).BackColor = QBColor(R)
If LB(6) = 34 Then txtBall(34).BackColor = QBColor(R)
If LB(6) = 35 Then txtBall(35).BackColor = QBColor(R)
If LB(6) = 36 Then txtBall(36).BackColor = QBColor(R)
If LB(6) = 37 Then txtBall(37).BackColor = QBColor(R)
If LB(6) = 38 Then txtBall(38).BackColor = QBColor(R)
If LB(6) = 39 Then txtBall(39).BackColor = QBColor(R)
If LB(6) = 40 Then txtBall(40).BackColor = QBColor(R)
If LB(6) = 41 Then txtBall(41).BackColor = QBColor(R)
If LB(6) = 42 Then txtBall(42).BackColor = QBColor(R)
If LB(6) = 43 Then txtBall(43).BackColor = QBColor(R)
If LB(6) = 44 Then txtBall(44).BackColor = QBColor(R)
If LB(6) = 45 Then txtBall(45).BackColor = QBColor(R)
If LB(6) = 46 Then txtBall(46).BackColor = QBColor(R)
If LB(6) = 47 Then txtBall(47).BackColor = QBColor(R)
If LB(6) = 48 Then txtBall(48).BackColor = QBColor(R)
If LB(6) = 49 Then txtBall(49).BackColor = QBColor(R)
End If
If PD = 1 Then
If PB(1) = 1 Then txtBall(1).BackColor = QBColor(C)
If PB(1) = 2 Then txtBall(2).BackColor = QBColor(C)
If PB(1) = 3 Then txtBall(3).BackColor = QBColor(C)
If PB(1) = 4 Then txtBall(4).BackColor = QBColor(C)
If PB(1) = 5 Then txtBall(5).BackColor = QBColor(C)
If PB(1) = 6 Then txtBall(6).BackColor = QBColor(C)
If PB(1) = 7 Then txtBall(7).BackColor = QBColor(C)
If PB(1) = 8 Then txtBall(8).BackColor = QBColor(C)
If PB(1) = 9 Then txtBall(9).BackColor = QBColor(C)
If PB(1) = 10 Then txtBall(10).BackColor = QBColor(C)
If PB(1) = 11 Then txtBall(11).BackColor = QBColor(C)
If PB(1) = 12 Then txtBall(12).BackColor = QBColor(C)
If PB(1) = 13 Then txtBall(13).BackColor = QBColor(C)
If PB(1) = 14 Then txtBall(14).BackColor = QBColor(C)
If PB(1) = 15 Then txtBall(15).BackColor = QBColor(C)
If PB(1) = 16 Then txtBall(16).BackColor = QBColor(C)
If PB(1) = 17 Then txtBall(17).BackColor = QBColor(C)
If PB(1) = 18 Then txtBall(18).BackColor = QBColor(C)
If PB(1) = 19 Then txtBall(19).BackColor = QBColor(C)
If PB(1) = 20 Then txtBall(20).BackColor = QBColor(C)
If PB(1) = 21 Then txtBall(21).BackColor = QBColor(C)
If PB(1) = 22 Then txtBall(22).BackColor = QBColor(C)
If PB(1) = 23 Then txtBall(23).BackColor = QBColor(C)
If PB(1) = 24 Then txtBall(24).BackColor = QBColor(C)
If PB(1) = 25 Then txtBall(25).BackColor = QBColor(C)
If PB(1) = 26 Then txtBall(26).BackColor = QBColor(C)
If PB(1) = 27 Then txtBall(27).BackColor = QBColor(C)
If PB(1) = 28 Then txtBall(28).BackColor = QBColor(C)
If PB(1) = 29 Then txtBall(29).BackColor = QBColor(C)
If PB(1) = 30 Then txtBall(30).BackColor = QBColor(C)
If PB(1) = 31 Then txtBall(31).BackColor = QBColor(C)
If PB(1) = 32 Then txtBall(32).BackColor = QBColor(C)
If PB(1) = 33 Then txtBall(33).BackColor = QBColor(C)
If PB(1) = 34 Then txtBall(34).BackColor = QBColor(C)
If PB(1) = 35 Then txtBall(35).BackColor = QBColor(C)
If PB(1) = 36 Then txtBall(36).BackColor = QBColor(C)
If PB(1) = 37 Then txtBall(37).BackColor = QBColor(C)
If PB(1) = 38 Then txtBall(38).BackColor = QBColor(C)
If PB(1) = 39 Then txtBall(39).BackColor = QBColor(C)
If PB(1) = 40 Then txtBall(40).BackColor = QBColor(C)
If PB(1) = 41 Then txtBall(41).BackColor = QBColor(C)
If PB(1) = 42 Then txtBall(42).BackColor = QBColor(C)
If PB(1) = 43 Then txtBall(43).BackColor = QBColor(C)
If PB(1) = 44 Then txtBall(44).BackColor = QBColor(C)
If PB(1) = 45 Then txtBall(45).BackColor = QBColor(C)
If PB(1) = 46 Then txtBall(46).BackColor = QBColor(C)
If PB(1) = 47 Then txtBall(47).BackColor = QBColor(C)
If PB(1) = 48 Then txtBall(48).BackColor = QBColor(C)
If PB(1) = 49 Then txtBall(49).BackColor = QBColor(C)
If PB(2) = 1 Then txtBall(1).BackColor = QBColor(C)
If PB(2) = 2 Then txtBall(2).BackColor = QBColor(C)
If PB(2) = 3 Then txtBall(3).BackColor = QBColor(C)
If PB(2) = 4 Then txtBall(4).BackColor = QBColor(C)
If PB(2) = 5 Then txtBall(5).BackColor = QBColor(C)
If PB(2) = 6 Then txtBall(6).BackColor = QBColor(C)
If PB(2) = 7 Then txtBall(7).BackColor = QBColor(C)
If PB(2) = 8 Then txtBall(8).BackColor = QBColor(C)
If PB(2) = 9 Then txtBall(9).BackColor = QBColor(C)
If PB(2) = 10 Then txtBall(10).BackColor = QBColor(C)
If PB(2) = 11 Then txtBall(11).BackColor = QBColor(C)
If PB(2) = 12 Then txtBall(12).BackColor = QBColor(C)
If PB(2) = 13 Then txtBall(13).BackColor = QBColor(C)
If PB(2) = 14 Then txtBall(14).BackColor = QBColor(C)
If PB(2) = 15 Then txtBall(15).BackColor = QBColor(C)
If PB(2) = 16 Then txtBall(16).BackColor = QBColor(C)
If PB(2) = 17 Then txtBall(17).BackColor = QBColor(C)
If PB(2) = 18 Then txtBall(18).BackColor = QBColor(C)
If PB(2) = 19 Then txtBall(19).BackColor = QBColor(C)
If PB(2) = 20 Then txtBall(20).BackColor = QBColor(C)
If PB(2) = 21 Then txtBall(21).BackColor = QBColor(C)
If PB(2) = 22 Then txtBall(22).BackColor = QBColor(C)
If PB(2) = 23 Then txtBall(23).BackColor = QBColor(C)
If PB(2) = 24 Then txtBall(24).BackColor = QBColor(C)
If PB(2) = 25 Then txtBall(25).BackColor = QBColor(C)
If PB(2) = 26 Then txtBall(26).BackColor = QBColor(C)
If PB(2) = 27 Then txtBall(27).BackColor = QBColor(C)
If PB(2) = 28 Then txtBall(28).BackColor = QBColor(C)
If PB(2) = 29 Then txtBall(29).BackColor = QBColor(C)
If PB(2) = 30 Then txtBall(30).BackColor = QBColor(C)
If PB(2) = 31 Then txtBall(31).BackColor = QBColor(C)
If PB(2) = 32 Then txtBall(32).BackColor = QBColor(C)
If PB(2) = 33 Then txtBall(33).BackColor = QBColor(C)
If PB(2) = 34 Then txtBall(34).BackColor = QBColor(C)
If PB(2) = 35 Then txtBall(35).BackColor = QBColor(C)
If PB(2) = 36 Then txtBall(36).BackColor = QBColor(C)
If PB(2) = 37 Then txtBall(37).BackColor = QBColor(C)
If PB(2) = 38 Then txtBall(38).BackColor = QBColor(C)
If PB(2) = 39 Then txtBall(39).BackColor = QBColor(C)
If PB(2) = 40 Then txtBall(40).BackColor = QBColor(C)
If PB(2) = 41 Then txtBall(41).BackColor = QBColor(C)
If PB(2) = 42 Then txtBall(42).BackColor = QBColor(C)
If PB(2) = 43 Then txtBall(43).BackColor = QBColor(C)
If PB(2) = 44 Then txtBall(44).BackColor = QBColor(C)
If PB(2) = 45 Then txtBall(45).BackColor = QBColor(C)
If PB(2) = 46 Then txtBall(46).BackColor = QBColor(C)
If PB(2) = 47 Then txtBall(47).BackColor = QBColor(C)
If PB(2) = 48 Then txtBall(48).BackColor = QBColor(C)
If PB(2) = 49 Then txtBall(49).BackColor = QBColor(C)
If PB(3) = 1 Then txtBall(1).BackColor = QBColor(C)
If PB(3) = 2 Then txtBall(2).BackColor = QBColor(C)
If PB(3) = 3 Then txtBall(3).BackColor = QBColor(C)
If PB(3) = 4 Then txtBall(4).BackColor = QBColor(C)
If PB(3) = 5 Then txtBall(5).BackColor = QBColor(C)
If PB(3) = 6 Then txtBall(6).BackColor = QBColor(C)
If PB(3) = 7 Then txtBall(7).BackColor = QBColor(C)
If PB(3) = 8 Then txtBall(8).BackColor = QBColor(C)
If PB(3) = 9 Then txtBall(9).BackColor = QBColor(C)
If PB(3) = 10 Then txtBall(10).BackColor = QBColor(C)
If PB(3) = 11 Then txtBall(11).BackColor = QBColor(C)
If PB(3) = 12 Then txtBall(12).BackColor = QBColor(C)
If PB(3) = 13 Then txtBall(13).BackColor = QBColor(C)
If PB(3) = 14 Then txtBall(14).BackColor = QBColor(C)
If PB(3) = 15 Then txtBall(15).BackColor = QBColor(C)
If PB(3) = 16 Then txtBall(16).BackColor = QBColor(C)
If PB(3) = 17 Then txtBall(17).BackColor = QBColor(C)
If PB(3) = 18 Then txtBall(18).BackColor = QBColor(C)
If PB(3) = 19 Then txtBall(19).BackColor = QBColor(C)
If PB(3) = 20 Then txtBall(20).BackColor = QBColor(C)
If PB(3) = 21 Then txtBall(21).BackColor = QBColor(C)
If PB(3) = 22 Then txtBall(22).BackColor = QBColor(C)
If PB(3) = 23 Then txtBall(23).BackColor = QBColor(C)
If PB(3) = 24 Then txtBall(24).BackColor = QBColor(C)
If PB(3) = 25 Then txtBall(25).BackColor = QBColor(C)
If PB(3) = 26 Then txtBall(26).BackColor = QBColor(C)
If PB(3) = 27 Then txtBall(27).BackColor = QBColor(C)
If PB(3) = 28 Then txtBall(28).BackColor = QBColor(C)
If PB(3) = 29 Then txtBall(29).BackColor = QBColor(C)
If PB(3) = 30 Then txtBall(30).BackColor = QBColor(C)
If PB(3) = 31 Then txtBall(31).BackColor = QBColor(C)
If PB(3) = 32 Then txtBall(32).BackColor = QBColor(C)
If PB(3) = 33 Then txtBall(33).BackColor = QBColor(C)
If PB(3) = 34 Then txtBall(34).BackColor = QBColor(C)
If PB(3) = 35 Then txtBall(35).BackColor = QBColor(C)
If PB(3) = 36 Then txtBall(36).BackColor = QBColor(C)
If PB(3) = 37 Then txtBall(37).BackColor = QBColor(C)
If PB(3) = 38 Then txtBall(38).BackColor = QBColor(C)
If PB(3) = 39 Then txtBall(39).BackColor = QBColor(C)
If PB(3) = 40 Then txtBall(40).BackColor = QBColor(C)
If PB(3) = 41 Then txtBall(41).BackColor = QBColor(C)
If PB(3) = 42 Then txtBall(42).BackColor = QBColor(C)
If PB(3) = 43 Then txtBall(43).BackColor = QBColor(C)
If PB(3) = 44 Then txtBall(44).BackColor = QBColor(C)
If PB(3) = 45 Then txtBall(45).BackColor = QBColor(C)
If PB(3) = 46 Then txtBall(46).BackColor = QBColor(C)
If PB(3) = 47 Then txtBall(47).BackColor = QBColor(C)
If PB(3) = 48 Then txtBall(48).BackColor = QBColor(C)
If PB(3) = 49 Then txtBall(49).BackColor = QBColor(C)
If PB(4) = 1 Then txtBall(1).BackColor = QBColor(C)
If PB(4) = 2 Then txtBall(2).BackColor = QBColor(C)
If PB(4) = 3 Then txtBall(3).BackColor = QBColor(C)
If PB(4) = 4 Then txtBall(4).BackColor = QBColor(C)
If PB(4) = 5 Then txtBall(5).BackColor = QBColor(C)
If PB(4) = 6 Then txtBall(6).BackColor = QBColor(C)
If PB(4) = 7 Then txtBall(7).BackColor = QBColor(C)
If PB(4) = 8 Then txtBall(8).BackColor = QBColor(C)
If PB(4) = 9 Then txtBall(9).BackColor = QBColor(C)
If PB(4) = 10 Then txtBall(10).BackColor = QBColor(C)
If PB(4) = 11 Then txtBall(11).BackColor = QBColor(C)
If PB(4) = 12 Then txtBall(12).BackColor = QBColor(C)
If PB(4) = 13 Then txtBall(13).BackColor = QBColor(C)
If PB(4) = 14 Then txtBall(14).BackColor = QBColor(C)
If PB(4) = 15 Then txtBall(15).BackColor = QBColor(C)
If PB(4) = 16 Then txtBall(16).BackColor = QBColor(C)
If PB(4) = 17 Then txtBall(17).BackColor = QBColor(C)
If PB(4) = 18 Then txtBall(18).BackColor = QBColor(C)
If PB(4) = 19 Then txtBall(19).BackColor = QBColor(C)
If PB(4) = 20 Then txtBall(20).BackColor = QBColor(C)
If PB(4) = 21 Then txtBall(21).BackColor = QBColor(C)
If PB(4) = 22 Then txtBall(22).BackColor = QBColor(C)
If PB(4) = 23 Then txtBall(23).BackColor = QBColor(C)
If PB(4) = 24 Then txtBall(24).BackColor = QBColor(C)
If PB(4) = 25 Then txtBall(25).BackColor = QBColor(C)
If PB(4) = 26 Then txtBall(26).BackColor = QBColor(C)
If PB(4) = 27 Then txtBall(27).BackColor = QBColor(C)
If PB(4) = 28 Then txtBall(28).BackColor = QBColor(C)
If PB(4) = 29 Then txtBall(29).BackColor = QBColor(C)
If PB(4) = 30 Then txtBall(30).BackColor = QBColor(C)
If PB(4) = 31 Then txtBall(31).BackColor = QBColor(C)
If PB(4) = 32 Then txtBall(32).BackColor = QBColor(C)
If PB(4) = 33 Then txtBall(33).BackColor = QBColor(C)
If PB(4) = 34 Then txtBall(34).BackColor = QBColor(C)
If PB(4) = 35 Then txtBall(35).BackColor = QBColor(C)
If PB(4) = 36 Then txtBall(36).BackColor = QBColor(C)
If PB(4) = 37 Then txtBall(37).BackColor = QBColor(C)
If PB(4) = 38 Then txtBall(38).BackColor = QBColor(C)
If PB(4) = 39 Then txtBall(39).BackColor = QBColor(C)
If PB(4) = 40 Then txtBall(40).BackColor = QBColor(C)
If PB(4) = 41 Then txtBall(41).BackColor = QBColor(C)
If PB(4) = 42 Then txtBall(42).BackColor = QBColor(C)
If PB(4) = 43 Then txtBall(43).BackColor = QBColor(C)
If PB(4) = 44 Then txtBall(44).BackColor = QBColor(C)
If PB(4) = 45 Then txtBall(45).BackColor = QBColor(C)
If PB(4) = 46 Then txtBall(46).BackColor = QBColor(C)
If PB(4) = 47 Then txtBall(47).BackColor = QBColor(C)
If PB(4) = 48 Then txtBall(48).BackColor = QBColor(C)
If PB(4) = 49 Then txtBall(49).BackColor = QBColor(C)
If PB(5) = 1 Then txtBall(1).BackColor = QBColor(C)
If PB(5) = 2 Then txtBall(2).BackColor = QBColor(C)
If PB(5) = 3 Then txtBall(3).BackColor = QBColor(C)
If PB(5) = 4 Then txtBall(4).BackColor = QBColor(C)
If PB(5) = 5 Then txtBall(5).BackColor = QBColor(C)
If PB(5) = 6 Then txtBall(6).BackColor = QBColor(C)
If PB(5) = 7 Then txtBall(7).BackColor = QBColor(C)
If PB(5) = 8 Then txtBall(8).BackColor = QBColor(C)
If PB(5) = 9 Then txtBall(9).BackColor = QBColor(C)
If PB(5) = 10 Then txtBall(10).BackColor = QBColor(C)
If PB(5) = 11 Then txtBall(11).BackColor = QBColor(C)
If PB(5) = 12 Then txtBall(12).BackColor = QBColor(C)
If PB(5) = 13 Then txtBall(13).BackColor = QBColor(C)
If PB(5) = 14 Then txtBall(14).BackColor = QBColor(C)
If PB(5) = 15 Then txtBall(15).BackColor = QBColor(C)
If PB(5) = 16 Then txtBall(16).BackColor = QBColor(C)
If PB(5) = 17 Then txtBall(17).BackColor = QBColor(C)
If PB(5) = 18 Then txtBall(18).BackColor = QBColor(C)
If PB(5) = 19 Then txtBall(19).BackColor = QBColor(C)
If PB(5) = 20 Then txtBall(20).BackColor = QBColor(C)
If PB(5) = 21 Then txtBall(21).BackColor = QBColor(C)
If PB(5) = 22 Then txtBall(22).BackColor = QBColor(C)
If PB(5) = 23 Then txtBall(23).BackColor = QBColor(C)
If PB(5) = 24 Then txtBall(24).BackColor = QBColor(C)
If PB(5) = 25 Then txtBall(25).BackColor = QBColor(C)
If PB(5) = 26 Then txtBall(26).BackColor = QBColor(C)
If PB(5) = 27 Then txtBall(27).BackColor = QBColor(C)
If PB(5) = 28 Then txtBall(28).BackColor = QBColor(C)
If PB(5) = 29 Then txtBall(29).BackColor = QBColor(C)
If PB(5) = 30 Then txtBall(30).BackColor = QBColor(C)
If PB(5) = 31 Then txtBall(31).BackColor = QBColor(C)
If PB(5) = 32 Then txtBall(32).BackColor = QBColor(C)
If PB(5) = 33 Then txtBall(33).BackColor = QBColor(C)
If PB(5) = 34 Then txtBall(34).BackColor = QBColor(C)
If PB(5) = 35 Then txtBall(35).BackColor = QBColor(C)
If PB(5) = 36 Then txtBall(36).BackColor = QBColor(C)
If PB(5) = 37 Then txtBall(37).BackColor = QBColor(C)
If PB(5) = 38 Then txtBall(38).BackColor = QBColor(C)
If PB(5) = 39 Then txtBall(39).BackColor = QBColor(C)
If PB(5) = 40 Then txtBall(40).BackColor = QBColor(C)
If PB(5) = 41 Then txtBall(41).BackColor = QBColor(C)
If PB(5) = 42 Then txtBall(42).BackColor = QBColor(C)
If PB(5) = 43 Then txtBall(43).BackColor = QBColor(C)
If PB(5) = 44 Then txtBall(44).BackColor = QBColor(C)
If PB(5) = 45 Then txtBall(45).BackColor = QBColor(C)
If PB(5) = 46 Then txtBall(46).BackColor = QBColor(C)
If PB(5) = 47 Then txtBall(47).BackColor = QBColor(C)
If PB(5) = 48 Then txtBall(48).BackColor = QBColor(C)
If PB(5) = 49 Then txtBall(49).BackColor = QBColor(C)
If PB(6) = 1 Then txtBall(1).BackColor = QBColor(C)
If PB(6) = 2 Then txtBall(2).BackColor = QBColor(C)
If PB(6) = 3 Then txtBall(3).BackColor = QBColor(C)
If PB(6) = 4 Then txtBall(4).BackColor = QBColor(C)
If PB(6) = 5 Then txtBall(5).BackColor = QBColor(C)
If PB(6) = 6 Then txtBall(6).BackColor = QBColor(C)
If PB(6) = 7 Then txtBall(7).BackColor = QBColor(C)
If PB(6) = 8 Then txtBall(8).BackColor = QBColor(C)
If PB(6) = 9 Then txtBall(9).BackColor = QBColor(C)
If PB(6) = 10 Then txtBall(10).BackColor = QBColor(C)
If PB(6) = 11 Then txtBall(11).BackColor = QBColor(C)
If PB(6) = 12 Then txtBall(12).BackColor = QBColor(C)
If PB(6) = 13 Then txtBall(13).BackColor = QBColor(C)
If PB(6) = 14 Then txtBall(14).BackColor = QBColor(C)
If PB(6) = 15 Then txtBall(15).BackColor = QBColor(C)
If PB(6) = 16 Then txtBall(16).BackColor = QBColor(C)
If PB(6) = 17 Then txtBall(17).BackColor = QBColor(C)
If PB(6) = 18 Then txtBall(18).BackColor = QBColor(C)
If PB(6) = 19 Then txtBall(19).BackColor = QBColor(C)
If PB(6) = 20 Then txtBall(20).BackColor = QBColor(C)
If PB(6) = 21 Then txtBall(21).BackColor = QBColor(C)
If PB(6) = 22 Then txtBall(22).BackColor = QBColor(C)
If PB(6) = 23 Then txtBall(23).BackColor = QBColor(C)
If PB(6) = 24 Then txtBall(24).BackColor = QBColor(C)
If PB(6) = 25 Then txtBall(25).BackColor = QBColor(C)
If PB(6) = 26 Then txtBall(26).BackColor = QBColor(C)
If PB(6) = 27 Then txtBall(27).BackColor = QBColor(C)
If PB(6) = 28 Then txtBall(28).BackColor = QBColor(C)
If PB(6) = 29 Then txtBall(29).BackColor = QBColor(C)
If PB(6) = 30 Then txtBall(30).BackColor = QBColor(C)
If PB(6) = 31 Then txtBall(31).BackColor = QBColor(C)
If PB(6) = 32 Then txtBall(32).BackColor = QBColor(C)
If PB(6) = 33 Then txtBall(33).BackColor = QBColor(C)
If PB(6) = 34 Then txtBall(34).BackColor = QBColor(C)
If PB(6) = 35 Then txtBall(35).BackColor = QBColor(C)
If PB(6) = 36 Then txtBall(36).BackColor = QBColor(C)
If PB(6) = 37 Then txtBall(37).BackColor = QBColor(C)
If PB(6) = 38 Then txtBall(38).BackColor = QBColor(C)
If PB(6) = 39 Then txtBall(39).BackColor = QBColor(C)
If PB(6) = 40 Then txtBall(40).BackColor = QBColor(C)
If PB(6) = 41 Then txtBall(41).BackColor = QBColor(C)
If PB(6) = 42 Then txtBall(42).BackColor = QBColor(C)
If PB(6) = 43 Then txtBall(43).BackColor = QBColor(C)
If PB(6) = 44 Then txtBall(44).BackColor = QBColor(C)
If PB(6) = 45 Then txtBall(45).BackColor = QBColor(C)
If PB(6) = 46 Then txtBall(46).BackColor = QBColor(C)
If PB(6) = 47 Then txtBall(47).BackColor = QBColor(C)
If PB(6) = 48 Then txtBall(48).BackColor = QBColor(C)
If PB(6) = 49 Then txtBall(49).BackColor = QBColor(C)
End If
If SW = 1 Then
txtBall(1).Caption = Str(B(1))
txtBall(2).Caption = Str(B(2))
txtBall(3).Caption = Str(B(3))
txtBall(4).Caption = Str(B(4))
txtBall(5).Caption = Str(B(5))
txtBall(6).Caption = Str(B(6))
txtBall(7).Caption = Str(B(7))
txtBall(8).Caption = Str(B(8))
txtBall(9).Caption = Str(B(9))
txtBall(10).Caption = Str(B(10))
txtBall(11).Caption = Str(B(11))
txtBall(12).Caption = Str(B(12))
txtBall(13).Caption = Str(B(13))
txtBall(14).Caption = Str(B(14))
txtBall(15).Caption = Str(B(15))
txtBall(16).Caption = Str(B(16))
txtBall(17).Caption = Str(B(17))
txtBall(18).Caption = Str(B(18))
txtBall(19).Caption = Str(B(19))
txtBall(20).Caption = Str(B(20))
txtBall(21).Caption = Str(B(21))
txtBall(22).Caption = Str(B(22))
txtBall(23).Caption = Str(B(23))
txtBall(24).Caption = Str(B(24))
txtBall(25).Caption = Str(B(25))
txtBall(26).Caption = Str(B(26))
txtBall(27).Caption = Str(B(27))
txtBall(28).Caption = Str(B(28))
txtBall(29).Caption = Str(B(29))
txtBall(30).Caption = Str(B(30))
txtBall(31).Caption = Str(B(31))
txtBall(32).Caption = Str(B(32))
txtBall(33).Caption = Str(B(33))
txtBall(34).Caption = Str(B(34))
txtBall(35).Caption = Str(B(35))
txtBall(36).Caption = Str(B(36))
txtBall(37).Caption = Str(B(37))
txtBall(38).Caption = Str(B(38))
txtBall(39).Caption = Str(B(39))
txtBall(40).Caption = Str(B(40))
txtBall(41).Caption = Str(B(41))
txtBall(42).Caption = Str(B(42))
txtBall(43).Caption = Str(B(43))
txtBall(44).Caption = Str(B(44))
txtBall(45).Caption = Str(B(45))
txtBall(46).Caption = Str(B(46))
txtBall(47).Caption = Str(B(47))
txtBall(48).Caption = Str(B(48))
txtBall(49).Caption = Str(B(49))
End If
If SW = 0 Then
txtBall(1).Caption = Str(Per(1))
txtBall(2).Caption = Str(Per(2))
txtBall(3).Caption = Str(Per(3))
txtBall(4).Caption = Str(Per(4))
txtBall(5).Caption = Str(Per(5))
txtBall(6).Caption = Str(Per(6))
txtBall(7).Caption = Str(Per(7))
txtBall(8).Caption = Str(Per(8))
txtBall(9).Caption = Str(Per(9))
txtBall(10).Caption = Str(Per(10))
txtBall(11).Caption = Str(Per(11))
txtBall(12).Caption = Str(Per(12))
txtBall(13).Caption = Str(Per(13))
txtBall(14).Caption = Str(Per(14))
txtBall(15).Caption = Str(Per(15))
txtBall(16).Caption = Str(Per(16))
txtBall(17).Caption = Str(Per(17))
txtBall(18).Caption = Str(Per(18))
txtBall(19).Caption = Str(Per(19))
txtBall(20).Caption = Str(Per(20))
txtBall(21).Caption = Str(Per(21))
txtBall(22).Caption = Str(Per(22))
txtBall(23).Caption = Str(Per(23))
txtBall(24).Caption = Str(Per(24))
txtBall(25).Caption = Str(Per(25))
txtBall(26).Caption = Str(Per(26))
txtBall(27).Caption = Str(Per(27))
txtBall(28).Caption = Str(Per(28))
txtBall(29).Caption = Str(Per(29))
txtBall(30).Caption = Str(Per(30))
txtBall(31).Caption = Str(Per(31))
txtBall(32).Caption = Str(Per(32))
txtBall(33).Caption = Str(Per(33))
txtBall(34).Caption = Str(Per(34))
txtBall(35).Caption = Str(Per(35))
txtBall(36).Caption = Str(Per(36))
txtBall(37).Caption = Str(Per(37))
txtBall(38).Caption = Str(Per(38))
txtBall(39).Caption = Str(Per(39))
txtBall(40).Caption = Str(Per(40))
txtBall(41).Caption = Str(Per(41))
txtBall(42).Caption = Str(Per(42))
txtBall(43).Caption = Str(Per(43))
txtBall(44).Caption = Str(Per(44))
txtBall(45).Caption = Str(Per(45))
txtBall(46).Caption = Str(Per(46))
txtBall(47).Caption = Str(Per(47))
txtBall(48).Caption = Str(Per(48))
txtBall(49).Caption = Str(Per(49))
End If
End Function

Public Function Order3()
M = 41: L = 1
If B(41) < B(1) Then L = L + 1
If B(41) < B(2) Then L = L + 1
If B(41) < B(3) Then L = L + 1
If B(41) < B(4) Then L = L + 1
If B(41) < B(5) Then L = L + 1
If B(41) < B(6) Then L = L + 1
If B(41) < B(7) Then L = L + 1
If B(41) < B(8) Then L = L + 1
If B(41) < B(9) Then L = L + 1
If B(41) < B(10) Then L = L + 1
If B(41) < B(11) Then L = L + 1
If B(41) < B(12) Then L = L + 1
If B(41) < B(13) Then L = L + 1
If B(41) < B(14) Then L = L + 1
If B(41) < B(15) Then L = L + 1
If B(41) < B(16) Then L = L + 1
If B(41) < B(17) Then L = L + 1
If B(41) < B(18) Then L = L + 1
If B(41) < B(19) Then L = L + 1
If B(41) < B(20) Then L = L + 1
If B(41) < B(21) Then L = L + 1
If B(41) < B(22) Then L = L + 1
If B(41) < B(23) Then L = L + 1
If B(41) < B(24) Then L = L + 1
If B(41) < B(25) Then L = L + 1
If B(41) < B(26) Then L = L + 1
If B(41) < B(27) Then L = L + 1
If B(41) < B(28) Then L = L + 1
If B(41) < B(29) Then L = L + 1
If B(41) < B(30) Then L = L + 1
If B(41) < B(31) Then L = L + 1
If B(41) < B(32) Then L = L + 1
If B(41) < B(33) Then L = L + 1
If B(41) < B(34) Then L = L + 1
If B(41) < B(35) Then L = L + 1
If B(41) < B(36) Then L = L + 1
If B(41) < B(37) Then L = L + 1
If B(41) < B(38) Then L = L + 1
If B(41) < B(39) Then L = L + 1
If B(41) < B(40) Then L = L + 1
If B(41) <= B(42) Then L = L + 1
If B(41) <= B(43) Then L = L + 1
If B(41) <= B(44) Then L = L + 1
If B(41) <= B(45) Then L = L + 1
If B(41) <= B(46) Then L = L + 1
If B(41) <= B(47) Then L = L + 1
If B(41) <= B(48) Then L = L + 1
If B(41) <= B(49) Then L = L + 1
 orderdisp
M = 42: L = 1
If B(42) < B(1) Then L = L + 1
If B(42) < B(2) Then L = L + 1
If B(42) < B(3) Then L = L + 1
If B(42) < B(4) Then L = L + 1
If B(42) < B(5) Then L = L + 1
If B(42) < B(6) Then L = L + 1
If B(42) < B(7) Then L = L + 1
If B(42) < B(8) Then L = L + 1
If B(42) < B(9) Then L = L + 1
If B(42) < B(10) Then L = L + 1
If B(42) < B(11) Then L = L + 1
If B(42) < B(12) Then L = L + 1
If B(42) < B(13) Then L = L + 1
If B(42) < B(14) Then L = L + 1
If B(42) < B(15) Then L = L + 1
If B(42) < B(16) Then L = L + 1
If B(42) < B(17) Then L = L + 1
If B(42) < B(18) Then L = L + 1
If B(42) < B(19) Then L = L + 1
If B(42) < B(20) Then L = L + 1
If B(42) < B(21) Then L = L + 1
If B(42) < B(22) Then L = L + 1
If B(42) < B(23) Then L = L + 1
If B(42) < B(24) Then L = L + 1
If B(42) < B(25) Then L = L + 1
If B(42) < B(26) Then L = L + 1
If B(42) < B(27) Then L = L + 1
If B(42) < B(28) Then L = L + 1
If B(42) < B(29) Then L = L + 1
If B(42) < B(30) Then L = L + 1
If B(42) < B(31) Then L = L + 1
If B(42) < B(32) Then L = L + 1
If B(42) < B(33) Then L = L + 1
If B(42) < B(34) Then L = L + 1
If B(42) < B(35) Then L = L + 1
If B(42) < B(36) Then L = L + 1
If B(42) < B(37) Then L = L + 1
If B(42) < B(38) Then L = L + 1
If B(42) < B(39) Then L = L + 1
If B(42) < B(40) Then L = L + 1
If B(42) < B(41) Then L = L + 1
If B(42) <= B(43) Then L = L + 1
If B(42) <= B(44) Then L = L + 1
If B(42) <= B(45) Then L = L + 1
If B(42) <= B(46) Then L = L + 1
If B(42) <= B(47) Then L = L + 1
If B(42) <= B(48) Then L = L + 1
If B(42) <= B(49) Then L = L + 1
 orderdisp
M = 43: L = 1
If B(43) < B(1) Then L = L + 1
If B(43) < B(2) Then L = L + 1
If B(43) < B(3) Then L = L + 1
If B(43) < B(4) Then L = L + 1
If B(43) < B(5) Then L = L + 1
If B(43) < B(6) Then L = L + 1
If B(43) < B(7) Then L = L + 1
If B(43) < B(8) Then L = L + 1
If B(43) < B(9) Then L = L + 1
If B(43) < B(10) Then L = L + 1
If B(43) < B(11) Then L = L + 1
If B(43) < B(12) Then L = L + 1
If B(43) < B(13) Then L = L + 1
If B(43) < B(14) Then L = L + 1
If B(43) < B(15) Then L = L + 1
If B(43) < B(16) Then L = L + 1
If B(43) < B(17) Then L = L + 1
If B(43) < B(18) Then L = L + 1
If B(43) < B(19) Then L = L + 1
If B(43) < B(20) Then L = L + 1
If B(43) < B(21) Then L = L + 1
If B(43) < B(22) Then L = L + 1
If B(43) < B(23) Then L = L + 1
If B(43) < B(24) Then L = L + 1
If B(43) < B(25) Then L = L + 1
If B(43) < B(26) Then L = L + 1
If B(43) < B(27) Then L = L + 1
If B(43) < B(28) Then L = L + 1
If B(43) < B(29) Then L = L + 1
If B(43) < B(30) Then L = L + 1
If B(43) < B(31) Then L = L + 1
If B(43) < B(32) Then L = L + 1
If B(43) < B(33) Then L = L + 1
If B(43) < B(34) Then L = L + 1
If B(43) < B(35) Then L = L + 1
If B(43) < B(36) Then L = L + 1
If B(43) < B(37) Then L = L + 1
If B(43) < B(38) Then L = L + 1
If B(43) < B(39) Then L = L + 1
If B(43) < B(40) Then L = L + 1
If B(43) < B(41) Then L = L + 1
If B(43) < B(42) Then L = L + 1
If B(43) <= B(44) Then L = L + 1
If B(43) <= B(45) Then L = L + 1
If B(43) <= B(46) Then L = L + 1
If B(43) <= B(47) Then L = L + 1
If B(43) <= B(48) Then L = L + 1
If B(43) <= B(49) Then L = L + 1
 orderdisp
M = 44: L = 1
If B(44) < B(1) Then L = L + 1
If B(44) < B(2) Then L = L + 1
If B(44) < B(3) Then L = L + 1
If B(44) < B(4) Then L = L + 1
If B(44) < B(5) Then L = L + 1
If B(44) < B(6) Then L = L + 1
If B(44) < B(7) Then L = L + 1
If B(44) < B(8) Then L = L + 1
If B(44) < B(9) Then L = L + 1
If B(44) < B(10) Then L = L + 1
If B(44) < B(11) Then L = L + 1
If B(44) < B(12) Then L = L + 1
If B(44) < B(13) Then L = L + 1
If B(44) < B(14) Then L = L + 1
If B(44) < B(15) Then L = L + 1
If B(44) < B(16) Then L = L + 1
If B(44) < B(17) Then L = L + 1
If B(44) < B(18) Then L = L + 1
If B(44) < B(19) Then L = L + 1
If B(44) < B(20) Then L = L + 1
If B(44) < B(21) Then L = L + 1
If B(44) < B(22) Then L = L + 1
If B(44) < B(23) Then L = L + 1
If B(44) < B(24) Then L = L + 1
If B(44) < B(25) Then L = L + 1
If B(44) < B(26) Then L = L + 1
If B(44) < B(27) Then L = L + 1
If B(44) < B(28) Then L = L + 1
If B(44) < B(29) Then L = L + 1
If B(44) < B(30) Then L = L + 1
If B(44) < B(31) Then L = L + 1
If B(44) < B(32) Then L = L + 1
If B(44) < B(33) Then L = L + 1
If B(44) < B(34) Then L = L + 1
If B(44) < B(35) Then L = L + 1
If B(44) < B(36) Then L = L + 1
If B(44) < B(37) Then L = L + 1
If B(44) < B(38) Then L = L + 1
If B(44) < B(39) Then L = L + 1
If B(44) < B(40) Then L = L + 1
If B(44) < B(41) Then L = L + 1
If B(44) < B(42) Then L = L + 1
If B(44) < B(43) Then L = L + 1
If B(44) <= B(45) Then L = L + 1
If B(44) <= B(46) Then L = L + 1
If B(44) <= B(47) Then L = L + 1
If B(44) <= B(48) Then L = L + 1
If B(44) <= B(49) Then L = L + 1
 orderdisp
M = 45: L = 1
If B(45) < B(1) Then L = L + 1
If B(45) < B(2) Then L = L + 1
If B(45) < B(3) Then L = L + 1
If B(45) < B(4) Then L = L + 1
If B(45) < B(5) Then L = L + 1
If B(45) < B(6) Then L = L + 1
If B(45) < B(7) Then L = L + 1
If B(45) < B(8) Then L = L + 1
If B(45) < B(9) Then L = L + 1
If B(45) < B(10) Then L = L + 1
If B(45) < B(11) Then L = L + 1
If B(45) < B(12) Then L = L + 1
If B(45) < B(13) Then L = L + 1
If B(45) < B(14) Then L = L + 1
If B(45) < B(15) Then L = L + 1
If B(45) < B(16) Then L = L + 1
If B(45) < B(17) Then L = L + 1
If B(45) < B(18) Then L = L + 1
If B(45) < B(19) Then L = L + 1
If B(45) < B(20) Then L = L + 1
If B(45) < B(21) Then L = L + 1
If B(45) < B(22) Then L = L + 1
If B(45) < B(23) Then L = L + 1
If B(45) < B(24) Then L = L + 1
If B(45) < B(25) Then L = L + 1
If B(45) < B(26) Then L = L + 1
If B(45) < B(27) Then L = L + 1
If B(45) < B(28) Then L = L + 1
If B(45) < B(29) Then L = L + 1
If B(45) < B(30) Then L = L + 1
If B(45) < B(31) Then L = L + 1
If B(45) < B(32) Then L = L + 1
If B(45) < B(33) Then L = L + 1
If B(45) < B(34) Then L = L + 1
If B(45) < B(35) Then L = L + 1
If B(45) < B(36) Then L = L + 1
If B(45) < B(37) Then L = L + 1
If B(45) < B(38) Then L = L + 1
If B(45) < B(39) Then L = L + 1
If B(45) < B(40) Then L = L + 1
If B(45) < B(41) Then L = L + 1
If B(45) < B(42) Then L = L + 1
If B(45) < B(43) Then L = L + 1
If B(45) < B(44) Then L = L + 1
If B(45) <= B(46) Then L = L + 1
If B(45) <= B(47) Then L = L + 1
If B(45) <= B(48) Then L = L + 1
If B(45) <= B(49) Then L = L + 1
 orderdisp
M = 46: L = 1
If B(46) < B(1) Then L = L + 1
If B(46) < B(2) Then L = L + 1
If B(46) < B(3) Then L = L + 1
If B(46) < B(4) Then L = L + 1
If B(46) < B(5) Then L = L + 1
If B(46) < B(6) Then L = L + 1
If B(46) < B(7) Then L = L + 1
If B(46) < B(8) Then L = L + 1
If B(46) < B(9) Then L = L + 1
If B(46) < B(10) Then L = L + 1
If B(46) < B(11) Then L = L + 1
If B(46) < B(12) Then L = L + 1
If B(46) < B(13) Then L = L + 1
If B(46) < B(14) Then L = L + 1
If B(46) < B(15) Then L = L + 1
If B(46) < B(16) Then L = L + 1
If B(46) < B(17) Then L = L + 1
If B(46) < B(18) Then L = L + 1
If B(46) < B(19) Then L = L + 1
If B(46) < B(20) Then L = L + 1
If B(46) < B(21) Then L = L + 1
If B(46) < B(22) Then L = L + 1
If B(46) < B(23) Then L = L + 1
If B(46) < B(24) Then L = L + 1
If B(46) < B(25) Then L = L + 1
If B(46) < B(26) Then L = L + 1
If B(46) < B(27) Then L = L + 1
If B(46) < B(28) Then L = L + 1
If B(46) < B(29) Then L = L + 1
If B(46) < B(30) Then L = L + 1
If B(46) < B(31) Then L = L + 1
If B(46) < B(32) Then L = L + 1
If B(46) < B(33) Then L = L + 1
If B(46) < B(34) Then L = L + 1
If B(46) < B(35) Then L = L + 1
If B(46) < B(36) Then L = L + 1
If B(46) < B(37) Then L = L + 1
If B(46) < B(38) Then L = L + 1
If B(46) < B(39) Then L = L + 1
If B(46) < B(40) Then L = L + 1
If B(46) < B(41) Then L = L + 1
If B(46) < B(42) Then L = L + 1
If B(46) < B(43) Then L = L + 1
If B(46) < B(44) Then L = L + 1
If B(46) < B(45) Then L = L + 1
If B(46) <= B(47) Then L = L + 1
If B(46) <= B(48) Then L = L + 1
If B(46) <= B(49) Then L = L + 1
 orderdisp
M = 47: L = 1
If B(47) < B(1) Then L = L + 1
If B(47) < B(2) Then L = L + 1
If B(47) < B(3) Then L = L + 1
If B(47) < B(4) Then L = L + 1
If B(47) < B(5) Then L = L + 1
If B(47) < B(6) Then L = L + 1
If B(47) < B(7) Then L = L + 1
If B(47) < B(8) Then L = L + 1
If B(47) < B(9) Then L = L + 1
If B(47) < B(10) Then L = L + 1
If B(47) < B(11) Then L = L + 1
If B(47) < B(12) Then L = L + 1
If B(47) < B(13) Then L = L + 1
If B(47) < B(14) Then L = L + 1
If B(47) < B(15) Then L = L + 1
If B(47) < B(16) Then L = L + 1
If B(47) < B(17) Then L = L + 1
If B(47) < B(18) Then L = L + 1
If B(47) < B(19) Then L = L + 1
If B(47) < B(20) Then L = L + 1
If B(47) < B(21) Then L = L + 1
If B(47) < B(22) Then L = L + 1
If B(47) < B(23) Then L = L + 1
If B(47) < B(24) Then L = L + 1
If B(47) < B(25) Then L = L + 1
If B(47) < B(26) Then L = L + 1
If B(47) < B(27) Then L = L + 1
If B(47) < B(28) Then L = L + 1
If B(47) < B(29) Then L = L + 1
If B(47) < B(30) Then L = L + 1
If B(47) < B(31) Then L = L + 1
If B(47) < B(32) Then L = L + 1
If B(47) < B(33) Then L = L + 1
If B(47) < B(34) Then L = L + 1
If B(47) < B(35) Then L = L + 1
If B(47) < B(36) Then L = L + 1
If B(47) < B(37) Then L = L + 1
If B(47) < B(38) Then L = L + 1
If B(47) < B(39) Then L = L + 1
If B(47) < B(40) Then L = L + 1
If B(47) < B(41) Then L = L + 1
If B(47) < B(42) Then L = L + 1
If B(47) < B(43) Then L = L + 1
If B(47) < B(44) Then L = L + 1
If B(47) < B(45) Then L = L + 1
If B(47) < B(46) Then L = L + 1
If B(47) <= B(48) Then L = L + 1
If B(47) <= B(49) Then L = L + 1
 orderdisp
M = 48: L = 1
If B(48) < B(1) Then L = L + 1
If B(48) < B(2) Then L = L + 1
If B(48) < B(3) Then L = L + 1
If B(48) < B(4) Then L = L + 1
If B(48) < B(5) Then L = L + 1
If B(48) < B(6) Then L = L + 1
If B(48) < B(7) Then L = L + 1
If B(48) < B(8) Then L = L + 1
If B(48) < B(9) Then L = L + 1
If B(48) < B(10) Then L = L + 1
If B(48) < B(11) Then L = L + 1
If B(48) < B(12) Then L = L + 1
If B(48) < B(13) Then L = L + 1
If B(48) < B(14) Then L = L + 1
If B(48) < B(15) Then L = L + 1
If B(48) < B(16) Then L = L + 1
If B(48) < B(17) Then L = L + 1
If B(48) < B(18) Then L = L + 1
If B(48) < B(19) Then L = L + 1
If B(48) < B(20) Then L = L + 1
If B(48) < B(21) Then L = L + 1
If B(48) < B(22) Then L = L + 1
If B(48) < B(23) Then L = L + 1
If B(48) < B(24) Then L = L + 1
If B(48) < B(25) Then L = L + 1
If B(48) < B(26) Then L = L + 1
If B(48) < B(27) Then L = L + 1
If B(48) < B(28) Then L = L + 1
If B(48) < B(29) Then L = L + 1
If B(48) < B(30) Then L = L + 1
If B(48) < B(31) Then L = L + 1
If B(48) < B(32) Then L = L + 1
If B(48) < B(33) Then L = L + 1
If B(48) < B(34) Then L = L + 1
If B(48) < B(35) Then L = L + 1
If B(48) < B(36) Then L = L + 1
If B(48) < B(37) Then L = L + 1
If B(48) < B(38) Then L = L + 1
If B(48) < B(39) Then L = L + 1
If B(48) < B(40) Then L = L + 1
If B(48) < B(41) Then L = L + 1
If B(48) < B(42) Then L = L + 1
If B(48) < B(43) Then L = L + 1
If B(48) < B(44) Then L = L + 1
If B(48) < B(45) Then L = L + 1
If B(48) < B(46) Then L = L + 1
If B(48) < B(47) Then L = L + 1
If B(48) <= B(49) Then L = L + 1
 orderdisp
M = 49: L = 1
If B(49) < B(1) Then L = L + 1
If B(49) < B(2) Then L = L + 1
If B(49) < B(3) Then L = L + 1
If B(49) < B(4) Then L = L + 1
If B(49) < B(5) Then L = L + 1
If B(49) < B(6) Then L = L + 1
If B(49) < B(7) Then L = L + 1
If B(49) < B(8) Then L = L + 1
If B(49) < B(9) Then L = L + 1
If B(49) < B(10) Then L = L + 1
If B(49) < B(11) Then L = L + 1
If B(49) < B(12) Then L = L + 1
If B(49) < B(13) Then L = L + 1
If B(49) < B(14) Then L = L + 1
If B(49) < B(15) Then L = L + 1
If B(49) < B(16) Then L = L + 1
If B(49) < B(17) Then L = L + 1
If B(49) < B(18) Then L = L + 1
If B(49) < B(19) Then L = L + 1
If B(49) < B(20) Then L = L + 1
If B(49) < B(21) Then L = L + 1
If B(49) < B(22) Then L = L + 1
If B(49) < B(23) Then L = L + 1
If B(49) < B(24) Then L = L + 1
If B(49) < B(25) Then L = L + 1
If B(49) < B(26) Then L = L + 1
If B(49) < B(27) Then L = L + 1
If B(49) < B(28) Then L = L + 1
If B(49) < B(29) Then L = L + 1
If B(49) < B(30) Then L = L + 1
If B(49) < B(31) Then L = L + 1
If B(49) < B(32) Then L = L + 1
If B(49) < B(33) Then L = L + 1
If B(49) < B(34) Then L = L + 1
If B(49) < B(35) Then L = L + 1
If B(49) < B(36) Then L = L + 1
If B(49) < B(37) Then L = L + 1
If B(49) < B(38) Then L = L + 1
If B(49) < B(39) Then L = L + 1
If B(49) < B(40) Then L = L + 1
If B(49) < B(41) Then L = L + 1
If B(49) < B(42) Then L = L + 1
If B(49) < B(43) Then L = L + 1
If B(49) < B(44) Then L = L + 1
If B(49) < B(45) Then L = L + 1
If B(49) < B(46) Then L = L + 1
If B(49) < B(47) Then L = L + 1
If B(49) < B(48) Then L = L + 1
orderdisp
End Function

Public Function MRandom()
MnuFixPlayer.Checked = False
mnuRandom.Checked = True
mnuEColdNums.Checked = False
mnuEHotNums.Checked = False
MnuAverage.Checked = False
ECN = 0
EHN = 0
SetRan = 1
SetFix = 0
AV = 0
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(20).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
lblPlayerStat.Caption = "(Random Numbers)"
mnuSumFilter.Enabled = True
mnuRangeFilter.Enabled = True
txtBlock.Caption = Str(IPUT)
For J = 1 To 6
txtH(J).Caption = Str(P(J))
Next
PC = 1
txtH(0).Left = 3240
txtH(0).Top = 4560
txtH(1).Left = 3240
txtH(1).Top = 4800
txtH(2).Left = 3720
txtH(2).Top = 4800
txtH(3).Left = 4200
txtH(3).Top = 4800
txtH(4).Left = 4680
txtH(4).Top = 4800
txtH(5).Left = 5160
txtH(5).Top = 4800
txtH(6).Left = 5640
txtH(6).Top = 4800
For J = 0 To 6
txtC(J).Visible = True
txtH(J).BackColor = QBColor(12)
Next
txtH(0).Caption = "6 Most Common Occurring Numbers"
SystemRestart
End Function
Public Function MFixed()
EHN = 0
ECN = 0
SetRan = 0
SetFix = 1
Sumfil = 0
RangeFil = 0
AV = 0
MnuFixPlayer.Checked = True
mnuRandom.Checked = False
mnuEColdNums.Checked = False
mnuEHotNums.Checked = False
mnuSumFilter.Caption = "Sum Filter"
TOOL10.Picture = ImageList1.ListImages(29).Picture
TOOL11.Picture = ImageList1.ListImages(27).Picture
mnuSumFilter.Enabled = False
mnuSumFilter.Checked = False
mnuRangeFilter.Checked = False
mnuRangeFilter.Caption = "Range Filter"
mnuRangeFilter.Enabled = False
lblsumtext.Visible = False
lblSum.Visible = False
lblRanText.Visible = False
lblran.Visible = False
lblPlayerStat.Caption = "('Fixed' Numbers)"
MnuAverage.Checked = False
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(19).Picture
TOOL8.Picture = ImageList1.ListImages(12).Picture
For J = 1 To 6
txtH(J).Caption = Str(P(J))
Next
txtH(0).Left = 3240
txtH(0).Top = 4560
txtH(1).Left = 3240
txtH(1).Top = 4800
txtH(2).Left = 3720
txtH(2).Top = 4800
txtH(3).Left = 4200
txtH(3).Top = 4800
txtH(4).Left = 4680
txtH(4).Top = 4800
txtH(5).Left = 5160
txtH(5).Top = 4800
txtH(6).Left = 5640
txtH(6).Top = 4800
For J = 0 To 6
txtC(J).Visible = True
txtH(J).BackColor = QBColor(12)
Next
txtH(0).Caption = "6 Most Common Occurring Numbers"
FrmFixed.Show 1
If DIC = 1 Then
DIC = 0
MRandom
GoTo 1000
End If
For J = 1 To 6
lblPb(J).Caption = Str(SF(J))
PB(J) = SF(J)
Next
SystemRestart
1000
End Function
Public Function MActual()
TOOL3.Picture = ImageList1.ListImages(34).Picture
TOOL2.Picture = ImageList1.ListImages(19).Picture
mnuActualNumbers.Checked = True
mnuRandomNumbers.Checked = False
PC = 1
LsetAct = 1
LsetRan = 0
clc
For J = 1 To 6
txtH(J).Caption = Str(P(J))
Next
txtC(1).Caption = Str(P(6))
txtC(2).Caption = Str(P(5))
txtC(3).Caption = Str(P(4))
txtC(4).Caption = Str(P(3))
txtC(5).Caption = Str(P(2))
txtC(6).Caption = Str(P(1))
lblLotteryStat.Caption = "(Actual Numbers)"
'Get #FA, FB, FC
'LB(1) = Str(FC.Ball1)
'LB(2) = Str(FC.Ball2)
'LB(3) = Str(FC.Ball3)
'LB(4) = Str(FC.Ball4)
'LB(5) = Str(FC.ball5)
'LB(6) = Str(FC.ball6)
'LB(0) = Str(FC.BallBon)
'For J = 0 To 6
'lbllb(J).Caption = Str(LB(J))
'Next
'lblDrawDate.Caption = Trim(FC.Date)
'Matcher
'NumCounter
'NumDisplay
'IPUT = FD
'txtBlock.Caption = Str(IPUT)
'NUMENT = NUMENT - Stake
'FB = 0
'ActBalls
If ECN = 1 Then
ColdBalls
EHN = 0
SetRan = 0
SetFix = 0
End If
If EHN = 1 Then
HotBalls
SetRan = 0
SetFix = 0
End If
If SetRan = 1 Then
playerballs
SetFix = 0
End If
If SetFix = 1 Then
FixedPlayerBalls
End If
Matcher
FB = 1
clc
For J = 1 To 49
txtBall(J).Caption = Str(B(J))
Next
Order
Matcher
displayNums
SystemRestart
End Function
Public Function MHot()
If EHN = 1 Then
MsgBox "Already entering in Hot numbers!"
GoTo 9
End If
mnuSumFilter.Caption = "Sum Filter"
TOOL10.Picture = ImageList1.ListImages(29).Picture
TOOL11.Picture = ImageList1.ListImages(27).Picture
mnuSumFilter.Enabled = False
mnuSumFilter.Checked = False
mnuRangeFilter.Checked = False
mnuRangeFilter.Caption = "Range Filter"
mnuRangeFilter.Enabled = False
lblsumtext.Visible = False
lblSum.Visible = False
lblRanText.Visible = False
lblran.Visible = False
Sumfil = 0
RangeFil = 0
AV = 0
MnuFixPlayer.Checked = False
mnuRandom.Checked = False
mnuEColdNums.Checked = False
mnuEHotNums.Checked = True
MnuAverage.Checked = False
TOOL5.Picture = ImageList1.ListImages(36).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(19).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
txtH(0).Left = 3240
txtH(0).Top = 4560
txtH(1).Left = 3240
txtH(1).Top = 4800
txtH(2).Left = 3720
txtH(2).Top = 4800
txtH(3).Left = 4200
txtH(3).Top = 4800
txtH(4).Left = 4680
txtH(4).Top = 4800
txtH(5).Left = 5160
txtH(5).Top = 4800
txtH(6).Left = 5640
txtH(6).Top = 4800
For J = 0 To 6
txtC(J).Visible = True
txtH(J).BackColor = QBColor(12)
Next
txtH(0).Caption = "6 Most Common Occurring Numbers"
For J = 1 To 6
txtH(J).Caption = Str(P(J))
Next
SetUpHotNumbers
9
End Function

Public Function MCold()
If ECN = 1 Then
MsgBox "Already entering in Cold numbers!"
GoTo 9
End If
mnuSumFilter.Caption = "Sum Filter"
TOOL10.Picture = ImageList1.ListImages(29).Picture
TOOL11.Picture = ImageList1.ListImages(27).Picture
mnuSumFilter.Enabled = False
mnuSumFilter.Checked = False
mnuRangeFilter.Checked = False
mnuRangeFilter.Caption = "Range Filter"
mnuRangeFilter.Enabled = False
lblsumtext.Visible = False
lblSum.Visible = False
lblRanText.Visible = False
lblran.Visible = False
MnuFixPlayer.Checked = False
mnuRandom.Checked = False
mnuEColdNums.Checked = True
mnuEHotNums.Checked = False
MnuAverage.Checked = False
For J = 1 To 6
txtH(J).Caption = Str(P(J))
Next
Sumfil = 0
RangeFil = 0
AV = 0
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(22).Picture
TOOL7.Picture = ImageList1.ListImages(19).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
txtH(0).Left = 3240
txtH(0).Top = 4560
txtH(1).Left = 3240
txtH(1).Top = 4800
txtH(2).Left = 3720
txtH(2).Top = 4800
txtH(3).Left = 4200
txtH(3).Top = 4800
txtH(4).Left = 4680
txtH(4).Top = 4800
txtH(5).Left = 5160
txtH(5).Top = 4800
txtH(6).Left = 5640
txtH(6).Top = 4800
For J = 0 To 6
txtC(J).Visible = True
txtH(J).BackColor = QBColor(12)
Next
txtH(0).Caption = "6 Most Common Occurring Numbers"
SetUpColdNumbers
9
End Function
Public Function MSumFilter()
Sumfil = Sumfil + 1
If Sumfil = 2 Then
Sumfil = 0
TOOL10.Picture = ImageList1.ListImages(29).Picture
mnuSumFilter.Caption = "Sum Filter"
mnuSumFilter.Checked = False
lblsumtext.Visible = False
lblSum.Visible = False
GoTo 500
End If
If Sumfil = 1 Then
TOOL10.Picture = ImageList1.ListImages(30).Picture
mnuSumFilter.Checked = True
lblsumtext.Visible = True
lblSum.Visible = True
FrmSumFilter.Show 1
If DIC = 1 Then
Sumfil = 0
TOOL10.Picture = ImageList1.ListImages(29).Picture
DIC = 0
mnuSumFilter.Caption = "Sum Filter"
mnuSumFilter.Checked = False
lblsumtext.Visible = False
lblSum.Visible = False
GoTo 500
End If
End If
MnuFixPlayer.Checked = False
mnuRandom.Checked = True
mnuEColdNums.Checked = False
mnuEHotNums.Checked = False
ECN = 0
EHN = 0
SetRan = 1
SetFix = 0
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(20).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
lblPlayerStat.Caption = "(Random Numbers)"
mnuSumFilter.Enabled = True
mnuRangeFilter.Enabled = True
mnuSumFilter.Caption = "Sum Filter        (ON)  Set to   Max:  " + Str(USL) + "   Min:  " + Str(LSL)
500
End Function

Public Function MRangeFilter()
ECN = 0
EHN = 0
SetRan = 1
SetFix = 0
RangeFil = RangeFil + 1
If RangeFil = 2 Then
RangeFil = 0
TOOL11.Picture = ImageList1.ListImages(27).Picture
mnuRangeFilter.Caption = "Range Filter"
mnuRangeFilter.Checked = False
lblRanText.Visible = False
lblran.Visible = False
GoTo 500
End If
If RangeFil = 1 Then
TOOL11.Picture = ImageList1.ListImages(28).Picture
mnuRangeFilter.Checked = True
lblRanText.Visible = True
lblran.Visible = True
frmRangeFilter.Show 1
If DIC = 1 Then
RangeFil = 0
TOOL11.Picture = ImageList1.ListImages(27).Picture
DIC = 0
mnuRangeFilter.Caption = "Range Filter"
mnuRangeFilter.Checked = False
lblRanText.Visible = False
lblran.Visible = False
GoTo 500
End If
End If
MnuFixPlayer.Checked = False
mnuRandom.Checked = True
mnuEColdNums.Checked = False
mnuEHotNums.Checked = False
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(20).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
lblPlayerStat.Caption = "(Random Numbers)"
mnuRangeFilter.Enabled = True
mnuRangeFilter.Caption = "Range Filter        (ON)  Set to   Max:  " + Str(URL) + "   Min:  " + Str(LRL)
500
End Function

Public Function MLRandom()
mnuRandomNumbers.Checked = True
mnuActualNumbers.Checked = False
TOOL2.Picture = ImageList1.ListImages(20).Picture
TOOL3.Picture = ImageList1.ListImages(33).Picture
LsetAct = 0
LsetRan = 1
lblLotteryStat.Caption = "(Random Numbers)"
lblDrawDate.Caption = ""
txtBlock.Caption = Str(IPUT)
PC = 1
SystemRestart
End Function

Public Function DispLott()
LD = LD + 1
If LD >= 2 Then LD = 0
If LD = 1 Then
mnuLotteryDisp.Checked = True
TOOL14.Picture = ImageList1.ListImages(18).Picture
Else
mnuLotteryDisp.Checked = False
TOOL14.Picture = ImageList1.ListImages(17).Picture
End If
clc
displayNums
NumDisplay
End Function

Public Function DispPlay()
PD = PD + 1
If PD >= 2 Then PD = 0
If PD = 1 Then
mnuPlayDisp.Checked = True
TOOL15.Picture = ImageList1.ListImages(24).Picture
Else
mnuPlayDisp.Checked = False
TOOL15.Picture = ImageList1.ListImages(23).Picture
End If
clc
displayNums
NumDisplay
End Function

Public Function DispHot()
HD = HD + 1
If HD >= 2 Then HD = 0
If HD = 1 Then
MNUhotdISP.Checked = True
TOOL16.Picture = ImageList1.ListImages(16).Picture
Else
MNUhotdISP.Checked = False
TOOL16.Picture = ImageList1.ListImages(15).Picture
End If
clc
displayNums
NumDisplay
End Function

Public Function DispCold()
CD = CD + 1
If CD >= 2 Then CD = 0
If CD = 1 Then
mnuColdDisp.Checked = True
TOOL17.Picture = ImageList1.ListImages(40).Picture
Else
mnuColdDisp.Checked = False
TOOL17.Picture = ImageList1.ListImages(39).Picture
End If
clc
displayNums
NumDisplay
End Function



Private Sub txtwn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
HPP = 25
FrmQuickHelp.Show
End If
End Sub

Private Sub Vspeed_Change()
ZZ = Vspeed.Value
Timer1.Interval = ZZ
End Sub


Private Sub Vspeed_Scroll()
ZZ = Vspeed.Value
Timer1.Interval = ZZ
End Sub



Public Function NumCounter2()
If Stack = 0 Then GoTo 100
Per(1) = B(1) / Stack * 100
Per(1) = Left(Per(1), 4)
Per(2) = B(2) / Stack * 100
Per(2) = Left(Per(2), 4)
Per(3) = B(3) / Stack * 100
Per(3) = Left(Per(3), 4)
Per(4) = B(4) / Stack * 100
Per(4) = Left(Per(4), 4)
Per(5) = B(5) / Stack * 100
Per(5) = Left(Per(5), 4)
Per(6) = B(6) / Stack * 100
Per(6) = Left(Per(6), 4)
Per(7) = B(7) / Stack * 100
Per(7) = Left(Per(7), 4)
Per(8) = B(8) / Stack * 100
Per(8) = Left(Per(8), 4)
Per(9) = B(9) / Stack * 100
Per(9) = Left(Per(9), 4)
Per(10) = B(10) / Stack * 100
Per(10) = Left(Per(10), 4)
Per(11) = B(11) / Stack * 100
Per(11) = Left(Per(11), 4)
Per(12) = B(12) / Stack * 100
Per(12) = Left(Per(12), 4)
Per(13) = B(13) / Stack * 100
Per(13) = Left(Per(13), 4)
Per(14) = B(14) / Stack * 100
Per(14) = Left(Per(14), 4)
Per(15) = B(15) / Stack * 100
Per(15) = Left(Per(15), 4)
Per(16) = B(16) / Stack * 100
Per(16) = Left(Per(16), 4)
Per(17) = B(17) / Stack * 100
Per(17) = Left(Per(17), 4)
Per(18) = B(18) / Stack * 100
Per(18) = Left(Per(18), 4)
Per(19) = B(19) / Stack * 100
Per(19) = Left(Per(19), 4)
Per(20) = B(20) / Stack * 100
Per(20) = Left(Per(20), 4)
Per(21) = B(21) / Stack * 100
Per(21) = Left(Per(21), 4)
Per(22) = B(22) / Stack * 100
Per(22) = Left(Per(22), 4)
Per(23) = B(23) / Stack * 100
Per(23) = Left(Per(23), 4)
Per(24) = B(24) / Stack * 100
Per(24) = Left(Per(24), 4)
Per(25) = B(25) / Stack * 100
Per(25) = Left(Per(25), 4)
Per(26) = B(26) / Stack * 100
Per(26) = Left(Per(26), 4)
Per(27) = B(27) / Stack * 100
Per(27) = Left(Per(27), 4)
Per(28) = B(28) / Stack * 100
Per(28) = Left(Per(28), 4)
Per(29) = B(29) / Stack * 100
Per(29) = Left(Per(29), 4)
Per(30) = B(30) / Stack * 100
Per(30) = Left(Per(30), 4)
Per(31) = B(31) / Stack * 100
Per(31) = Left(Per(31), 4)
Per(32) = B(32) / Stack * 100
Per(32) = Left(Per(32), 4)
Per(33) = B(33) / Stack * 100
Per(33) = Left(Per(33), 4)
Per(34) = B(34) / Stack * 100
Per(34) = Left(Per(34), 4)
Per(35) = B(35) / Stack * 100
Per(35) = Left(Per(35), 4)
Per(36) = B(36) / Stack * 100
Per(36) = Left(Per(36), 4)
Per(37) = B(37) / Stack * 100
Per(37) = Left(Per(37), 4)
Per(38) = B(38) / Stack * 100
Per(38) = Left(Per(38), 4)
Per(39) = B(39) / Stack * 100
Per(39) = Left(Per(39), 4)
Per(40) = B(40) / Stack * 100
Per(40) = Left(Per(40), 4)
Per(41) = B(41) / Stack * 100
Per(41) = Left(Per(41), 4)
Per(42) = B(42) / Stack * 100
Per(42) = Left(Per(42), 4)
Per(43) = B(43) / Stack * 100
Per(43) = Left(Per(43), 4)
Per(44) = B(44) / Stack * 100
Per(44) = Left(Per(44), 4)
Per(45) = B(45) / Stack * 100
Per(45) = Left(Per(45), 4)
Per(46) = B(46) / Stack * 100
Per(46) = Left(Per(46), 4)
Per(47) = B(47) / Stack * 100
Per(47) = Left(Per(47), 4)
Per(48) = B(48) / Stack * 100
Per(48) = Left(Per(48), 4)
Per(49) = B(49) / Stack * 100
Per(49) = Left(Per(49), 4)
100
End Function



Public Function Random6()
TRB(1) = TRB(1) + LB(1)
TRB(2) = TRB(2) + LB(2)
TRB(3) = TRB(3) + LB(3)
TRB(4) = TRB(4) + LB(4)
TRB(5) = TRB(5) + LB(5)
TRB(6) = TRB(6) + LB(6)
If Stack = 0 Then GoTo 100
RB(1) = TRB(1) / (Stack + 1)
RB(2) = TRB(2) / (Stack + 1)
RB(3) = TRB(3) / (Stack + 1)
RB(4) = TRB(4) / (Stack + 1)
RB(5) = TRB(5) / (Stack + 1)
RB(6) = TRB(6) / (Stack + 1)
100
End Function

Public Function MAverage()
If AV = 1 Then
MsgBox "Already entering in Average 6"
GoTo 9
End If
AV = 1
EHN = 0
ECN = 0
SetFix = 0
Sumfil = 0
RangeFil = 0
If Stack <= 1 Then
For J = 1 To 6
RB(J) = LB(J)
txtH(J).Caption = Str(LB(J))
Next
End If
TOOL10.Picture = ImageList1.ListImages(29).Picture
TOOL11.Picture = ImageList1.ListImages(27).Picture
mnuSumFilter.Enabled = False
mnuSumFilter.Checked = False
mnuSumFilter.Caption = "Sum Filter"
mnuRangeFilter.Checked = False
mnuRangeFilter.Caption = "Range Filter"
mnuRangeFilter.Enabled = False
lblsumtext.Visible = False
lblSum.Visible = False
lblRanText.Visible = False
lblran.Visible = False
MnuFixPlayer.Checked = False
mnuRandom.Checked = False
mnuEColdNums.Checked = False
mnuEHotNums.Checked = False
MnuAverage.Checked = True
lblPlayerStat.Caption = "(Average 6)"
TOOL5.Picture = ImageList1.ListImages(35).Picture
TOOL6.Picture = ImageList1.ListImages(21).Picture
TOOL7.Picture = ImageList1.ListImages(19).Picture
TOOL8.Picture = ImageList1.ListImages(11).Picture
With txtH(0)
.Left = 3240
.Top = 4800
.Caption = "Average 6"
End With
With txtH(1)
.Left = 3240
.Top = 5040
.Caption = RB(1)
End With
With txtH(2)
.Left = 3720
.Top = 5040
.Caption = RB(2)
End With
With txtH(3)
.Left = 4200
.Top = 5040
.Caption = RB(3)
End With
With txtH(4)
.Left = 4680
.Top = 5040
.Caption = RB(4)
End With
With txtH(5)
.Left = 5160
.Top = 5040
.Caption = RB(5)
End With
With txtH(6)
.Left = 5640
.Top = 5040
.Caption = RB(6)
End With
For J = 0 To 6
txtH(J).BackColor = QBColor(10)
txtC(J).Visible = False
Next
9
End Function
