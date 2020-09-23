VERSION 4.00
Begin VB.Form frmExit 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prize Money History"
   ClientHeight    =   5430
   ClientLeft      =   315
   ClientTop       =   915
   ClientWidth     =   9030
   BeginProperty Font 
      name            =   "Arial"
      charset         =   0
      weight          =   400
      size            =   11.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   5835
   Icon            =   "frmExit.frx":0000
   Left            =   255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Top             =   570
   Width           =   9150
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   48
      Text            =   "£"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   45
      Text            =   "£"
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   6840
      TabIndex        =   43
      Text            =   "£"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   42
      Text            =   "£"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   41
      Text            =   "£"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   40
      Text            =   "£"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   39
      Text            =   "£ "
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Amend"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   23
      Text            =   "£"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Text            =   "£"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Text            =   "£"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Text            =   "£"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Text            =   "£"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4335
      Left            =   8640
      TabIndex        =   18
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   17
      Text            =   "£ 2945989"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   16
      Text            =   "£ 118233"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Index           =   2
      Left            =   3240
      TabIndex        =   15
      Text            =   "£ 2463"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   14
      Text            =   "£ 85"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   12
      Text            =   "£ 10"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Text            =   "3"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Text            =   "23"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Text            =   "690"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Text            =   "43623"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Text            =   "948598"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "E&xit"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total estimated Lottery takings to-date:"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Lottery winnings paid out to-date:"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   46
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total to-date"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   44
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bonus"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   38
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   1        2        3        4        5        6"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   37
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   36
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8400
      TabIndex        =   34
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7680
      TabIndex        =   33
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7080
      TabIndex        =   32
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6480
      TabIndex        =   31
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   30
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   29
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   28
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wednesday, September 25, 1997"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   27
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "194"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Average to-date"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Awarded"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4 Ball Match"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3 Ball Match:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5 Ball Match"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5 + Bonus"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Jackpot"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Winning Tickets"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
frmExit.Hide
End Sub





