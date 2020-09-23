VERSION 4.00
Begin VB.Form FrmSpent 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profit & Loss Result...."
   ClientHeight    =   5265
   ClientLeft      =   750
   ClientTop       =   1080
   ClientWidth     =   8190
   ControlBox      =   0   'False
   Height          =   5670
   Icon            =   "FrmSpent.frx":0000
   Left            =   690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Top             =   735
   Width           =   8310
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   8055
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"FrmSpent.frx":030A
      Height          =   855
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"FrmSpent.frx":03B9
      Height          =   855
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"FrmSpent.frx":0474
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   3700
      Width           =   3615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "This figure represents the amount of stake money that has been placed into the lottery that has been lost as a percentage."
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"FrmSpent.frx":053A
      Height          =   975
      Left            =   4320
      TabIndex        =   14
      Top             =   2060
      Width           =   3615
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"FrmSpent.frx":061D
      Height          =   975
      Left            =   4320
      TabIndex        =   13
      Top             =   3700
      Width           =   3615
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   1920
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line5 
      X1              =   4320
      X2              =   5520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      X1              =   4320
      X2              =   6120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   4320
      X2              =   7560
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   2760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Spent on the Lottery"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Winnings so far"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Odds of winning:"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "% Stake lost"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "% Stake won back"
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date of draw  (Actual draws only) "
      BeginProperty Font 
         name            =   "System"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   4200
      TabIndex        =   4
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   4200
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "FrmSpent"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
FrmSpent.Hide
End Sub


