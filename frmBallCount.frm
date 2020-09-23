VERSION 4.00
Begin VB.Form FrmStake 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Doctor Lottery  Custom Settings"
   ClientHeight    =   2295
   ClientLeft      =   960
   ClientTop       =   1605
   ClientWidth     =   7560
   ClipControls    =   0   'False
   BeginProperty Font 
      name            =   "Arial"
      charset         =   0
      weight          =   400
      size            =   9.75
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2700
   Icon            =   "frmBallCount.frx":0000
   Left            =   900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Top             =   1260
   Width           =   7680
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6600
      Picture         =   "frmBallCount.frx":030A
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   7
      Top             =   720
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6600
      Picture         =   "frmBallCount.frx":128C
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   6
      Top             =   1440
      Width           =   870
   End
   Begin VB.TextBox TxtStake 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restore Default setting"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   400
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use this setting from now on"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   400
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Settings"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Price payable for each Lottery Stake:"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   885
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change the Value of each Lottery Stake"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmStake"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub Form_Activate()
Opt1.Value = True
TxtStake.Text = "£" + Str(Stake)
End Sub

Private Sub Form_Load()
'lblIN.Caption = "The Right hand side of the Screen respresents a National Lottery Game Card, each number in the box shows how many times each ball has shown in a simulated draw."
'lblINT.Caption = "The Most Common Occouring numbers are highlighted in RED, and the Least Common Occourring Numbers are in BLUE. If you select - Enter 'HOT' Numbers,-  the program will calculate which numbers are the Most Common Occourring after each draw, and will enter them in the subsequent draw as the Players Numbers"
'lblNt2.Caption = "The 'Count Ball numbers?' push button will disable counting the number of times each lottery number has come up, and in turn disable the 'Enter HOT numbers' and 'Enter COLD numbers' functions."
End Sub







Private Sub Opt1_Click()
Opt1.Value = True
End Sub


Private Sub Opt2_Click()
Opt2.Value = True
Stake = 1
TxtStake.Text = "£" + Str(Stake)
End Sub


Private Sub Picture1_Click()
FrmStake.Hide
End Sub

Private Sub Picture2_Click()
Dim PIX As Variant
Dim HC As Boolean
PIX = TxtStake.Text
HC = IsNumeric(PIX)
If HC = False Then
Beep
MsgBox "The Lottery Stake MUST be a monetary Value!", vbCritical
GoTo 100
End If
If (Str(PIX)) > 9999 Then
Beep
MsgBox "£" + (Str(PIX)) + " for a Lottery Stake is too much!", vbCritical
PIX = 1
GoTo 100
End If
If Str(PIX) <= 0 Then
Beep
MsgBox "Negative or Non Values are unacceptable!", vbCritica
PIX = 1
GoTo 100
End If
Stake = PIX
ZC.Stak = Stake
Put #ZA, 300, ZC
FrmStake.Hide
100
End Sub




