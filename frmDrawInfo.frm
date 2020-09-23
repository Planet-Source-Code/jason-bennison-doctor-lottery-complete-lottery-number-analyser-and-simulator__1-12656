VERSION 4.00
Begin VB.Form frmPrizes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doctor Lottery  Custom Settings"
   ClientHeight    =   4140
   ClientLeft      =   705
   ClientTop       =   1455
   ClientWidth     =   8310
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
   Height          =   4545
   Icon            =   "frmDrawInfo.frx":0000
   Left            =   645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Top             =   1110
   Width           =   8430
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   6100
      Picture         =   "frmDrawInfo.frx":030A
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   19
      Top             =   3360
      Width           =   870
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use default settings"
      Height          =   495
      Left            =   3840
      TabIndex        =   18
      Top             =   3480
      Width           =   2055
   End
   Begin VB.OptionButton OPtion1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use these settings from now on"
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   3480
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.Frame FrmSettings 
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
      Left            =   480
      TabIndex        =   16
      Top             =   3240
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   6960
      Picture         =   "frmDrawInfo.frx":128C
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   15
      Top             =   3360
      Width           =   870
   End
   Begin VB.TextBox Text 
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
      Index           =   6
      Left            =   6480
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   0
      Left            =   6480
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   5
      Left            =   6480
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   4
      Left            =   6480
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   3
      Left            =   6480
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   2
      Left            =   6480
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text 
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
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change the value of the Lottery Prizes"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   18
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 6 Lottery Ball Numbers:"
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 5 + Bonus Ball Number:"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 5 Lottery Ball Numbers:"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 4 Lottery Ball Numbers:"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 3 Lottery Ball Numbers:"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 2 Lottery Ball Numbers:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prize Money to be awarded for matching 1 Lottery Ball Number:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   5895
   End
End
Attribute VB_Name = "frmPrizes"
Attribute VB_Creatable = False
Attribute VB_Exposed = False





Private Sub Check1_Click()
Check1.Value = 1

End Sub


Private Sub Check2_Click()
Check2.Value = 1
End Sub

Private Sub Form_Activate()
OPtion1.Value = True
For J = 0 To 6
Text(J).Text = "£" + Str(PRIZE(J))
Next
End Sub

Private Sub OPtion1_Click()
OPtion1.Value = True
End Sub

Private Sub Option2_Click()
Option2.Value = True
PRIZE(1) = 0
PRIZE(2) = 0
PRIZE(3) = 10
PRIZE(4) = 70
PRIZE(5) = 1500
PRIZE(6) = 8000000
PRIZE(0) = 100000
For J = 0 To 6
Text(J).Text = "£" + Str(PRIZE(J))
Next

End Sub

Private Sub Picture1_Click()
Dim PIX(0 To 6) As Variant
Dim HC(0 To 6) As Boolean
For J = 0 To 6
PIX(J) = Text(J).Text
Next
For J = 0 To 6
HC(J) = IsNumeric(PIX(J))
If HC(J) = False Then
If (J) = 0 Then
Beep
MsgBox "Prize for 5 + Bonus Ball Match is not Monetary!", vbCritical
GoTo 100
End If
Beep
MsgBox "Prize for " + Str(J) + " Ball Match is not Monetary!", vbCritical
GoTo 100
End If
Next
If (Str(PIX(1))) > 99999999 Then
Beep
MsgBox "£" + (Str(PIX(1))) + " For matching just 1 ball? - No Way!", vbCritical
PIX(1) = 0
GoTo 100
End If
If (Str(PIX(2))) > 99999999 Then
Beep
MsgBox "£" + (Str(PIX(2))) + " For matching just 2 balls? - No Way!", vbCritical
PIX(2) = 0
GoTo 100
End If
If (Str(PIX(3))) > 99999999 Then
Beep
MsgBox "£" + (Str(PIX(3))) + " For matching 3 balls? - No Way!", vbCritical
PIX(3) = 10
GoTo 100
End If
If (Str(PIX(4))) > 99999999 Then
Beep
MsgBox "£" + (Str(PIX(4))) + " For matching 4 balls? - No Way!", vbCritical
PIX(4) = 70
GoTo 100
End If
If (Str(PIX(5))) > 99999999 Then
Beep
MsgBox "£" + (Str(PIX(5))) + " For matching 5 balls? - No Way!", vbCritical
PIX(5) = 1500
GoTo 100
End If
If (Str(PIX(6))) > 999999999 Then
Beep
MsgBox "£" + (Str(PIX(6))) + " For matching all 6 balls? - No Way!", vbCritical
PIX(6) = 8000000
GoTo 100
End If
If "£" + (Str(PIX(0))) > 99999999 Then
Beep
MsgBox "£" + (Str(PIX(0))) + " For matching 5 + Bonus Ball? - No Way!", vbCritical
PIX(0) = 100000
GoTo 100
End If
For J = 0 To 6
If PIX(J) < 0 Then
Beep
MsgBox "Negative Values are unacceptable!", vbCritical
GoTo 100
End If
Next
For J = 0 To 6
PRIZE(J) = Str(PIX(J))
Next
ZC.Priz1 = PRIZE(1)
ZC.Priz2 = PRIZE(2)
ZC.Priz3 = PRIZE(3)
ZC.Priz4 = PRIZE(4)
ZC.Priz5 = PRIZE(5)
ZC.Priz6 = PRIZE(6)
ZC.Priz0 = PRIZE(0)
Put #ZA, 300, ZC
frmPrizes.Hide
100
End Sub

Private Sub Picture2_Click()
frmPrizes.Hide
End Sub

Private Sub Text_Click(Index As Integer)
OPtion1.Value = True
End Sub


