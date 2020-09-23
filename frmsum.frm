VERSION 4.00
Begin VB.Form FrmSumFilter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sum Filter"
   ClientHeight    =   3165
   ClientLeft      =   1470
   ClientTop       =   1410
   ClientWidth     =   6855
   BeginProperty Font 
      name            =   "Arial"
      charset         =   0
      weight          =   400
      size            =   11.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   3570
   Icon            =   "frmsum.frx":0000
   Left            =   1410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Top             =   1065
   Width           =   6975
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use these settings from now on"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Value           =   -1  'True
      Width           =   3495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restore default settings"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Screen out all Player Numbers:"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   6615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5880
      Picture         =   "frmsum.frx":030A
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   1440
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5880
      Picture         =   "frmsum.frx":128C
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Min.  21)"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Max. 279)"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1500
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Upper Sum Limit:"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lower Sum Limit:"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   14.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sum Filter Settings"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   21.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "FrmSumFilter"
Attribute VB_Creatable = False
Attribute VB_Exposed = False






Private Sub Form_Activate()
Text1.Text = Str(LSL)
Text2.Text = Str(USL)
End Sub

Private Sub OPtion1_Click()
OPtion1.Value = True
LSL = 100
USL = 200
Text1.Text = Str(LSL)
Text2.Text = Str(USL)
End Sub

Private Sub Option2_Click()
Option2.Value = True

End Sub

Private Sub Picture1_Click()
Dim dialogtype As Integer
Dim dialogtitle As String
Dim dialogmsg As String
Dim response As Integer
dialogtype = vbYesNo + vbExclamation
dialogtitle = "Engine"
dialogmsg = "This may take a long time, Continue?"
Dim PIX(1 To 2) As Variant
Dim HC As Boolean
If Text1.Text = "" Then Text1.Text = "100"
If Text2.Text = "" Then Text2.Text = "200"
PIX(1) = Str(Text1.Text)
HC = IsNumeric(PIX(1))
If HC = False Then
Beep
MsgBox "The Lower Sum Limit must be a number!", vbCritical
GoTo 100
End If
If (Str(PIX(1))) <= 20 Then
Beep
MsgBox "Lower Sum Limit must be greater than 21.", vbCritical
PIX(1) = 21
GoTo 100
End If
If Str(PIX(1)) >= 279 Then
Beep
MsgBox "Lower Sum Limit exceeds any possible Upper Sum Limit!", vbCritica
PIX(1) = 21
GoTo 100
End If
PIX(2) = Str(Trim(Text2.Text))
HC = IsNumeric(PIX(2))
If HC = False Then
Beep
MsgBox "The Upper Sum Limit must be a number!", vbCritical
GoTo 100
End If
If (Str(PIX(2))) <= 20 Then
Beep
MsgBox "Upper Sum Limit exceeds any possible Lower Sum Limit.", vbCritical
PIX(2) = 21
GoTo 100
End If
If Str(PIX(2)) > 279 Then
Beep
MsgBox "Upper Sum Limit must not be greater than 279.", vbCritical
PIX(2) = 21
GoTo 100
End If
If PIX(2) <= 52 Then
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
90
End If
If PIX(1) >= 230 Then
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
End If
If PIX(1) <= 99 Then
MsgBox "The Lower Sum Limit is below 100 and will have no effect"
PIX(1) = 100
GoTo 100
End If
If PIX(1) >= PIX(2) Then
MsgBox "The Lower Sum Limit (" + PIX(1) + ") is Greater than the Upper Sum Limit (" + PIX(2) + ")"
GoTo 100
End If
LSL = PIX(1)
USL = PIX(2)
ZC.UPSUML = USL
ZC.LOSUML = LSL
Put #ZA, 300, ZC
FrmSumFilter.Hide
100
End Sub
Private Sub Picture2_Click()
DIC = 1
FrmSumFilter.Hide
End Sub


