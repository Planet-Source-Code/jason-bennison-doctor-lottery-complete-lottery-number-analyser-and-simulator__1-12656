VERSION 4.00
Begin VB.Form frmRangeFilter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Range Filter"
   ClientHeight    =   3165
   ClientLeft      =   1485
   ClientTop       =   2700
   ClientWidth     =   6840
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
   Icon            =   "frmRange.frx":0000
   Left            =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Top             =   2355
   Width           =   6960
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Restore default settings"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   2400
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use these settings from now on"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Value           =   -1  'True
      Width           =   3855
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
      Left            =   3120
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   8
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
      Left            =   3120
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5880
      Picture         =   "frmRange.frx":030A
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   6
      Top             =   1440
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5880
      Picture         =   "frmRange.frx":128C
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   5
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Max. 48)"
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
      Left            =   4200
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Range Filter Settings"
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
      TabIndex        =   4
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lower Range Limit:"
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
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Upper Range Limit:"
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
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Min. 6)"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmRangeFilter"
Attribute VB_Creatable = False
Attribute VB_Exposed = False







Private Sub Form_Activate()
Text1.Text = Str(LRL)
Text2.Text = Str(URL)
End Sub

Private Sub OPtion1_Click()
OPtion1.Value = True
End Sub

Private Sub Option2_Click()
Option2.Value = True
LRL = 10
URL = 48
Text1.Text = Str(LRL)
Text2.Text = Str(URL)
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
PIX(1) = 10
PIX(2) = 20
Dim HC As Boolean
If Text1.Text = "" Then Text1.Text = "10"
If Text2.Text = "" Then Text2.Text = "40"
PIX(1) = Str(Text1.Text)
HC = IsNumeric(PIX(1))
If HC = False Then
Beep
MsgBox "The Lower Range Limit must be a number!", vbCritical
GoTo 100
End If
If (Str(PIX(1))) < 6 Then
Beep
MsgBox "Lower Range Limit must be greater than 6.", vbCritical
PIX(1) = 10
GoTo 100
End If
If (Str(PIX(1))) >= 48 Then
Beep
MsgBox "Lower Range Limit exceeds any possible Upper Range Limit!", vbCritica
PIX(1) = 6
GoTo 100
End If
PIX(2) = Str(Trim(Text2.Text))
HC = IsNumeric(PIX(2))
If HC = False Then
Beep
MsgBox "The Upper Range Limit must be a number!", vbCritical
GoTo 100
End If
If (Str(PIX(2))) <= 6 Then
Beep
MsgBox "Upper Range Limit exceeds any possible Lower Range Limit.", vbCritical
PIX(2) = 40
GoTo 100
End If
If Str(PIX(2)) > 48 Then
Beep
MsgBox "Upper Range Limit must not be greater than 48.", vbCritical
PIX(2) = 40
GoTo 100
End If
If PIX(2) <= 15 Then
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
GoTo 90
End If
If PIX(1) >= 46 Then
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
End If
90
If PIX(1) <= 9 Then
MsgBox "The Lower Range Limit is too low and will have no effect."
PIX(1) = 10
Text1.Text = Str(PIX(1))
LRL = PIX(1)
GoTo 100
End If
If PIX(1) > PIX(2) Then
MsgBox "The Lower Range Limit (" + PIX(1) + ") is greater than the Upper Range Limit (" + PIX(2) + ")"
LRL = 10
URL = 45
GoTo 100
End If
LRL = PIX(1)
URL = PIX(2)
ZC.UPRANL = URL
ZC.LORANL = LRL
Put #ZA, 300, ZC
frmRangeFilter.Hide
100
End Sub
Private Sub Picture2_Click()
DIC = 1
frmRangeFilter.Hide
End Sub


