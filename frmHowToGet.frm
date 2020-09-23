VERSION 4.00
Begin VB.Form FrmRuns 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Doctor Lottery"
   ClientHeight    =   2205
   ClientLeft      =   1065
   ClientTop       =   1590
   ClientWidth     =   7320
   BeginProperty Font 
      name            =   "Arial"
      charset         =   0
      weight          =   400
      size            =   11.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   2610
   Icon            =   "frmHowToGet.frx":0000
   Left            =   1005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Top             =   1245
   Width           =   7440
   Begin VB.TextBox TxtRuns 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   6360
      Picture         =   "frmHowToGet.frx":030A
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   720
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   6360
      Picture         =   "frmHowToGet.frx":128C
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   1440
      Width           =   870
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"frmHowToGet.frx":220E
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter the number of Draws :"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Execute Number of Consecutive Draws"
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
      TabIndex        =   3
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "FrmRuns"
Attribute VB_Creatable = False
Attribute VB_Exposed = False




Private Sub Form_Activate()
TxtRuns.Text = Str(IPUT)
End Sub

Private Sub Picture1_Click()
FrmRuns.Hide
End Sub


Private Sub Picture2_Click()
Dim dialogtype As Integer
Dim dialogtitle As String
Dim dialogmsg As String
Dim response As Integer
dialogtype = vbYesNo + vbExclamation
dialogtitle = "Number of Consecutive Runs"
dialogmsg = "This may take a long time, Continue?"
Dim PIX As Variant
Dim HC As Boolean
PIX = TxtRuns.Text
HC = IsNumeric(PIX)
If HC = False Then
Beep
MsgBox "You Must enter a Number!, or press Cancel.", vbCritical
GoTo 100
End If
If (Str(PIX)) > 99999999 Then
Beep
MsgBox (Str(PIX)) + " Draws will take too long!", vbCritical
PIX = 200
GoTo 100
End If
If Str(PIX) > 10000 Then
response = MsgBox(dialogmsg, dialogtype, dialogtitle)
If response = vbNo Then
GoTo 100
End If
End If
If Str(PIX) <= 0 Then
Beep
MsgBox "Negative Values are unacceptable!", vbCritica
PIX = 1
End If
IPUT = PIX
ZC.NCDR = IPUT
Put #ZA, 300, ZC
FrmRuns.Hide
100
End Sub
