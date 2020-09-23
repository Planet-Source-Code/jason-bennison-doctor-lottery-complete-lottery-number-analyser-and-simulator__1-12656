VERSION 4.00
Begin VB.Form frmSuggestions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Was it You?........"
   ClientHeight    =   5280
   ClientLeft      =   735
   ClientTop       =   585
   ClientWidth     =   8340
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Height          =   5685
   Icon            =   "frmSuggestions.frx":0000
   Left            =   675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Top             =   240
   Width           =   8460
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   2880
      Picture         =   "frmSuggestions.frx":0442
      ScaleHeight     =   1320
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSuggestions.frx":2464
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSuggestions.frx":2508
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   7815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSuggestions.frx":26A8
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   7695
   End
   Begin VB.Label lblSuggest 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Any Suggestions?  - Tell me about them at the address below."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   7920
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmSuggestions"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
frmSuggestions.Hide
End Sub

Private Sub Form_Click()
frmSuggestions.Hide
End Sub

Private Sub Form_DblClick()
frmSuggestions.Hide
End Sub

Private Sub Label1_Click()
frmSuggestions.Hide
End Sub

Private Sub Label3_Click()
frmSuggestions.Hide
End Sub


Private Sub lblSuggest_Click()
frmSuggestions.Hide
End Sub

Private Sub Picture1_Click()
frmSuggestions.Hide
End Sub


