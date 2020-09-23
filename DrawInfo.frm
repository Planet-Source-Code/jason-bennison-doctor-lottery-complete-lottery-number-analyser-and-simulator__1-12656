VERSION 4.00
Begin VB.Form FrmFixed 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doctor Lottery"
   ClientHeight    =   2400
   ClientLeft      =   1335
   ClientTop       =   1545
   ClientWidth     =   6615
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   2805
   Icon            =   "DrawInfo.frx":0000
   Left            =   1275
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Top             =   1200
   Width           =   6735
   Begin VB.TextBox Text 
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
      Index           =   6
      Left            =   4320
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text 
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
      Index           =   5
      Left            =   3600
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text 
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
      Index           =   4
      Left            =   2880
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text 
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
      Index           =   3
      Left            =   2160
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text 
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
      Index           =   2
      Left            =   1440
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text 
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
      Index           =   1
      Left            =   720
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5640
      Picture         =   "DrawInfo.frx":030A
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   1680
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5640
      Picture         =   "DrawInfo.frx":128C
      ScaleHeight     =   720
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter your own 'Favourite 6' Numbers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   $"DrawInfo.frx":220E
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5295
   End
End
Attribute VB_Name = "FrmFixed"
Attribute VB_Creatable = False
Attribute VB_Exposed = False





Private Sub Form_Activate()
FrmFixed.Left = (Screen.Width / 2) - (FrmFixed.Width / 2)
FrmFixed.Top = (Screen.Height / 2) - (FrmFixed.Height / 2)
DIC = 0
If SF(1) = 0 Then
SF(1) = 1: SF(2) = 2: SF(3) = 3: SF(4) = 4: SF(5) = 5: SF(6) = 6
End If
Text(1).Text = SF(1)
Text(2).Text = SF(2)
Text(3).Text = SF(3)
Text(4).Text = SF(4)
Text(5).Text = SF(5)
Text(6).Text = SF(6)
End Sub

Private Sub Picture1_Click()
DIC = 1
FrmFixed.Hide
End Sub


Private Sub Picture2_Click()
Dim PIX(1 To 6) As Variant
Dim HC(1 To 6) As Boolean
For J = 1 To 6
PIX(J) = Text(J).Text
Next
For J = 1 To 6
HC(J) = IsNumeric(PIX(J))
If HC(J) = False Then
Beep
MsgBox "Ball " + (Str(J)) + " is not valid!", vbCritical
GoTo 1000
End If
Next
For J = 1 To 6
If (Str(PIX(J))) > 49 Then
Beep
MsgBox "Ball " + (Str(J)) + " is greater than 49. Lottery Numbers range from 1 to 49", vbCritical
GoTo 1000
End If
If (Str(PIX(J))) <= 0 Then
Beep
MsgBox "Ball " + (Str(PIX(J))) + " is less than 1. Lottery Numbers range from 1 to 49", vbCritical
GoTo 1000
End If
Next
For J = 1 To 6
SF(J) = PIX(J)
Next
For J = 1 To 6
If SF(J) = SF(1) Then
If J = 1 Then GoTo 10
MsgBox "Numbers 1 and " + (Str(J)) + " are the same"
GoTo 1000
10
End If
Next
For J = 1 To 6
If SF(J) = SF(2) Then
If J = 2 Then GoTo 20
MsgBox "Numbers 2 and " + (Str(J)) + " are the same"
GoTo 1000
20
End If
Next
For J = 1 To 6
If SF(J) = SF(3) Then
If J = 3 Then GoTo 30
MsgBox "Numbers 3 and " + (Str(J)) + " are the same"
GoTo 1000
30
End If
Next
For J = 1 To 6
If SF(J) = SF(4) Then
If J = 4 Then GoTo 40
MsgBox "Numbers 4 and " + (Str(J)) + " are the same"
GoTo 1000
40
End If
Next
For J = 1 To 6
If SF(J) = SF(5) Then
If J = 5 Then GoTo 50
MsgBox "Numbers 5 and " + (Str(J)) + " are the same"
GoTo 1000
50
End If
Next
For J = 1 To 6
If SF(J) = SF(6) Then
If J = 6 Then GoTo 60
MsgBox "Numbers 6 and " + (Str(J)) + " are the same"
GoTo 1000
60
End If
Next
ZC.PL1 = SF(1)
ZC.PL2 = SF(2)
ZC.PL3 = SF(3)
ZC.PL4 = SF(4)
ZC.PL5 = SF(5)
ZC.PL6 = SF(6)
Put #ZA, 300, ZC
FrmFixed.Hide
1000
End Sub


