VERSION 4.00
Begin VB.Form frmAbout 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Lottery  (Freeware Version 1.1)"
   ClientHeight    =   2925
   ClientLeft      =   1440
   ClientTop       =   1335
   ClientWidth     =   7140
   Height          =   3330
   Icon            =   "frmaboutLot.frx":0000
   Left            =   1380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Top             =   990
   Width           =   7260
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmaboutLot.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Jay.Soft@USA.Net"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2475
      TabIndex        =   6
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   1455
      Left            =   720
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "Programming by:              Jason Bennison"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Proof Reading by:             Belinda Bennison"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Technical Consultant:     Samuel Bennison"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   5445
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Published by JaySoft, 52 Hemsby Road, Chessington, Surrey KT9 2DZ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   5685
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Doctor Lottery 1.1   for Windows"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   720
      X2              =   5160
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

End Sub


Private Sub Form_Click()
frmAbout.Hide
frmLottery.Show
End Sub


Private Sub Form_Load()
frmAbout.Left = (Screen.Width / 2) - (frmAbout.Width / 2)
frmAbout.Top = (Screen.Height / 2) - (frmAbout.Height / 2)
End Sub


Private Sub Label4_Click()
frmAbout.Hide
frmLottery.Show
End Sub

Private Sub Label5_Click()
frmAbout.Hide
frmLottery.Show
End Sub

Private Sub Label6_Click()
frmAbout.Hide
frmLottery.Show
End Sub


Private Sub Picture1_Click()
frmAbout.Hide
frmLottery.Show
End Sub


