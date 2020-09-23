VERSION 4.00
Begin VB.Form frmorders 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Getting the Doctor Lottery Lottery Results Data file..."
   ClientHeight    =   5445
   ClientLeft      =   1065
   ClientTop       =   1005
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Height          =   5850
   Icon            =   "frmAdjustPrizes.frx":0000
   Left            =   1005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Top             =   660
   Width           =   7620
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   2520
      Picture         =   "frmAdjustPrizes.frx":0442
      ScaleHeight     =   1320
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   4080
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   1680
      Picture         =   "frmAdjustPrizes.frx":2464
      ScaleHeight     =   1050
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   120
      Width           =   3915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Price includes Post + Packing"
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   5160
      Width           =   2115
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Order the Lottery Results DAT file for just £14.99"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "This copy of Doctor Lottery does not need to be registered."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAdjustPrizes.frx":70D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   6135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 1.1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   1050
   End
End
Attribute VB_Name = "frmorders"
Attribute VB_Creatable = False
Attribute VB_Exposed = False






Private Sub Form_Load()
frmorders.Left = (Screen.Width / 2) - (frmorders.Width / 2)
frmorders.Top = (Screen.Height / 2) - (frmorders.Height / 2)
End Sub


Private Sub Label1_Click()
frmorders.Hide
End Sub

Private Sub Picture2_Click()
Close
frmorders.Hide
End
End Sub


