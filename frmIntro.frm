VERSION 4.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Lottery  (c)  JaySoft "
   ClientHeight    =   6180
   ClientLeft      =   495
   ClientTop       =   480
   ClientWidth     =   8730
   ControlBox      =   0   'False
   Height          =   6585
   Icon            =   "frmIntro.frx":0000
   Left            =   435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Top             =   135
   Width           =   8850
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   5160
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   7
      Top             =   4800
      Width           =   2655
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1200
      Picture         =   "frmIntro.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   3210
      TabIndex        =   5
      Top             =   5280
      Width           =   3210
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   2400
      Picture         =   "frmIntro.frx":14CC
      ScaleHeight     =   1050
      ScaleWidth      =   3915
      TabIndex        =   4
      Top             =   345
      Width           =   3915
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   177
      ImageHeight     =   88
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmIntro.frx":613E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmIntro.frx":8170
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "This FREEWARE program has no operating restrictions"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   5175
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   327682
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmIntro.frx":A1A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   3960
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmIntro.frx":A285
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   7455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmIntro.frx":A394
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   7335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version 1.1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   0
      Top             =   960
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      FillColor       =   &H00FF0000&
      Height          =   1215
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Form_Load()
frmIntro.Left = (Screen.Width / 2) - (frmIntro.Width / 2)
frmIntro.Top = (Screen.Height / 2) - (frmIntro.Height / 2)
Picture1.Picture = ImageList2.ListImages(1).Picture
End Sub


Private Sub Picture1_Click()
Picture1.Picture = ImageList2.ListImages(2).Picture
frmLottery.Show
frmIntro.Hide
End Sub


