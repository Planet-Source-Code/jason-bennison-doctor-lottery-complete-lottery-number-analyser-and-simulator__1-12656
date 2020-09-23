VERSION 4.00
Begin VB.Form frmColdNumbers 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cold Numbers"
   ClientHeight    =   1590
   ClientLeft      =   1530
   ClientTop       =   1845
   ClientWidth     =   6735
   Height          =   1995
   Icon            =   "Form1.frx":0000
   Left            =   1470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Top             =   1500
   Width           =   6855
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":030A
      Top             =   0
      Width           =   6735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmColdNumbers"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
frmColdNumbers.Hide
End Sub


Private Sub Label2_Click()

End Sub


Private Sub Text1_Change()
frmColdNumbers.Hide
End Sub


