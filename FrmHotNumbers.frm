VERSION 4.00
Begin VB.Form FrmHotNumbers 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hot Numbers"
   ClientHeight    =   1695
   ClientLeft      =   1545
   ClientTop       =   1845
   ClientWidth     =   6645
   Height          =   2100
   Icon            =   "FrmHotNumbers.frx":0000
   Left            =   1485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Top             =   1500
   Width           =   6765
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmHotNumbers.frx":030A
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "FrmHotNumbers"
Attribute VB_Creatable = False
Attribute VB_Exposed = False





Private Sub Text1_Click()
FrmHotNumbers.Hide
End Sub


