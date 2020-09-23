VERSION 4.00
Begin VB.Form FrmTips 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tip Wizard"
   ClientHeight    =   4140
   ClientLeft      =   2070
   ClientTop       =   1500
   ClientWidth     =   5160
   ControlBox      =   0   'False
   Height          =   4545
   Left            =   2010
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Top             =   1155
   Width           =   5280
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1440
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2160
      Picture         =   "FrmTips.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   3720
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   120
      Picture         =   "FrmTips.frx":0292
      ScaleHeight     =   645
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblBords 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4935
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   182
      ImageHeight     =   81
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":07C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":45C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":83B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":C1B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":C454
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":C6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":C998
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":CB32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":CCCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":CE66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmTips.frx":D000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTips"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim TIP As Integer
Dim BR As Integer

Private Sub Form_Activate()
Picture3.Picture = ImageList1.ListImages(5).Picture
Randomize Timer
BR% = 33
' PiD(1) Did You Know?   PiD(3) Don't   PiD(2)  Do
TIP% = Int(Rnd * BR%) + 1
If TIP% = 1 Then GoTo 10
If TIP% = 2 Then GoTo 20
If TIP% = 3 Then GoTo 30
If TIP% = 4 Then GoTo 40
If TIP% = 5 Then GoTo 50
If TIP% = 6 Then GoTo 60
If TIP% = 7 Then GoTo 70
If TIP% = 8 Then GoTo 80
If TIP% = 9 Then GoTo 90
If TIP% = 10 Then GoTo 100
If TIP% = 11 Then GoTo 110
If TIP% = 12 Then GoTo 120
If TIP% = 13 Then GoTo 130
If TIP% = 14 Then GoTo 140
If TIP% = 15 Then GoTo 150
If TIP% = 16 Then GoTo 160
If TIP% = 17 Then GoTo 170
If TIP% = 18 Then GoTo 180
If TIP% = 19 Then GoTo 190
If TIP% = 20 Then GoTo 200
If TIP% = 21 Then GoTo 210
If TIP% = 22 Then GoTo 220
If TIP% = 23 Then GoTo 230
If TIP% = 24 Then GoTo 240
If TIP% = 25 Then GoTo 250
If TIP% = 26 Then GoTo 260
If TIP% = 27 Then GoTo 270
If TIP% = 28 Then GoTo 280
If TIP% = 29 Then GoTo 290
If TIP% = 30 Then GoTo 300
If TIP% = 31 Then GoTo 310
If TIP% = 32 Then GoTo 320
If TIP% = 33 Then GoTo 330
10 Label1.Caption = "When filling in your Lottery Game Card, do this at home, not at the last minute at your newsagents!"
Picture13.Picture = ImageList1.ListImages(2).Picture
GoTo 9999
20 Label1.Caption = "Use birthdates, or anniversaries as your lottery entries, there are only 31 days in a month and 49 Lottery numbers, you will be missing out on up to 18 Lottery numbers in your draw!"
Picture13.Picture = ImageList1.ListImages(3).Picture
GoTo 9999
30 Label1.Caption = "Use the National Lottery Game Card as a means of determining your Lottery entry, it has 4 columns of ten numbers and 1 column of nine numbers, as well as 9 rows of five numbers and 1 row of 4 numbers!"
Picture13.Picture = ImageList1.ListImages(3).Picture
GoTo 9999
40 Label1.Caption = "Use Lottery entries that have a SUM close to 150!"
Picture13.Picture = ImageList1.ListImages(2).Picture
GoTo 9999
50 Label1.Caption = "Use Lottery entries that have a RANGE close to 37!"
Picture13.Picture = ImageList1.ListImages(2).Picture
GoTo 9999
60 Label1.Caption = "You can use the Range Filter function on the Player Numbers menu to screen out ranges outside the average range of all the Lottery Draws drawn so far?"
Picture13.Picture = ImageList1.ListImages(1).Picture
Picture3.Picture = ImageList1.ListImages(4).Picture
GoTo 9999
70 Label1.Caption = "You can use the Sum Filter function on the Player Numbers menu to screen out sums outside the average sum of all the Lottery Draws drawn so far?"
Picture3.Picture = ImageList1.ListImages(6).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
80 Label1.Caption = "You can change the Lottery Prize values from the Prizes menu, Wednesday lottery prizes are generally less than Saturdays Prizes."
Picture3.Picture = ImageList1.ListImages(7).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
90 Label1.Caption = "The first Wednesday Lottery draw took place on February 5, 1997, Lottery draw was      09 25 28 29 31 35 and the bonus ball was 37."
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
100 Label1.Caption = "Use Lottery Draw entries that consist of all Odd numbers or all Even numbers, the chances of this happening in a Lottery draw is extremely remote."
Picture13.Picture = ImageList1.ListImages(3).Picture
GoTo 9999
110 Label1.Caption = "That 28 pence of every Pound you spend on the Lottery goes to good causes?"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
120 Label1.Caption = "The Government is guaranteed a win of 12 pence of every Pound you spend on the Lottery? It is called Government Duty."
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
130 Label1.Caption = "You can order all the Lottery Results in a DAT file for use with Doctor Lottery for just £14.99 Inc."
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
140 Label1.Caption = "You can use your own favourite lottery numbers by using the Fix The Players Numbers function in the Player numbers menu."
Picture3.Picture = ImageList1.ListImages(8).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
150 Label1.Caption = "If you HAVE to use random numbers as your Lottery Entry, always make sure they are random!, don't just choose them  by looking at the Game Card."
Picture3.Picture = ImageList1.ListImages(9).Picture
Picture13.Picture = ImageList1.ListImages(2).Picture
GoTo 9999
160 Label1.Caption = "Doctor Lottery Version 2.1 can actually tell you what formula has the greatest likelihood of winning the best prizes, and will show you the evidence to prove it!."
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
170 Label1.Caption = "The Biggest Lottery Jackpot win was £22.5 Million won by a lucky punter from Blackpool!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
180 Label1.Caption = "You can enter in the Most Common Occurring numbers (Hot Numbers) into the draw by using the Enter Hot Numbers function in the Lottery Numbers Menu"
Picture3.Picture = ImageList1.ListImages(10).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
190 Label1.Caption = "You can enter in the Least Common Occurring numbers (Cold Numbers) into the draw by using the Enter Cold Numbers function in the Lottery Numbers Menu"
Picture3.Picture = ImageList1.ListImages(11).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
200 Label1.Caption = "The Biggest Lottery Jackpot win was £22.5 Million won by a lucky punter from Blackpool!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
210 Label1.Caption = "Use numbers derived from Starsigns, or spiritual prediction methods, they have no scientific proof that they work, and if they did we would all be Millionaires by now!"
Picture13.Picture = ImageList1.ListImages(3).Picture
GoTo 9999
220 Label1.Caption = "£22.5 Million (the largest single Jackpot win so far) would earn £1.7 Million or just over 7 Pence per second if left in a savings account!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
230 Label1.Caption = "5% of your stake money on the Lottery goes straight to organisers Camelot?"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
240 Label1.Caption = "If entering large amounts of Lottery entries, try not to use the same prediction method for all your entries."
Picture13.Picture = ImageList1.ListImages(3).Picture
GoTo 9999
250 Label1.Caption = "80% of all Lottery winners remained anonymous, the best policy is not to tell anyone you have won!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
260 Label1.Caption = "You can change the value of each Lottery Stake by selecting 'Change the Value of Each Stake' in the Prizes Menu, - you never know when the price will go up!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
270 Label1.Caption = "You can change the value of each Lottery Stake by selecting 'Change the Value of Each Stake' in the Prizes Menu, - you never know when the price will go up!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
280 Label1.Caption = "The maximum Range possible for the National Lottery numbers is 48, the minumum is 6!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
290 Label1.Caption = "The maximum Sum possible for the National Lottery numbers is 279, the minimum is 21!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
300 Label1.Caption = "We MIGHT be getting a second National Lottery by the end of 1998, PM Tony Blair wants a 'Non Profit Lottery' bill through Parliament!"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
310 Label1.Caption = "There is a Doctor Lottery 2 coming soon, and guess what!  it will be FREE! (on Magazine Cover CD ROMS and on the Internet.)"
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
320 Label1.Caption = "Doctor Lottery can not just calculate which 6 numbers are the Most Common Occurring, but will automatically enter them into the next and subsequent draws for you!. - use the Hot Numbers function on the toolbar."
Picture3.Picture = ImageList1.ListImages(10).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
330 Label1.Caption = "Doctor Lottery can not just calculate which 6 numbers are the Least Common Occurring, but will automatically enter them into the next and subsequent draws for you!. - use the Cold Numbers function on the toolbar."
Picture3.Picture = ImageList1.ListImages(11).Picture
Picture13.Picture = ImageList1.ListImages(1).Picture
GoTo 9999
9999
End Sub
Private Sub Form_Click()
FrmTips.Hide
End Sub


Private Sub Label1_Click()
FrmTips.Hide
End Sub

Private Sub lblBords_Click()
FrmTips.Hide
End Sub

Private Sub Picture1_Click()
FrmTips.Hide
End Sub

Private Sub Picture13_Click()
FrmTips.Hide
End Sub

Private Sub Picture2_Click()
FrmTips.Hide
End Sub


