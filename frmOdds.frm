VERSION 4.00
Begin VB.Form FrmQuickHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Doctor Lottery Quick Help"
   ClientHeight    =   3375
   ClientLeft      =   885
   ClientTop       =   1710
   ClientWidth     =   8055
   BeginProperty Font 
      name            =   "Arial"
      charset         =   0
      weight          =   400
      size            =   11.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   3780
   Icon            =   "frmOdds.frx":0000
   Left            =   825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Top             =   1365
   Width           =   8175
   Begin VB.Label Main 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   0
      X2              =   0
      Y1              =   375
      Y2              =   3350
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   8040
      X2              =   8040
      Y1              =   375
      Y2              =   3350
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   8040
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Header 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         name            =   "Arial"
         charset         =   0
         weight          =   700
         size            =   12
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   8040
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "FrmQuickHelp"
Attribute VB_Creatable = False
Attribute VB_Exposed = False






Private Sub Form_Activate()
If HPP = 0 Then
Header.Caption = "Object not found"
Main.Caption = "You right-Clicked an object or item on the screen that does not have Help."
End If
If HPP = 1 Then
Header.Caption = "Sum Filter"
Main.Caption = "This can only be active when the Players numbers being entered against the Lottery Draw are set to Random. Doctor Lottery automatically calculates the Aggregate Total or Sum for each lottery draw that have taken place since you last pressed Restart or opened Doctor Lottery. This function works with both, Actual Lottery Draw results and with Randomly generated (Simulated) Lottery Results.  For more information on Sum Filter, see the felp file."
End If
If HPP = 2 Then
Header.Caption = "System Restart"
Main.Caption = "This resets Doctor Lottery, sets all values to zero, prepares the next draw for entering against your next Lottery Entry. System Restart has no effect on the Toolbar, Menu settings or the content of the stored Lottery results in the results Data File."
End If
If HPP = 3 Then
Header.Caption = "Random Lottery Numbers"
Main.Caption = "Initiates the Random Number Generator, produces a set of 6 random numbers and a Bonus Ball number ranging between 1 and 49 with no duplication. Doctor Lottery enters these numbers as if they are lottery Draw Numbers. Unlike the number of Actual Lottery Draws, which will always be finite, there is no limit to the number of times a set of randomly generated Lottery Numbers that can be produced consecutively. This function is effectively a Lottery 'Simulator'"
End If
If HPP = 4 Then
Header.Caption = "Actual Lottery Numbers"
Main.Caption = "Doctor Lottery will call for the Lottery Draw Results that are stored in the file 'Results.DAT' which holds all the result data for all the National Lottery draws taken so far. When selected, it will display the date of the first draw in the Draw Date field located above the Lottery Numbers in readiness for entering against your own chosen number sequence. The file 'Results.DAT' is not supplies with the freeware edition of Doctor Lottery, but can be ordered from JaySoft. See Help Contents."
End If
If HPP = 5 Then
Header.Caption = "View Lottery Results"
Main.Caption = "Displays the Lottery Draws Dialogue box, allowing you to view, edit or add the latest draw result from the Lottery. To view the Lottery Draw Results, Doctor Lottery requires the results data file 'Results.DAT' to be present prior to displaying them."
End If
If HPP = 6 Then
Header.Caption = "Enter HOT Numbers"
Main.Caption = "Doctor Lottery automatically calculates the number of times each lottery ball number has appeared (whether using random or actual draw results), the program then evaluates which numbers have appeared most often and will enter the 6 most 'hottest' or Most Common Occurring numbers as your next lottery entry. This function determines the theory that if a particular number keeps appearing in the Lottery Draw result, it may be prudent to use that number again for subsequent or future draws."
End If
If HPP = 7 Then
Header.Caption = "Enter COLD Numbers"
Main.Caption = "Doctor Lottery automatically calculates the number of times each lottery ball number has appeared (whether using random or actual draw results), the program then evaluates which numbers have appeared least often and will enter the 6  'coldest' or Least Common Occurring numbers as your next lottery entry. This function determines the theory that if a particular number has not appeared in the Lottery Draw result for a while, it could be described as being 'Overdue', it may be prudent to use that number for subsequent or future draws."
End If
If HPP = 8 Then
Header.Caption = "Enter RANDOM Numbers"
Main.Caption = "Doctor Lottery will initiate the Random Number Generator and produce a set of 6 RANDOM numbers ranging from 1 to 49 with no duplication. Doctor Lottery will then enter these numbers as if they are your Lottery Draw Entries. These numbers can be edited using tools such as Range Filter or Sum Filter etc.  There is no limit to the number of randomly generated Lottery Entries you can make."
End If
If HPP = 9 Then
Header.Caption = "Enter FIXED or PRE-SET Numbers"
Main.Caption = "Doctor Lottery displays the fixed player numbers dialogue box inviting you to enter your own 6 numbers to be entered and placed against the lottery draw Results. Duplicate numbers are not allowed."
End If
If HPP = 10 Then
Header.Caption = "Range Filter"
Main.Caption = "This can only be active when the Players numbers being entered against the Lottery Draw are set to Random. Doctor Lottery automatically calculates the Range for each lottery draw that have taken place since you last pressed Restart or opened Doctor Lottery. This function works with both, Actual Lottery Draw results and with Randomly generated (Simulated) Lottery Results.  For more information on Range Filter, see the felp file."
End If
If HPP = 11 Then
Header.Caption = "Insert Prize Values"
Main.Caption = "Displays the Prize Values Dialogue box. This enables you to set your own prize values for matching any number of Lottery Balls."
End If
If HPP = 12 Then
Header.Caption = "Tip Wizard"
Main.Caption = "Displays the Tip Wizard Dialogue Box with some snippets about the Lottery in general."
End If
If HPP = 13 Then
Header.Caption = "Help Contents"
Main.Caption = "Displays the Doctor Lottery Contents page of the Help file."
End If
If HPP = 14 Then
Header.Caption = "Display Current Lottery Numbers"
Main.Caption = "This is set to ON by default. it shows the currently displayed Lottery Draw Numbers onto the Gamecard by highlighting the ball number fields in white."
End If
If HPP = 15 Then
Header.Caption = "Display Current Players Numbers"
Main.Caption = "This shows the currently entered Player Numbers onto the Gamecard by highlighting the ball number fields in yellow."
End If
If HPP = 16 Then
Header.Caption = "Display HOT Numbers"
Main.Caption = "This shows the current 6 Most Common Occurring numbers of all the Lottery Draws so far on to the Gamecard by highlighting the ball number fields in red."
End If
If HPP = 17 Then
Header.Caption = "Display COLD Numbers"
Main.Caption = "This shows the current 6 Least Common Occurring numbers of all the Lottery Draws so far on to the Gamecard by highlighting the ball number fields in blue."
End If
If HPP = 18 Then
Header.Caption = "Exit"
Main.Caption = "Shows some advertising and away you go."
End If
If HPP = 19 Then
Header.Caption = "Select Number of Draws"
Main.Caption = "This push button displays the Number of Draws dialogue box. This enables you to pre-set the number of draws you want to run consecutively when the RUN button located beneath the toolbar - is pressed. A field located to the right displays the current setting"
End If
If HPP = 20 Then
Header.Caption = "Date of Actual Lottery Draw"
Main.Caption = "This is only active when you are running Actual Lottery Draw Results, it displays the date the currently displayed Lottery Draw was made by the Lottery. For running Actual Lottery Results, you need the data file 'Results.DAT' containing all the draw results and dates from the National Lottery which can be ordered from Jaysoft."
End If
If HPP = 21 Then
Header.Caption = "Continuous Running Display"
Main.Caption = "Doctor Lottery will perform Lottery Draws and submit Players Lottery entries using your chosen prediction method continuously and displaying the result of each draw. When in use, a vertical slider control will appear to the right, this allows the user control of the speed of each draw being displayed for visual representation."
End If
If HPP = 22 Then
Header.Caption = "Continuous cycle (RUN)"
Main.Caption = "This control is used in conjunction with the 'Select Number of Draws' button located to the right. Doctor Lottery will perform continuously that number of draws shown in the Select Number of Draws Field, the program will display the collective result. There is no maximum number of draws you can run, however the more draws you do, the longer it will take, speed is dependent on the speed of your Maths Co-processor."
End If
If HPP = 23 Then
Header.Caption = "1 Draw Cycle (STEP)"
Main.Caption = "Doctor Lottery will perform 1 draw and enter 1 set of players numbers against it, indicate which and how many numbers matched by highlighing the number fields in red and display the result."
End If
If HPP = 24 Then
Header.Caption = "Total spent on the Lottery"
Main.Caption = "This represents the total amount that has been spent on Lottery Draw Stakes since opening Doctor Lottery or pressing 'Restart'"
End If
If HPP = 25 Then
Header.Caption = "Total Winnings so far"
Main.Caption = "This represents the total amount that has been 'Won Back' from the Lottery since opening Doctor Lottery or pressing 'Restart'.  If you can get this figure to be greater than the 'Total Spent on the Lottery' figure, you have made a PROFIT. "
End If
If HPP = 26 Then
Header.Caption = "% Stake Money Lost/Won back"
Main.Caption = "This represents the total amount as a percentage that has been paid out in Lottery Stakes and have NOT won back from the Lottery since opening Doctor Lottery or last pressing 'Restart'.There is a Bar Chart at the base of the screen marked '% Stakes won back', this represents a visual aid to determing how much you have lost on the Lottery. If the Bar Chart 'Fills' you have made a profit."
End If
If HPP = 27 Then
Header.Caption = "% Stake Money Won Back"
Main.Caption = "This represents the total amount as a percentage that has been won back from the stake money paid out to the Lottery since opening Doctor Lottery, or last pressing 'Restart'."
End If
If HPP = 28 Then
Header.Caption = "Sum:"
Main.Caption = "This is the Sum, or Aggregate Total of the currently displayed Lottery Numbers."
End If
If HPP = 29 Then
Header.Caption = "Range:"
Main.Caption = "This is the Range of the currently displayed Lottery Draw Numbers."
End If
If HPP = 30 Then
Header.Caption = "Number of Draws:"
Main.Caption = "This represents the number of Lottery Draws, Actual or Simulated, that have been performed since opening Doctor Lottery or last pressing 'Restart'."
End If
If HPP = 31 Then
Header.Caption = "Odds of winning - 1 in:"
Main.Caption = "This represents the chances of winning ANY Lottery Prize regardless of value, and comparing this with the number of Lottery Entries you have made since opening Doctor Lottery or last pressing 'Restart'.  The Odds figure is calculated by taking the number of cash prizes that have been won, dividing this with the number of Lottery Entries that have been placed, multiplied by 100."
End If
If HPP = 33 Then
Header.Caption = "Sum:"
Main.Caption = "Only Visible when the Sum Filter is active. This is the Sum, or Aggregate Total of the currently displayed Player Entry Numbers."
End If
If HPP = 32 Then
Header.Caption = "Ran (Range)"
Main.Caption = "Only visible when the Range Filter is active. This is the Range of the currently displayed Player Entry Numbers."
End If
If HPP = 34 Then
Header.Caption = "Lottery Numbers"
Main.Caption = "These are the currently displyed Lottery Numbers, where a field is highlighted in red, this shows a match with a player entry number, the matching player number is also highlighted. To the left of the Lottery Numbers display is the method that the Lottery Numbers are being produced or obtained."
End If
If HPP = 35 Then
Header.Caption = "Player Numbers"
Main.Caption = "These are the currently displayed Player Entry Numbers, where a field is highlighted in red, this shows a match with a Lottery Ball Number, the matching Lottery Ball number is also highlighted. To the left of the Player Numbers display is the method the Player Numbers are being produced or obtained."
End If
If HPP = 36 Then
Header.Caption = "Number of Prizes that have been won."
Main.Caption = "These are number of times any number of Lottery Ball Numbers have matched a Players Entry number. To the right of this is the currently set prize values awarded."
End If
If HPP = 37 Then
Header.Caption = "Average Range"
Main.Caption = "The Average Range is calculated by taking all the Ranges of all the Lottery Draws, Actual or simulated, adding them together then dividing that figure by the number of draws performed."
End If
If HPP = 38 Then
Header.Caption = "Highest Range"
Main.Caption = "This is the highest occurring Range the Lottery results (Simulated or Actual) that has been recorded since opening Doctor Lottery or pressing 'Restart'."
End If
If HPP = 39 Then
Header.Caption = "Lowest Range"
Main.Caption = "This is the lowest occurring Range the Lottery result (simulated or actual) that has been recorded since opening Doctor Lottery or pressing 'Restart'."
End If
If HPP = 40 Then
Header.Caption = "Average Sum"
Main.Caption = "The Average Sum (Aggregate Total) is calculated by taking all the Sums of all the Lottery Draws, Actual or simulated, adding them together then dividing that figure by the number of draws performed."
End If
If HPP = 41 Then
Header.Caption = "Highest Sum"
Main.Caption = "This is the highest occurring Sum the Lottery result (Simulated or Actual) that has been recorded since opening Doctor Lottery or pressing 'Restart'."
End If
If HPP = 42 Then
Header.Caption = "Lowest Sum"
Main.Caption = "This is the lowest occurring Sum the Lottery result (Simulated or Actual) that has been recorded since opening Doctor Lottery or pressing 'Restart'."
End If
If HPP = 43 Then
If AV = 0 Then
Header.Caption = "6 Most Common Occurring Numbers"
Main.Caption = "These are the numbers that have appeared as Lottery Ball Numbers (Actual or Simulated) most frequently. They read from left to right starting with the Most Common Occurring Number first. For more information on what you can do with Most Common Occurring Numbers, see the help file."
End If
If AV = 1 Then
Header.Caption = "Average 6"
Main.Caption = "These numbers are the Average 6 numbers of the Lottery draws taken so far (actual or simulated). If you take the first number of every lottery draw, add them together and divide them by the number of draws, you get an average for that ball. The process is repeated for the rest of the Lottery balls and then entered in to the Lottery draw as your next entry. For more detail on Average 6, Please see 'Prediction Theories' in the help file."
End If
End If
If HPP = 44 Then
Header.Caption = "6 Least Common Occurring Numbers"
Main.Caption = "These are the numbers that have appeared as Lottery Ball Numbers (Actual or Simulated) least frequently. They read from left to right starting with the Least Common Occurring Number first. For more information on what you can do with Least Common Occurring Numbers, see the help file."
End If
If HPP = 45 Then
Header.Caption = "Progress..."
Main.Caption = "This is only active when you are running Lottery Draws consecutively using the 'RUN' button. It is a visual aid to determining at a glance how long the process is taking."
End If
If HPP = 46 Then
Header.Caption = "Game Card"
Main.Caption = "The Game card occupies most of the right hand side of the screen when Doctor Lottery is running. It is a visual aid to determining at a glance which numbers are currently being played. It is set out similar to a Lottery Game card. It consists of 49 fields, each with a number between 1 and 49 above it, by default, inside the field displays the number of times that number has appeared in a Lottery Result, you can change this to display the Percentage of times each number has appeared in a draw from the Gamecard Menu."
End If
If HPP = 100 Then
Header.Caption = "Getting Doctor Lottery Quick Help"
Main.Caption = "Simply Right-Click anything you want help on, and this box will appear with information about the object you clicked on. For More in-depth information, see Help Contents."
End If
HPP = 0
End Sub

Private Sub Form_Click()
FrmQuickHelp.Hide
End Sub

Private Sub Header_Click()
HPP = 0
FrmQuickHelp.Hide
End Sub

Private Sub Main_Click()
HPP = 0
FrmQuickHelp.Hide
End Sub


Private Sub Text1_Click()
HPP = 0
FrmQuickHelp.Hide
End Sub


