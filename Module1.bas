Attribute VB_Name = "Module1"
Public J As Integer
Public FE As Long
Public AV As Integer
Public FA As Integer
Public FB As Long
Public FD As Long
Public HPP As Integer
Public USL As Integer
Public LSL As Integer
Public URL As Integer
Public LRL As Integer
Public DIC As Integer
Public LsetAct As Integer
Public LsetRan As Integer
Public LB(0 To 6) As Integer
Public LS(0 To 6) As Integer
Public P(1 To 49) As Integer
Public RecNum
Public SF(1 To 6) As Integer
Public TRB(1 To 6) As Long
Public RB(1 To 6) As Long
Public IPUT As Long
Public FC As DrawsInfo
Public gPrevLine1
Public PRIZE(0 To 6) As Currency
Public Stake As Currency
Public ZE
Public ZA
Public ZD
Public ZC As PrizeInfo
Public ZZ As Integer
Type DrawsInfo
Ball1 As Integer
Ball2 As Integer
Ball3 As Integer
Ball4 As Integer
ball5 As Integer
ball6 As Integer
BallBon As Integer
Date As String * 31
End Type
Type PrizeInfo
Priz1 As Currency
Priz2 As Currency
Priz3 As Currency
Priz4 As Currency
Priz5 As Currency
Priz0 As Currency
Priz6 As Currency
Stak As Currency
PL1 As Integer
PL2 As Integer
PL3 As Integer
PL4 As Integer
PL5 As Integer
PL6 As Integer
ZZ1 As Integer
UPSUML As Integer
LOSUML As Integer
NCDR As Long
UPRANL As Integer
LORANL As Integer
End Type

