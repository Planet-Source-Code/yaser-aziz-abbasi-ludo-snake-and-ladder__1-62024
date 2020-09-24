VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWinSck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ludo Snake And Ladders Multiplayer"
   ClientHeight    =   6870
   ClientLeft      =   240
   ClientTop       =   1500
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":1272
   ScaleHeight     =   6870
   ScaleWidth      =   11880
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8160
      Top             =   5280
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   9720
      TabIndex        =   24
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtMessage 
      Height          =   375
      Left            =   8760
      TabIndex        =   23
      Top             =   5400
      Width           =   2895
   End
   Begin VB.ListBox lstMessages 
      Height          =   3960
      Left            =   8760
      TabIndex        =   22
      Top             =   1200
      Width           =   2895
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6390
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRandom 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Computer Control"
      Enabled         =   0   'False
      Height          =   420
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5205
      Width           =   1290
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7740
      Top             =   4560
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6855
      Top             =   5415
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6390
      Top             =   5430
   End
   Begin VB.CommandButton cmdPlayerColor 
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdPlayerColor 
      BackColor       =   &H00FF0000&
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdPlayerColor 
      BackColor       =   &H000080FF&
      Enabled         =   0   'False
      Height          =   345
      Index           =   3
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   765
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton cmdPlayerColor 
      BackColor       =   &H000000C0&
      Enabled         =   0   'False
      Height          =   345
      Index           =   4
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1125
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   2
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   5715
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   3
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   5715
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   4
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   5715
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   1
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   5715
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6405
      Top             =   4935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6405
      Top             =   4455
   End
   Begin VB.CommandButton cmdStartGame 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Start Game"
      Height          =   360
      Left            =   6488
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1635
   End
   Begin VB.CommandButton cmdStopDice 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Stop Dice"
      Height          =   615
      Left            =   6938
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4380
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdDiceStart 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Roll Dice"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4395
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblComments 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6405
      Width           =   6345
   End
   Begin VB.Image imgTrophy 
      Height          =   2175
      Left            =   2835
      Picture         =   "Form1.frx":855B4
      Stretch         =   -1  'True
      Top             =   2460
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   1995
      TabIndex        =   19
      Top             =   1725
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "s"
      Height          =   330
      Index           =   1
      Left            =   6885
      TabIndex        =   18
      Top             =   150
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "s"
      Height          =   330
      Index           =   2
      Left            =   6885
      TabIndex        =   17
      Top             =   510
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "s"
      Height          =   330
      Index           =   3
      Left            =   6885
      TabIndex        =   16
      Top             =   870
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblPlayerName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "s"
      Height          =   330
      Index           =   4
      Left            =   6885
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblPlayerTurn 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6525
      TabIndex        =   6
      Top             =   3150
      Width           =   1560
   End
   Begin VB.Label lblDiceResult 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Index           =   1
      Left            =   6780
      TabIndex        =   4
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lblDiceResult 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Index           =   3
      Left            =   7470
      TabIndex        =   3
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label lblDiceResult 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Index           =   2
      Left            =   7125
      TabIndex        =   2
      Top             =   3690
      Width           =   360
   End
   Begin VB.Image imgDice 
      Height          =   1080
      Index           =   1
      Left            =   6780
      Picture         =   "Form1.frx":86681
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgDice 
      Height          =   1080
      Index           =   2
      Left            =   6750
      Picture         =   "Form1.frx":88969
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgDice 
      Height          =   1080
      Index           =   3
      Left            =   6750
      Picture         =   "Form1.frx":8C194
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgDice 
      Height          =   1080
      Index           =   4
      Left            =   6750
      Picture         =   "Form1.frx":8EA70
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgDice 
      Height          =   1080
      Index           =   5
      Left            =   6750
      Picture         =   "Form1.frx":917F3
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgDice 
      Height          =   1080
      Index           =   6
      Left            =   6750
      Picture         =   "Form1.frx":946A8
      Stretch         =   -1  'True
      Top             =   1965
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intTotPlayers As Integer
Dim strPlayerName(1 To 4) As String
Dim intDiceResult(1 To 3) As Integer
Dim intRand As Integer
Dim intPlayerTurn As Integer
Dim intDiceTurn As Integer
Dim intTemp As Integer
Dim intCount As Integer
Dim intDiceTurnTemp As Integer
Dim intPlayerTurnTemp As Integer
Dim intPicBoxNo(1 To 4) As Integer
Dim strDirection As String
Dim intLeft As Integer
Dim intRight As Integer
Dim intUp As Integer
Dim intDown As Integer
Dim blnSixer(1 To 4) As Boolean
Dim blnMulti As Boolean
Dim intMultTurn As Integer
Dim strData As String
Dim blnMyTurn As Boolean

Private Sub cmdDiceStart_Click()
cmdRandom.Enabled = True
lblComments = ""
If intPlayerTurnTemp <> intPlayerTurn Then
    lblDiceResult(1) = ""
    lblDiceResult(2) = ""
    lblDiceResult(3) = ""
End If
cmdDiceStart.Visible = False
cmdStopDice.Visible = True
cmdStopDice.SetFocus
Timer1.Enabled = True
blnMyTurn = True
End Sub

Private Sub cmdRandom_Click()
If cmdRandom.Caption = "Computer Control" Then
cmdRandom.Caption = "Take Control"
Timer5.Enabled = True
Else
    cmdRandom.Caption = "Computer Control"
    Timer5.Enabled = False
End If
End Sub

Private Sub cmdSend_Click()
If txtMessage.Text = "" Then
    MsgBox "Please Enter a Message", vbInformation, "Enter Message"
Else
    Winsock1.SendData strPlayerName(1) & ": " & txtMessage.Text
    lstMessages.AddItem strPlayerName(1) & ": " & txtMessage.Text
End If
End Sub

Private Sub cmdStartGame_Click()
Winsock1.SendData "GSTART"
End Sub

Private Sub cmdStopDice_Click()
If intPlayerTurn = 1 Then
    Winsock1.SendData "INTRAN" & intRand
    'blnMyTurn = False
End If
cmdDiceStart.Visible = True
cmdDiceStart.Enabled = False
cmdStopDice.Visible = False
Timer1.Enabled = False
lblDiceResult(intDiceTurn) = intRand
intDiceResult(intDiceTurn) = intRand
intDiceTurnTemp = intDiceTurn
intPlayerTurnTemp = intPlayerTurn
If blnSixer(intPlayerTurn) Then
    If intDiceResult(intDiceTurn) + intPicBoxNo(intPlayerTurnTemp) <= 100 Then
        DiceMove
    Else
        intDiceTurn = 3
        lblComments.Caption = strPlayerName(intPlayerTurn) & ": Not moved dice going over hundered"
        If intPlayerTurn = 2 Then ' Need to check this may be put in the timer
            cmdDiceStart.Enabled = True
            cmdDiceStart.SetFocus
        End If
    End If
Else
    lblComments.Caption = strPlayerName(intPlayerTurn) & ": Not moved you didn't get a sixer yet"
    If intPlayerTurn = 2 Then ' Need to check this may be put in the timer
        cmdDiceStart.Enabled = True
        cmdDiceStart.SetFocus
    End If
End If
If intDiceResult(intDiceTurn) = 6 And blnSixer(intPlayerTurn) = False Then
    blnSixer(intPlayerTurn) = True
    'intDiceResult(intDiceTurn) = 0
End If
If Val(lblDiceResult(intDiceTurn).Caption) < 6 Or intDiceTurn = 3 Then
    Timer3.Enabled = False
    If intPlayerTurn = intTotPlayers Then
        intPlayerTurn = 1
        intDiceTurn = 1
    Else
        intPlayerTurn = intPlayerTurn + 1
        intDiceTurn = 1
    End If
    cmdPlayerColor(intPlayerTurnTemp).Visible = True
    lblPlayerTurn.Caption = strPlayerName(intPlayerTurn) & "'s Turn"
    Timer3.Enabled = True
Else
    intDiceTurn = intDiceTurn + 1
End If
End Sub

Private Sub Form_Load()
intTotPlayers = 2
strPlayerName(1) = InputBox("Enter you name")
strPlayerName(2) = "Player2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
intRand = Int(Rnd * 6) + 1
imgDice(intTemp).Visible = False
intTemp = intRand
imgDice(intRand).Visible = True
lblDiceResult(intDiceTurn) = intRand
End Sub

Private Sub DiceMove()
intCount = 1
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
intPicBoxNo(intPlayerTurnTemp) = intPicBoxNo(intPlayerTurnTemp) + 1
Select Case intPicBoxNo(intPlayerTurnTemp)
    Case 11, 21, 31, 41, 51, 61, 71, 81, 91
        picBox(intPlayerTurnTemp).Top = picBox(intPlayerTurnTemp).Top - 640
    Case 0 To 10, 22 To 30, 42 To 50, 62 To 70, 82 To 90
        picBox(intPlayerTurnTemp).Left = picBox(intPlayerTurnTemp).Left + 640
    Case 12 To 20, 32 To 40, 52 To 60, 72 To 80, 92 To 100
        picBox(intPlayerTurnTemp).Left = picBox(intPlayerTurnTemp).Left - 640
End Select
picBox(intPlayerTurnTemp).Cls
If intPicBoxNo(intPlayerTurnTemp) < 10 Then
    picBox(intPlayerTurnTemp).Print " " & intPicBoxNo(intPlayerTurnTemp)
Else
    picBox(intPlayerTurnTemp).Print "" & intPicBoxNo(intPlayerTurnTemp)
End If
If intCount = intDiceResult(intDiceTurnTemp) Then
    Select Case intPicBoxNo(intPlayerTurnTemp)
        Case 15, 32, 35, 38, 58, 65, 81, 90, 94, 98
            Snake
        Case 4, 8, 22, 25, 30, 44, 49, 60, 72, 85
            Ladder
        Case Else
            If intPlayerTurn = 1 Then
                cmdDiceStart.Enabled = True
                cmdDiceStart.SetFocus
            End If
    End Select
    If intPicBoxNo(intPlayerTurnTemp) = 100 Then
        lblWinner.Caption = strPlayerName(intPlayerTurnTemp) & " Wins"
        lblWinner.Visible = True
        imgTrophy.Visible = True
        Timer1.Enabled = False
        cmdDiceStart.Enabled = False
    End If
    Timer2.Enabled = False
End If
intCount = intCount + 1
End Sub

Private Sub Timer3_Timer()
If cmdPlayerColor(intPlayerTurn).Visible Then
    cmdPlayerColor(intPlayerTurn).Visible = False
Else
    cmdPlayerColor(intPlayerTurn).Visible = True
End If
End Sub

Private Sub Snake()
Select Case intPicBoxNo(intPlayerTurnTemp)
    Case 15
        intDown = 1
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 5
    Case 32
        intDown = 3
        intPicBoxNo(intPlayerTurnTemp) = 9
    Case 35
        intDown = 2
        intLeft = 3
        intPicBoxNo(intPlayerTurnTemp) = 18
    Case 38
        intDown = 2
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 19
    Case 58
        intDown = 1
        intLeft = 2
        intPicBoxNo(intPlayerTurnTemp) = 41
    Case 65
        intDown = 4
        intRight = 1
        intPicBoxNo(intPlayerTurnTemp) = 26
    Case 81
        intDown = 2
        intRight = 2
        intPicBoxNo(intPlayerTurnTemp) = 63
    Case 90
        intDown = 2
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 69
    Case 94
        intDown = 6
        intRight = 1
        intPicBoxNo(intPlayerTurnTemp) = 33
    Case 98
        intDown = 3
        intRight = 1
        intPicBoxNo(intPlayerTurnTemp) = 64
End Select
lblComments.Caption = strPlayerName(intPlayerTurn) & ": Oops you just got bitten by a snake"
picBox(intPlayerTurnTemp).Cls
Timer4.Enabled = True
End Sub
Private Sub Ladder()
Select Case intPicBoxNo(intPlayerTurnTemp)
    Case 4
        intUp = 2
        intPicBoxNo(intPlayerTurnTemp) = 24
    Case 8
        intUp = 3
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 34
    Case 22
        intUp = 2
        intPicBoxNo(intPlayerTurnTemp) = 42
    Case 25
        intUp = 5
        intRight = 3
        intPicBoxNo(intPlayerTurnTemp) = 42
    Case 30
        intUp = 1
        intPicBoxNo(intPlayerTurnTemp) = 31
    Case 44
        intUp = 3
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 78
    Case 49
        intUp = 1
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 53
    Case 60
        intUp = 1
        intPicBoxNo(intPlayerTurnTemp) = 61
    Case 72
        intUp = 2
        intLeft = 1
        intPicBoxNo(intPlayerTurnTemp) = 93
    Case 85
        intUp = 1
        intPicBoxNo(intPlayerTurnTemp) = 96
End Select
lblComments.Caption = strPlayerName(intPlayerTurn) & ": You just climbed a ladder "
picBox(intPlayerTurnTemp).Cls
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
If intUp > 0 Then
    intUp = intUp - 1
    picBox(intPlayerTurnTemp).Top = picBox(intPlayerTurnTemp).Top - 640
End If
If intDown > 0 Then
    intDown = intDown - 1
    picBox(intPlayerTurnTemp).Top = picBox(intPlayerTurnTemp).Top + 640
End If
If intLeft > 0 Then
    intLeft = intLeft - 1
    picBox(intPlayerTurnTemp).Left = picBox(intPlayerTurnTemp).Left - 640
End If
If intRight > 0 Then
    intRight = intRight - 1
    picBox(intPlayerTurnTemp).Left = picBox(intPlayerTurnTemp).Left + 640
End If
If intUp = 0 And intDown = 0 And intLeft = 0 And intRight = 0 Then
    picBox(intPlayerTurnTemp).Cls
    If intPicBoxNo(intPlayerTurnTemp) < 10 Then
        picBox(intPlayerTurnTemp).Print " " & intPicBoxNo(intPlayerTurnTemp)
    Else
        picBox(intPlayerTurnTemp).Print "" & intPicBoxNo(intPlayerTurnTemp)
    End If
    If intPlayerTurn = 1 Then
        cmdDiceStart.Enabled = True
        cmdDiceStart.SetFocus
    End If
    Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
Dim intRand As Integer
intRand = Int(Rnd * 3000) + 500
Timer5.Interval = intRand
If cmdDiceStart.Enabled = True And cmdDiceStart.Visible = True Then
    cmdDiceStart_Click
Else
    If cmdStopDice.Enabled = True And cmdStopDice.Visible = True Then
    cmdStopDice_Click
    End If
End If
If picBox(intPlayerTurnTemp) = 100 Then
Timer5.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
Winsock1.SendData "NMDATA" & strPlayerName(1)
Timer6.Enabled = False
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
Winsock1.SendData "CONTON"
Form5.Hide
Form1.Show
lstMessages.AddItem "Client connected"
Timer6.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData strData
Select Case Left(strData, 6)
Case "GSTART"
If MsgBox("Would you like to play Snake and Ladders with " & strPlayerName(2), vbYesNo, "Multiplayer Game") = vbYes Then
    Winsock1.SendData "GREADY"
    intPlayerTurn = 2
    GameStart
    cmdDiceStart.Visible = True
Else
    Winsock1.SendData "NOTRED"
End If
Case "GREADY"
    MsgBox "Game Started Your Turn"
    intPlayerTurn = 1
    GameStart
    cmdDiceStart.Enabled = True
    cmdDiceStart.Visible = True
    cmdDiceStart.SetFocus
Case "NOTRED"
    MsgBox "Game Invitation Refused"
Case "INTRAN"
    intRand = Mid(strData, 7, 1)
    cmdStopDice_Click
Case "CONTON"
    Form5.Hide
    Form1.Show
    lstMessages.AddItem "You are connected"
    Winsock1.SendData "NMDATA" & strPlayerName(1)
Case "NMDATA"
    strPlayerName(2) = Mid(strData, 7, Len(strData))
Case Else
lstMessages.AddItem strData
End Select
End Sub
   
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
MsgBox Description, vbCritical, "Error"
End Sub


Public Sub GameStart()
Dim intCount As Integer
For intCount = 1 To intTotPlayers
    'strPlayerName(intCount) = Form3.txtPlayer(intCount)
    picBox(intCount).Visible = True
    picBox(intCount).Left = picBox(intCount).Left - 640
    lblPlayerName(intCount).Visible = True
    lblPlayerName(intCount).Caption = strPlayerName(intCount)
    cmdPlayerColor(intCount).Visible = True
Next intCount
cmdStartGame.Visible = False
intTemp = 1
intDiceTurn = 1
lblPlayerTurn.Caption = strPlayerName(intPlayerTurn) & "'s Turn"
Timer3.Enabled = True
End Sub
