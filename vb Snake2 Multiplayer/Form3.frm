VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Snakes And Ladders"
   ClientHeight    =   3330
   ClientLeft      =   1560
   ClientTop       =   1260
   ClientWidth     =   1500
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   1500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPlayers 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   4125
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   2850
      Begin VB.CommandButton cmdSubmitPlayers 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3000
         Width           =   1485
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   2745
         Width           =   1485
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   2070
         Width           =   1485
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   1425
         Width           =   1485
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Enter The Players Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   0
         TabIndex        =   10
         Top             =   15
         Width           =   1485
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Player 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   8
         Top             =   2430
         Width           =   1485
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Player 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   1755
         Width           =   1485
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   1110
         Width           =   1485
      End
      Begin VB.Label lblPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   1485
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSubmitPlayers_Click()
Dim blnInCorrect As Boolean
Dim intCount As Integer
For intCount = 1 To Form1.intTotPlayers
    If txtPlayer(intCount) = "" Then
        MsgBox "You must enter the players name to continue", vbInformation, "Player Names"
        txtPlayer(intCount).SetFocus
        blnInCorrect = True
        Exit For
    End If
Next intCount
If blnInCorrect = False Then
    Form3.Hide
    SinglePlayer.Show
End If
End Sub

Private Sub Form_Load()
Dim strTempPlayers As String
Dim blnCorrect As Boolean
Dim intCount As Integer
Do Until blnCorrect
    strTempPlayers = InputBox("Enter The Total Number Of Players", "Number of Players")
    If IsNumeric(strTempPlayers) = True And Val(strTempPlayers) < 5 And Val(strTempPlayers) > 1 Then
        Form1.intTotPlayers = Val(strTempPlayers)
        blnCorrect = True
    Else
        MsgBox "!Invalid number, Please Enter a valid number from 2 to 4", vbInformation, "Number of Players"
    End If
Loop
For intCount = 1 To Form1.intTotPlayers
    txtPlayer(intCount).Enabled = True
Next intCount
End Sub


Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to exit", vbYesNo) = vbNo Then
    Cancel = 1
Else
    End
End If
End Sub
