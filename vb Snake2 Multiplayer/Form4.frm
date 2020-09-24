VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Choose Game Type"
   ClientHeight    =   2970
   ClientLeft      =   2925
   ClientTop       =   1830
   ClientWidth     =   6930
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMultiplayer 
      BackColor       =   &H00008080&
      Caption         =   "Multiple Player (LAN)"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1830
      Width           =   1695
   End
   Begin VB.CommandButton cmdLocalPlayer 
      BackColor       =   &H00008080&
      Caption         =   "Local Player"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1830
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to play single player or multiple player over the Iternet or Lan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   960
      Left            =   330
      TabIndex        =   2
      Top             =   525
      Width           =   6270
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLocalPlayer_Click()
Form4.Hide
Form3.Show
End Sub

Private Sub cmdMultiplayer_Click()
Form4.Hide
Form5.Show
End Sub
