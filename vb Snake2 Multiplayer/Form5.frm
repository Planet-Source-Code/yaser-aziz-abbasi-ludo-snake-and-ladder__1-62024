VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000040&
   Caption         =   "Multiplayer"
   ClientHeight    =   3390
   ClientLeft      =   3750
   ClientTop       =   3225
   ClientWidth     =   2160
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   2160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15
      Top             =   120
   End
   Begin VB.TextBox txtNotice 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   225
      TabIndex        =   7
      Top             =   2850
      Width           =   1800
   End
   Begin VB.CommandButton cmdConnect 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Caption         =   "Connect "
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
      Left            =   360
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1485
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Text            =   "1981"
      Top             =   1860
      Width           =   1485
   End
   Begin VB.TextBox txtIp 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   1245
      Width           =   1485
   End
   Begin VB.CommandButton cmdHost 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Caption         =   "Host Game"
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
      Left            =   360
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Port"
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
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1485
   End
   Begin VB.Label lblIp 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Ip Adress"
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
      Left            =   360
      TabIndex        =   3
      Top             =   945
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Enter The Remote Ip and Port"
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
      Height          =   690
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1485
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnLstn As Boolean
Dim intCount As Integer
Private Sub cmdConnect_Click()
Form1.Winsock1.Connect txtIp, txtPort
cmdHost.Enabled = False
cmdConnect.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub cmdHost_Click()
Form1.Winsock1.LocalPort = txtPort
Form1.Winsock1.Listen
cmdHost.Enabled = False
cmdConnect.Enabled = False
txtIp.Enabled = False
txtPort.Enabled = False
blnLstn = True
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Form1.Show
Form1.Hide
txtIp = Form1.Winsock1.LocalIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
If blnLstn Then
    Select Case intCount
        Case 1
            txtNotice.Text = "Listening    "
        Case 2
            txtNotice.Text = "Listening .  "
        Case 3
            txtNotice.Text = "Listening .. "
        Case 4
            txtNotice.Text = "Listening ..."
    End Select
Else
    Select Case intCount
        Case 1
            txtNotice.Text = "Connecting    "
        Case 2
            txtNotice.Text = "Connecting .  "
        Case 3
            txtNotice.Text = "Connecting .. "
        Case 4
            txtNotice.Text = "Connecting ..."
    End Select
End If
intCount = intCount + 1
If intCount = 5 Then
    intCount = 1
End If
End Sub
