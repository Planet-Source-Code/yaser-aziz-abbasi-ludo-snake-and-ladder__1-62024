VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Snakes And Ladders"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10200
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1350
      Top             =   345
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   -120
      Picture         =   "Form2.frx":1272
      Top             =   -45
      Width           =   10425
   End
   Begin VB.Image Image2 
      Height          =   1440
      Left            =   -150
      Picture         =   "Form2.frx":389B
      Top             =   1380
      Width           =   10485
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Timer1_Timer()
Form2.Hide
Form4.Show
Timer1.Enabled = False
End Sub
