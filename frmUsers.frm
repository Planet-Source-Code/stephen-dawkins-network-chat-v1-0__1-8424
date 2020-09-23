VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Online"
   ClientHeight    =   3480
   ClientLeft      =   3285
   ClientTop       =   2415
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3375
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3345
      Top             =   1260
   End
   Begin VB.ListBox lstUsers 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3240
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    frmPrivateChat.cChat.SendMessage sUsername, "home", Chr(1) & "ping"
End Sub
