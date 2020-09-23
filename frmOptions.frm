VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Options"
   ClientHeight    =   1950
   ClientLeft      =   3210
   ClientTop       =   2775
   ClientWidth     =   3840
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   480
      Left            =   45
      TabIndex        =   2
      Top             =   1410
      Width           =   3735
   End
   Begin VB.CommandButton cmdPrvChat 
      Caption         =   "Hava A Private Chat"
      Height          =   480
      Left            =   45
      TabIndex        =   1
      Top             =   860
      Width           =   3735
   End
   Begin VB.CommandButton cmdPrvMsg 
      Caption         =   "Send A Private Message"
      Height          =   480
      Left            =   45
      TabIndex        =   0
      Top             =   330
      Width           =   3735
   End
   Begin VB.Label lblUser 
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   45
      Width           =   3735
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdPrvChat_Click()
    frmPrivateMessage.cChat.SendMessage sUsername, lblUser, Chr(1) & "prvchat"
End Sub

Private Sub cmdPrvMsg_Click()
    Dim sTo As String
    Dim sMessage As String
    
    sTo = lblUser
    sMessage = InputBox("Message to send to the user", "Message")
    If sMessage = "" Then Exit Sub
    frmPrivateMessage.SendPrivateMessage sTo, sMessage
End Sub

Private Sub Form_Load()
    Center Me
End Sub
