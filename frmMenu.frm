VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   2115
   ClientLeft      =   4065
   ClientTop       =   3855
   ClientWidth     =   2865
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2865
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   3840
      Picture         =   "frmMenu.frx":0442
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1425
      Width           =   300
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   15
      TabIndex        =   4
      Top             =   1695
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrivateMessage 
      Caption         =   "Private Messages"
      Height          =   375
      Left            =   15
      TabIndex        =   3
      Top             =   1275
      Width           =   2835
   End
   Begin VB.CommandButton cmdPrivateChat 
      Caption         =   "Private Chat"
      Height          =   375
      Left            =   15
      TabIndex        =   2
      Top             =   861
      Width           =   2835
   End
   Begin VB.CommandButton cmdJoinMain 
      Caption         =   "Join Main Chat Room"
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2835
   End
   Begin VB.CommandButton cmdJoinRoom 
      Caption         =   "Join Room"
      Height          =   375
      Left            =   15
      TabIndex        =   1
      Top             =   438
      Width           =   2835
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then
        Unload Me
    Else
        sUsername = InputBox("Username?")
    End If
End Sub

Private Sub cmdJoinMain_Click()
    CreateRoomChat "chat"
End Sub

Private Sub cmdJoinRoom_Click()
    Dim sRoom As String
try:
    sRoom = InputBox("Enter name of room to join. If the room doesn't exist then it will be created.", "Join Room")
    If sRoom = "" Then Exit Sub
    If Len(sRoom) > 8 Then
        MsgBox "Length of room must 8 or less letters long"
        GoTo try:
    End If
    sRoom = "userroom\" & sRoom
    CreateRoomChat sRoom
End Sub

Private Sub cmdPrivateChat_Click()
    Dim sUser As String
    
    sUser = InputBox("Enter the user you want to have a private chat with.")
    If sUser = "" Then Exit Sub
    frmPrivateMessage.PrvChat sUser
End Sub

Private Sub cmdPrivateMessage_Click()
    frmPrivateMessage.Show
End Sub

Private Sub cmdUsers_Click()
    frmUsers.Show
End Sub

Private Sub Form_Load()
    sUsername = Space(100)
    ret = GetComputerName(sUsername, 100)
    sUsername = Mid(sUsername, 1, InStr(1, sUsername, Chr(0)) - 1)
    CreateRoomChat "chat"
    Load frmPrivateMessage
    frmPrivateMessage.Hide
    Me.Icon = picMain.Picture
    Me.Move 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadAll
    End
End Sub
