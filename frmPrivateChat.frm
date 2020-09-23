VERSION 5.00
Begin VB.Form frmPrivateChat 
   Caption         =   "Private Chat With ''"
   ClientHeight    =   3540
   ClientLeft      =   1485
   ClientTop       =   2160
   ClientWidth     =   7260
   Icon            =   "frmPrivateChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7260
   Tag             =   "stephen"
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   210
      Left            =   -1000
      TabIndex        =   2
      Top             =   2685
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2385
      Top             =   2775
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   2115
      Width           =   6510
   End
   Begin VB.TextBox txtMsg 
      Height          =   2040
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   30
      Width           =   6510
   End
End
Attribute VB_Name = "frmPrivateChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cChat As New clsChat

Private Sub cmdSend_Click()
    If txtChat > "" Then
        cChat.SendMessage sUsername, "*", txtChat
        txtChat = ""
    End If
End Sub

Private Sub Form_Resize()
    txtChat.Move 30, Me.ScaleHeight - txtChat.Height - 30, Me.ScaleWidth - 60, txtChat.Height
    txtMsg.Move 30, 30, Me.ScaleWidth - 60, txtChat.Top - 60
End Sub

Public Sub Go(sUser As String, bMeInit As Boolean)
    Dim sTmp As String
    
    Me.Tag = sUser
    Me.Caption = "Private Chat With '" & sUser & "'"
    If bMeInit Then
        sTmp = sUsername & sUser
    Else
        sTmp = sUser & sUsername
    End If
    cChat.Init "prvchat\" & sTmp
    If cChat.GetError = 1 Then
        MsgBox "Unable to create mailslot."
        Unload Me
    End If
    Center Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cChat.CloseHnd
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim sMsg() As String
    Dim z() As String
    
    sMsg() = cChat.GetMessages
    
    If UBound(sMsg) < 0 Then Exit Sub
    On Error GoTo 0
    For i = 0 To UBound(sMsg)
        z() = Split(sMsg(i), Chr(0))
        txtMsg = txtMsg & "<" & z(0) & ">" & z(1) & vbNewLine
        txtMsg.SelStart = Len(txtMsg)
    Next
End Sub

