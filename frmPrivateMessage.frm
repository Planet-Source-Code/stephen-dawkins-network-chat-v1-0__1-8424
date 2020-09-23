VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrivateMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Private Message"
   ClientHeight    =   4035
   ClientLeft      =   1530
   ClientTop       =   1980
   ClientWidth     =   8520
   Icon            =   "frmPrivateMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lstMsg 
      Height          =   3495
      Left            =   105
      TabIndex        =   4
      Top             =   90
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Message"
         Object.Width           =   11959
      EndProperty
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Delete All"
      Height          =   345
      Left            =   4440
      TabIndex        =   3
      Top             =   3660
      Width           =   1410
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Message"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2985
      TabIndex        =   2
      Top             =   3660
      Width           =   1410
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "Reply"
      Height          =   345
      Left            =   1545
      TabIndex        =   1
      Top             =   3660
      Width           =   1410
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Message"
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   3660
      Width           =   1410
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8040
      Top             =   3690
   End
End
Attribute VB_Name = "frmPrivateMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cChat As New clsChat
Private iMsg As Integer

Private Sub cmdClearAll_Click()
    lstMsg.ListItems.Clear
End Sub

Private Sub cmdDelete_Click()
    lstMsg.ListItems.Remove lstMsg.SelectedItem.Index
End Sub

Private Sub cmdReply_Click()
    Dim sTo As String
    Dim sMessage As String
    
    sTo = lstMsg.SelectedItem.Text
    If sTo = "" Then Exit Sub
    sMessage = InputBox("Message to send to the user", "Message")
    If sMessage = "" Then Exit Sub
    cChat.SendMessage sUsername, sTo, sMessage
End Sub

Private Sub cmdSend_Click()
    Dim sTo As String
    Dim sMessage As String
    
    sTo = InputBox("User to send the message to.", "User")
    If sTo = "" Then Exit Sub
    sMessage = InputBox("Message to send to the user", "Message")
    If sMessage = "" Then Exit Sub
    cChat.SendMessage sUsername, sTo, sMessage
End Sub

Private Sub Form_Activate()
    iMsg = 0
    frmMenu.cmdPrivateMessage.Caption = "Private Messages"
End Sub

Private Sub Form_Load()
    cChat.Init "prvmsg"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Hide
End Sub

Private Sub lblFun_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Dim sTo As String
    Dim sMessage As String
    Dim sFrom As String
    
    sTo = InputBox("User to send the message to.", "User")
    If sTo = "" Then Exit Sub
    sFrom = InputBox("User to send the message to be sent from.", "User")
    If sFrom = "" Then Exit Sub
    sMessage = InputBox("Message to send to the user", "Message")
    If sMessage = "" Then Exit Sub
    SendMessageToWinPopUp sFrom, sTo, sMessage
End Sub

Private Sub lstMsg_ItemClick(ByVal Item As ListItem)
    cmdDelete.Enabled = True
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim sMsg() As String
    Dim z() As String
    Dim itm As ListItem
    
    sMsg() = cChat.GetMessages
    
    If UBound(sMsg) < 0 Then Exit Sub
    On Error GoTo 0
    For i = 0 To UBound(sMsg)
        tmp = sMsg(i)
        z() = Split(tmp, Chr(0))
        Select Case z(1)
        Case Chr(1) & "prvchat"
            If MsgBox("Do you want a private chat with '" & z(0) & "'.", vbYesNo) = vbYes Then
                CreatePrivateChat z(0), False
                cChat.SendMessage sUsername, z(0), Chr(1) & "prvchatyes"
            Else
                cChat.SendMessage sUsername, z(0), Chr(1) & "prvchatno"
            End If
        Case Chr(1) & "prvchatyes"
            CreatePrivateChat z(0), True
        Case Chr(1) & "prvchatno"
            MsgBox z(0) & " refused private chat."
        Case Else
            Set itm = lstMsg.ListItems.Add(1, , z(0))
            itm.SubItems(1) = z(1)
            If Me.Visible = False Then
                iMsg = iMsg + UBound(sMsg) + 1
                frmMenu.cmdPrivateMessage.Caption = "Private Messages (" & iMsg & ")"
            End If
        End Select
    Next
End Sub

Public Sub PrvChat(sUser As String)
    cChat.SendMessage sUsername, sUser, Chr(1) & "prvchat"
End Sub

Public Sub SendPrivateMessage(sTo As String, sMessage As String)
    cChat.SendMessage sUsername, sTo, sMessage
End Sub
