VERSION 5.00
Begin VB.Form frmRoom 
   Caption         =   " "
   ClientHeight    =   2895
   ClientLeft      =   1710
   ClientTop       =   2520
   ClientWidth     =   8775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   8775
   Begin VB.ListBox lstUsr 
      Height          =   2400
      Left            =   6900
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   1815
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   195
      Left            =   -1000
      TabIndex        =   2
      Top             =   4635
      Width           =   390
   End
   Begin VB.TextBox txtChat 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   2160
      Width           =   6810
   End
   Begin VB.TextBox txtMsg 
      Height          =   2070
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   45
      Width           =   6795
   End
   Begin VB.Timer tmrMain 
      Interval        =   1000
      Left            =   7890
      Top             =   2610
   End
End
Attribute VB_Name = "frmRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cChat As New clsChat

Private Sub cmdSend_Click()
    If txtChat > "" Then
        cChat.SendMessage sUsername, "*", txtChat
        txtChat = ""
        tmrMain_Timer
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstUsr.Move Me.ScaleWidth - lstUsr.Width, 45, lstUsr.Width, Me.ScaleHeight - 50
    txtChat.Move 30, Me.ScaleHeight - txtChat.Height, Me.ScaleWidth - lstUsr.Width - 75
    txtMsg.Move 30, 45, txtChat.Width, txtChat.Top - 75
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cChat.SendMessage sUsername, "*", Chr(1) & "bye"
    cChat.CloseHnd
End Sub

Private Sub lstUsr_DblClick()
    Dim sTo As String
    Dim sMessage As String
    
    'If sUsername = lstUsr.List(lstUsr.ListIndex) Then Exit Sub
    frmOptions.lblUser = lstUsr.List(lstUsr.ListIndex)
    frmOptions.Show
End Sub

Private Sub tmrMain_Timer()
    On Error Resume Next
    Dim sTmp() As String
    Dim tmp() As String

    sTmp() = cChat.GetMessages
    
    If UBound(sTmp) < 0 Then
        Exit Sub
    End If
    'On Error GoTo 0
    For i = 0 To UBound(sTmp)
        tmp() = Split(sTmp(i), Chr(0))
        
        Select Case tmp(1)
        Case Chr(1) & "userlist"
            cChat.SendMessage sUsername, tmp(0), Chr(1) & "username"
            AddUsr tmp(0)
        Case Chr(1) & "username"
            AddUsr tmp(0)
        Case Chr(1) & "prvchat"
            If MsgBox("Do you want a private chat with '" & tmp(0) & "'.", vbYesNo) = vbYes Then
                CreatePrivateChat tmp(0), False
                cChat.SendMessage sUsername, tmp(0), Chr(1) & "pvrchatyes"
            Else
                cChat.SendMessage sUsername, tmp(0), Chr(1) & "pvrchatno"
            End If
        Case Chr(1) & "pvrchatyes"
            CreatePrivateChat tmp(0), True
        Case Chr(1) & "pvrchatno"
            MsgBox tmp(0) & " refused private chat."
        Case Chr(1) & "bye"
            RemoveUsr tmp(0)
        Case Else
            AddUsr tmp(0)
            txtMsg = txtMsg & "<" & tmp(0) & ">" & tmp(1) & vbNewLine
            txtMsg.SelStart = Len(txtMsg)
        End Select
    Next
End Sub

Public Sub Go(sRoom As String)
    Me.Tag = sRoom
    Me.Caption = "Chatting in room '" & Me.Tag & "'"
    cChat.Init Me.Tag
    If cChat.GetError = 1 Then
        MsgBox "Unable to create the mailslot."
        Unload Me
    End If
    cChat.SendMessage sUsername, "*", Chr(1) & "userlist"
    Center Me
End Sub

Private Sub AddUsr(sUser As String)
    Dim bThere As Boolean
    
    bThere = False
    For X = 0 To lstUsr.ListCount - 1
        If lstUsr.List(X) = sUser Then
            bThere = True
        End If
    Next
    If Not bThere Then
        lstUsr.AddItem sUser
    End If
End Sub

Private Sub RemoveUsr(sUser As String)
    For i = 0 To lstUsr.ListCount - 1
        If lstUsr.List(i) = sUser Then
            lstUsr.RemoveItem i
            Exit Sub
        End If
    Next
End Sub

