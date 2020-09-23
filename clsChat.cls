VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Handle As Long
Private sMailslot As String
Private m_Error As Integer

Public Sub Init(sMail As String)
    Dim tSecr As SECURITY_ATTRIBUTES
    
    sMailslot = sMail
    tSecr.nLength = Len(tSecr)
    tSecr.bInheritHandle = False

    m_Handle = CreateMailslot("\\.\mailslot\" & sMailslot, 0, MAILSLOT_WAIT_FOREVER, tSecr)
    If m_Handle = -1 Then
        m_Error = 1
    End If
End Sub

Public Sub SendMessage(PopFrom As String, PopTo As String, MsgText As String)
    Dim Rc As Long
    Dim Mshandle As Long
    Dim Msgtxt As String
    Dim BytesWritten As Long
    Dim MailslotName As String
    
    MailslotName = "\\" + PopTo + "\mailslot\" & sMailslot
    Msgtxt = PopFrom & Chr(0) & MsgText
    Mshandle = CreateFile(MailslotName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, -1)
    
    Rc = WriteFile(Mshandle, Msgtxt, Len(Msgtxt), BytesWritten, 0)
    Rc = CloseHandle(Mshandle)
End Sub

Public Function GetMessages() As String()
    Dim sTmp As String
    Dim Buffer() As Byte
    Dim sMessages() As String
    Dim mTmp() As String
    Dim bDupe As Boolean
    
    FHandle = GetMailslotInfo(m_Handle, 0, NextSize, Waiting, MAILSLOT_WAIT_FOREVER)

    Do Until Waiting = 0
    
        ReDim Buffer(NextSize - 1)
        FHandle = ReadFile(m_Handle, Buffer(0), NextSize, ReadSize, ByVal 0&)
        
        sTmp = ""
        For c = 0 To NextSize - 1
            sTmp = sTmp & Chr(Buffer(c))
        Next
        For X = 0 To i - 1
            If mTmp(X) = sTmp Then
                bDupe = True
            End If
        Next
        If Not bDupe Then
            ReDim Preserve mTmp(i)
            mTmp(i) = sTmp
            i = i + 1
        End If
        bDupe = False
        FHandle = GetMailslotInfo(m_Handle, 0, NextSize, Waiting, MAILSLOT_WAIT_FOREVER)
    Loop
    GetMessages = mTmp()
End Function

Private Sub Class_Terminate()
    If m_Handle <> -1 Then
        CloseHandle m_Handle
    End If
End Sub

Public Sub CloseHnd()
    If m_Handle <> -1 Then
        CloseHandle m_Handle
    End If
End Sub

Public Property Get GetError() As Integer
    GetError = m_Error
End Property
