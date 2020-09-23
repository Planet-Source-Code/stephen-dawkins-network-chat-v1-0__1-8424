Attribute VB_Name = "modDeclares"
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Const MAILSLOT_WAIT_FOREVER = (-1)
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Declare Function CloseHandle Lib "kernel32" (ByVal hHandle As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFileName As Long, ByVal lpBuff As Any, ByVal nNrBytesToWrite As Long, lpNrOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwAccess As Long, ByVal dwShare As Long, ByVal lpSecurityAttrib As Long, ByVal dwCreationDisp As Long, ByVal dwAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailslot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CreateMailslot Lib "kernel32.dll" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public sUsername As String
Public fRooms() As New frmRoom
Public iRooms As Integer
Public pChats() As New frmPrivateChat
Public iChats As Integer

Public Sub CreateRoomChat(sRoom As String)
    ReDim Preserve fRooms(iRooms)
    fRooms(iRooms).Visible = True
    fRooms(iRooms).Go sRoom
    iRooms = iRooms + 1
End Sub

Public Sub CreatePrivateChat(sUser As String, bMeInit As Boolean)
    ReDim Preserve pChats(iChats)
    pChats(iChats).Visible = True
    pChats(iChats).Go sUser, bMeInit
    iChats = iChats + 1
End Sub

Public Sub UnloadAll()
    On Error Resume Next
    For i = 0 To iRooms - 1
        fRooms(i).cChat.CloseHnd
    Next
    For i = 0 To iChats - 1
        pChats(i).cChat.CloseHnd
    Next
    frmPrivateMessage.cChat.CloseHnd
End Sub

Public Sub Center(frm As Form)
    frm.Left = (Screen.Width / 2) - (frm.Width / 2)
    frm.Top = (Screen.Height / 2) - (frm.Height / 2)
End Sub

Public Sub SendMessageToWinPopUp(PopFrom As String, PopTo As String, MsgText As String)
    Dim Rc As Long
    Dim Mshandle As Long
    Dim Msgtxt As String
    Dim BytesWritten As Long
    Dim MailslotName As String
    
    MailslotName = "\\" + PopTo + "\mailslot\messngr"
    Msgtxt = PopFrom & Chr(0) & PopTo & Chr(0) & MsgText
    Mshandle = CreateFile(MailslotName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, -1)
    
    Rc = WriteFile(Mshandle, Msgtxt, Len(Msgtxt), BytesWritten, 0)
    Rc = CloseHandle(Mshandle)
End Sub
