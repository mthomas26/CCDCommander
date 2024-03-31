Attribute VB_Name = "EMail"
' Copyright (C) 2004-2024 Matthew Thomas
'
' This file is part of CCD Commander.
'
' CCD Commander is free software: you can redistribute it and/or modify it under the terms of the GNU
' General Public License as published by the Free Software Foundation, version 3 of the License.
'
' CCD Commander is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
' even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
' Public License for more details.
'
' You should have received a copy of the GNU General Public License along with CCD Commander.
' If not, see <https://www.gnu.org/licenses/>.
'
'---------------------------------------------------------------------------------------------------------

Option Explicit

Private Const AF_INET = 2
Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1
Private Const FD_READ = &H1&
Private Const FD_WRITE = &H2&
Private Const FD_CONNECT = &H10&
Private Const FD_CLOSE = &H20&
Private Const PF_INET = 2
Private Const SOCK_STREAM = 1
Private Const IPPROTO_TCP = 6
Private Const GWL_WNDPROC = (-4)
Private Const WINSOCKMSG = 1025
Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
Private Const INADDR_NONE = &HFFFF
Private Const SOL_SOCKET = &HFFFF&
Private Const SO_LINGER = &H80&
Private Const hostent_size = 16
Private Const sockaddr_size = 16
Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type
Private Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function Send Lib "wsock32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal BufLen As Long, ByVal flags As Long) As Long
Private Declare Function recv Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal BufLen As Long, ByVal flags As Long) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Private Declare Function Connect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Private saZero As sockaddr
Private WSAStartedUp As Boolean, Obj As TextBox
Private PrevProc As Long, lSocket As Long

Private lFromSocketValue As Long
Private ReadMessage As String
Private MessageReady As Boolean
Private TimeOut As Boolean

'subclassing functions
Private Sub HookForm(F As Form)
    PrevProc = SetWindowLong(F.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Private Sub UnHookForm(F As Form)
    If PrevProc <> 0 Then
        SetWindowLong F.hwnd, GWL_WNDPROC, PrevProc
        PrevProc = 0
    End If
End Sub
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WINSOCKMSG Then
        ProcessMessage wParam, lParam
    Else
        WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    End If
End Function
'our Winsock-message handler
Private Sub ProcessMessage(ByVal lFromSocket As Long, ByVal lParam As Long)
    Dim X As Long, ReadBuffer(1 To 1024) As Byte, strCommand As String
    Select Case lParam
        Case FD_CONNECT 'we are connected to microsoft.com
        Case FD_WRITE 'we can write to our connection
            lFromSocketValue = lFromSocket
        Case FD_READ 'we have data waiting to be processed
            'start reading the data
            Do
                X = recv(lFromSocket, ReadBuffer(1), 1024, 0)
                If X > 0 Then
                    ReadMessage = ReadMessage + Left$(StrConv(ReadBuffer, vbUnicode), X)
                End If
                If X <> 1024 Then Exit Do
            Loop
            If (InStr(ReadMessage, vbCrLf) > 0) Then
                MessageReady = True
            End If
        Case FD_CLOSE 'the connection with microsoft.com is closed
    End Select
End Sub

Private Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function
Private Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub
Private Function SendData(ByVal s&, vMessage As Variant) As Long
    Dim TheMsg() As Byte, sTemp$
    TheMsg = ""
    Select Case VarType(vMessage)
        Case 8209   'byte array
            sTemp = vMessage
            TheMsg = sTemp
        Case 8      'string, if we recieve a string, its assumed we are linemode
            sTemp = StrConv(vMessage, vbFromUnicode)
        Case Else
            sTemp = CStr(vMessage)
            sTemp = StrConv(vMessage, vbFromUnicode)
    End Select
    TheMsg = sTemp
    If UBound(TheMsg) > -1 Then
        SendData = Send(s, TheMsg(0), (UBound(TheMsg) - LBound(TheMsg) + 1), 0)
    End If
End Function
Private Function ConnectSock(ByVal Host$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, ByVal Async%) As Long
    Dim s&, SelectOps&, Dummy&
    Dim sockin As sockaddr
    Dim SockReadBuffer$
    
    SockReadBuffer$ = ""
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(Host$)

    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    retIpPort$ = getascip$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    s = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If SetSockLinger(s, 1, 0) = SOCKET_ERROR Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Not Async Then
        If Connect(s, sockin, sockaddr_size) <> 0 Then
            If s > 0 Then
                Dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            If s > 0 Then
                Dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(s, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            If s > 0 Then
                Dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If Connect(s, sockin, sockaddr_size) <> -1 Then
            If s > 0 Then
                Dummy = closesocket(s)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    End If
    ConnectSock = s
End Function
Private Function GetHostByNameAlias(ByVal hostname$) As Long
    On Error Resume Next
    Dim phe&
    Dim heDestHost As HostEnt
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function
Private Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function
Private Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
    Dim Linger As LingerType
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
        End If
    End If
End Function

Private Function GetResponce(Optional FileNo As Integer = 0) As String
    Dim StartTime As Date
    
    StartTime = Now
        
    TimeOut = False
    
    Do While MessageReady = False And DateDiff("s", StartTime, Now) < 15
        Call Wait(0.1)
    Loop
    
    If (MessageReady = False) Then
        TimeOut = True
    End If
    
    GetResponce = ReadMessage
    
    MessageReady = False
    ReadMessage = ""
    
    If (FileNo <> 0) Then
        Print #FileNo, GetResponce;
    End If
    
End Function

Private Sub SendMessage(strCommand As String, Optional FileNo As Integer = 0)
    If (FileNo <> 0) Then
        Print #FileNo, strCommand;
    End If
    
    SendData lFromSocketValue, strCommand
End Sub

Public Function SendEMailNoAuth(FormWithFocus As Form, SMTPAddress As String, SMTPPort As Long, FromAddress As String, ToAddress As String, Message As String, Optional FileNo As Integer = 0) As Boolean
    Dim myReadMessage As String
    
    Dim sSave As String
    FormWithFocus.AutoRedraw = True
    'Start subclassing
    HookForm FormWithFocus
    'create a new winsock session
    StartWinsock sSave
    
    'connect
    lSocket = ConnectSock(SMTPAddress, SMTPPort, 0, FormWithFocus.hwnd, False)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "220") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    Call SendMessage("HELO CCDCommander" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    Call SendMessage("MAIL FROM:<" & FromAddress & ">" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    Call SendMessage("RCPT TO:<" & ToAddress & ">" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    Call SendMessage("DATA" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "354") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    Call SendMessage(Message, FileNo)
        
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    Call SendMessage("QUIT" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "221") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailNoAuth = False
        GoTo SendEMailNoAuthCleanup
    End If
    
    SendEMailNoAuth = True
        
SendEMailNoAuthCleanup:
    'close our connection
    closesocket lSocket
    'end winsock session
    EndWinsock
    'stop subclassing
    UnHookForm FormWithFocus
End Function

Public Function SendEMailWithAuth(FormWithFocus As Form, SMTPAddress As String, SMTPPort As Long, FromAddress As String, ToAddress As String, Message As String, Username As String, Password As String, Optional FileNo As Integer = 0) As Boolean
    Dim myReadMessage As String
    
    Dim sSave As String
    FormWithFocus.AutoRedraw = True
    'Start subclassing
    HookForm FormWithFocus
    'create a new winsock session
    StartWinsock sSave
    
    'connect
    lSocket = ConnectSock(SMTPAddress, SMTPPort, 0, FormWithFocus.hwnd, False)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "220") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage("EHLO CCDCommander" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    If (InStr(myReadMessage, "250-AUTH") = 0) Or (InStr(myReadMessage, "LOGIN") = 0) Then
        'server doesn't support authentication - quit
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage("AUTH LOGIN" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "334") > 0 Or TimeOut
    
    If TimeOut Then
        'server isn't prompting for the username
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage(Enc64.Encode64(Username) & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "334") > 0 Or TimeOut
    
    If TimeOut Then
        'server isn't prompting for the password
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage(Enc64.Encode64(Password) & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "235") > 0 Or TimeOut
    
    If TimeOut Then
        'login failed
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage("MAIL FROM:<" & FromAddress & ">" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage("RCPT TO:<" & ToAddress & ">" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage("DATA" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "354") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage(Message, FileNo)
        
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "250") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    Call SendMessage("QUIT" & vbCrLf, FileNo)
    
    myReadMessage = ""
    Do
        myReadMessage = myReadMessage + GetResponce(FileNo)
    Loop Until InStr(myReadMessage, "221") > 0 Or TimeOut
    
    If TimeOut Then
        SendEMailWithAuth = False
        GoTo SendEMailWithAuthCleanup
    End If
    
    SendEMailWithAuth = True
        
SendEMailWithAuthCleanup:
    'close our connection
    closesocket lSocket
    'end winsock session
    EndWinsock
    'stop subclassing
    UnHookForm FormWithFocus
End Function

Public Function SendEMail(TopWindow As Form, Subject As String, Message As String, Optional FileNo As Integer = 0)
    Dim Counter As Integer
    Dim Result As Boolean
    Dim TotalMessage As String
    
    If (frmOptions.txtEMailScript.Text <> "") Then
        Call RunProgram.RunScriptDirect(frmOptions.txtEMailScript.Text, False, Chr(34) & Subject & Chr(34) & " " & Chr(34) & Message & Chr(34))
    End If
    
    Counter = 0
    For Counter = 0 To frmOptions.lstToAddresses.ListCount - 1
        TotalMessage = "Subject: " & Subject & vbCrLf & "From: CCDCommander <" & frmOptions.txtFromAddress.Text & ">" & vbCrLf & "To: " & frmOptions.lstToAddresses.List(Counter) & vbCrLf & vbCrLf & Message & vbCrLf & "." & vbCrLf
        
        If frmOptions.chkAuthentication.Value = vbUnchecked Then
            Result = SendEMailNoAuth(TopWindow, frmOptions.txtSMTPServer.Text, Settings.SMTPPort, frmOptions.txtFromAddress.Text, frmOptions.lstToAddresses.List(Counter), TotalMessage, FileNo)
        Else
            Result = SendEMailWithAuth(TopWindow, frmOptions.txtSMTPServer.Text, Settings.SMTPPort, frmOptions.txtFromAddress.Text, frmOptions.lstToAddresses.List(Counter), TotalMessage, frmOptions.txtSMTPUsername.Text, frmOptions.txtSMTPPassword.Text, FileNo)
        End If
        
        If Result = False Then
            Call AddToStatus("Error sending email to: " & frmOptions.lstToAddresses.List(Counter))
        End If
    Next Counter
End Function


