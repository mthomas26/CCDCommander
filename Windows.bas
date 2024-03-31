Attribute VB_Name = "Windows"
' Module Name: ModFindWindowLike
' (c) 2005 Wayne Phillips (http://www.everythingaccess.com)
' Written 02/06/2005

Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'Custom structure for passing in the parameters in/out of the hook enumeration function
'Could use global variables instead, but this is nicer.
Private Type FindWindowParameters

    strTitle As String 'INPUT
    hwnd As Long        'OUTPUT

End Type

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5

Public Function FnFindWindowLike(strWindowTitle As String) As Long
    'We'll pass a custom structure in as the parameter to store our result...
    Dim Parameters As FindWindowParameters
    Parameters.strTitle = strWindowTitle ' Input parameter

    Call EnumWindows(AddressOf EnumWindowProc, VarPtr(Parameters))
    
    FnFindWindowLike = Parameters.hwnd
End Function

Private Function EnumWindowProc(ByVal hwnd As Long, lParam As FindWindowParameters) As Long
    Dim strWindowTitle As String

    strWindowTitle = Space(260)
    Call GetWindowText(hwnd, strWindowTitle, 260)
    strWindowTitle = TrimNull(strWindowTitle) ' Remove extra null terminator
                                            
    If strWindowTitle Like lParam.strTitle Then
        lParam.hwnd = hwnd 'Store the result for later.
        EnumWindowProc = 0 'This will stop enumerating more windows
    End If
                        
    EnumWindowProc = 1
End Function

Private Function TrimNull(strNullTerminatedString As String)
    Dim lngPos As Long

    'Remove unnecessary null terminator
    lngPos = InStr(strNullTerminatedString, Chr$(0))

    If lngPos Then
        TrimNull = Left$(strNullTerminatedString, lngPos - 1)
    Else
        TrimNull = strNullTerminatedString
    End If
End Function

Public Function FnSetForegroundWindow(strWindowTitle As String) As Boolean

    Dim MyAppHWnd As Long
    Dim CurrentForegroundThreadID As Long
    Dim NewForegroundThreadID As Long
    Dim lngRetVal As Long
    
    Dim blnSuccessful As Boolean
    
    MyAppHWnd = FnFindWindowLike(strWindowTitle)
    
    If MyAppHWnd <> 0 Then
        
        'We've found the application window by the caption
            CurrentForegroundThreadID = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
            NewForegroundThreadID = GetWindowThreadProcessId(MyAppHWnd, ByVal 0&)
    
        'AttachThreadInput is used to ensure SetForegroundWindow will work
        'even if our application isn't currently the foreground window
        '(e.g. an automated app running in the background)
            Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, True)
            lngRetVal = SetForegroundWindow(MyAppHWnd)
            Call AttachThreadInput(CurrentForegroundThreadID, NewForegroundThreadID, False)
            
        If lngRetVal <> 0 Then
        
            'Now that the window is active, let's restore it from the taskbar
            If IsIconic(MyAppHWnd) Then
                Call ShowWindow(MyAppHWnd, SW_RESTORE)
            Else
                Call ShowWindow(MyAppHWnd, SW_SHOW)
            End If
            
            blnSuccessful = True
        
        Else
        
            'MsgBox "Found the window, but failed to bring it to the foreground!"
        
        End If
        
    Else
    
        'Failed to find the window caption
        'Therefore the app is probably closed.
        'MsgBox "Application Window '" + strWindowTitle + "' not found!"
    
    End If
    
     FnSetForegroundWindow = blnSuccessful
    
End Function
