VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaximPlateSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
' Declare Type for API call:
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128   '  Maintenance string for PSS usage
End Type

' API declarations:
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long

' Constant declarations:
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_CAPITAL = &H14
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const VER_PLATFORM_WIN32_NT = 2
Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Function IsCapsLockOn() As Boolean
        Dim o As OSVERSIONINFO

        o.dwOSVersionInfoSize = Len(o)
        GetVersionEx o
        Dim keys(0 To 255) As Byte
        GetKeyboardState keys(0)
        IsCapsLockOn = keys(VK_CAPITAL)
End Function

Private Sub ToggleCapsLock()
        Dim o As OSVERSIONINFO

        o.dwOSVersionInfoSize = Len(o)
        GetVersionEx o
        Dim keys(0 To 255) As Byte
        GetKeyboardState keys(0)

        If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '=====Win95
        'Toggle capslock
            keys(VK_CAPITAL) = Abs(Not keys(VK_CAPITAL))
            SetKeyboardState keys(0)
        ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then   '=====WinNT
          'Simulate Key Press>
            keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
          'Simulate Key Release
            keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY _
               Or KEYEVENTF_KEYUP, 0
        End If
End Sub

Public Function PlateSolve(FileName As String, PixelScale As Double, NorthAngle As Double, RA As Double, Dec As Double) As Boolean
    'This is the MaxIm/PinPoint LE function.  The full version on PinPoint is implemented in the PinPoint class
    Dim MyDoc As MaxIm.Document
    Dim StartTime As Date
    Dim Retry As Boolean
    Dim PinPointWindow As Long
    Dim ProcessButton As Long
    Dim PinPointErrorWindow As Long
    Dim Counter1 As Long
    Dim Counter2 As Long
    Dim PlateSolveStatus As Boolean
    Dim CapsLockOn As Boolean
        
    Call AddToStatus("Performing Plate Solve with MaxIm/PinPointLE...")
    
    CapsLockOn = IsCapsLockOn()
    
    If CapsLockOn Then ToggleCapsLock
    DoEvents
    
    Call SetForegroundWindow(frmMain.hwnd)
    
    Retry = False
    Set MyDoc = New MaxIm.Document
    
    With MyDoc
        Call .OpenFile(FileName)
        
        'Put pixel scale into FITS header
        Call .SetFITSKey("CDELT1", PixelScale / 3600)
        Call .SetFITSKey("CDELT2", PixelScale / 3600)
        
        'Force a value to this key just in case the image has a plate solve before calling this function
        Call .SetFITSKey("PLTSOLVD", False)
    End With

    Do
        Counter1 = 0
        Do
            Call Windows.FnSetForegroundWindow("MaxIm DL*")
            
            Call Wait(1)
            SendKeys "%(A)", True
            Call Wait(2)
            SendKeys "a", True
            'SendKeys "~", True  'This is the enter key
            
            'this should start pinpoint, try to get the handle to the pinpoint window
            Counter2 = 0
            Do
                Call Wait(1)
                PinPointWindow = FindWindow(vbNullString, "PinPoint Astrometry")
                Counter2 = Counter2 + 1
            Loop While PinPointWindow = 0 And Counter2 < 10 And Not Aborted
            
            Counter1 = Counter1 + 1
        Loop While PinPointWindow = 0 And Counter1 < 10 And Not Aborted
            
        If PinPointWindow <> 0 Then
            'PinPoint window is up and running!
        
            'Now find the handle to the Process button
            ProcessButton = FindWindowEx(PinPointWindow, 0, vbNullString, "P&rocess")
            
            'now try to execute the plate solve
            Counter1 = 0
            Do
                Call SetForegroundWindow(PinPointWindow)
                DoEvents
                
                'SendKeys "%(R)", True
                Call PostMessage(ProcessButton, &H201, &H1, &HA000A)
                Call PostMessage(ProcessButton, &H202, &H1, &HA000A)
                
                'that should start the pinpoint process - check the pinpoint button to see that it is disabled
                Counter2 = 0
                Do
                    Call Wait(0.1)
                    Counter2 = Counter2 + 1
                Loop Until IsWindowEnabled(ProcessButton) = 0 Or Counter2 = 20 Or Aborted
            
                Counter1 = Counter1 + 1
            Loop Until IsWindowEnabled(ProcessButton) = 0 Or Counter1 = 5 Or Aborted
        
            'plate solve is running! Now just wait for it to complete
            PinPointErrorWindow = 0
            StartTime = Now
            Do While (MyDoc.GetFITSKey("PLTSOLVD") = False) And (Not Aborted) And ((CDbl(DateDiff("s", StartTime, Now)) / 60) < Settings.PinPointLETimeout) And (PinPointErrorWindow = 0)
                Call Wait(1)
                PinPointErrorWindow = FindWindow(vbNullString, "PinPoint error")
            Loop
                
            If ((CDbl(DateDiff("s", StartTime, Now)) / 60) >= Settings.PinPointLETimeout) Or Aborted Or PinPointErrorWindow <> 0 Then
                'timed out...
                If Not Retry And (frmOptions.chkPinPointLERetry = vbChecked) Then
                    Call AddToStatus("Plate Solve time out reached.  Retrying...")
                    Retry = True
                Else
                    Call AddToStatus("Plate Solve time out reached.  Giving up.")
                    Retry = False
                End If
                
                PlateSolveStatus = False
            Else
                PlateSolveStatus = True
            End If
        End If
                
        If PinPointErrorWindow <> 0 Then
            Call SetForegroundWindow(PinPointErrorWindow)
            
            DoEvents
            
            SendKeys "%({F4})", True
            
            Call Wait(0.5)
        End If
        
        'close out the pinpoint window
        Do While IsWindowVisible(PinPointWindow) <> 0
            Call SetForegroundWindow(PinPointWindow)
            
            DoEvents
            
            SendKeys "%({F4})", True
            'Call SendMessage(PinPointWindow, &H112, &HF060, 0)
            
            Call Wait(0.5)
            
            PinPointWindow = FindWindow(vbNullString, "PinPoint Astrometry")
        Loop
                
        Call SetForegroundWindow(frmMain.hwnd)
        
        If Not Aborted And PlateSolveStatus Then
            With MyDoc
                NorthAngle = .PositionAngle
                RA = .CenterRA
                Dec = .CenterDec
                
                Call AddToStatus("Plate solve results:")
                Call AddToStatus("Pixel Scale = " & Format(.GetFITSKey("CDELT1") * 3600, "0.00") & " asp.")
                Call AddToStatus("Focal Length = " & Format(206.265 * CDbl(.GetFITSKey("XPIXSZ")) / Abs((.GetFITSKey("CDELT1") + .GetFITSKey("CDELT2")) / 2 * 3600), "0.00") & " mm.")
                Call AddToStatus("North Angle = " & Format(NorthAngle, "0.00") & " degrees.")
                Call AddToStatus("J2000 Coordinates = " & Misc.ConvertEquatorialToString(RA, Dec, False))
            End With
            PlateSolve = True
        Else
            PlateSolve = False
        End If
    Loop While Retry = True And PlateSolve = False
    
    Call MyDoc.Close
    
    If CapsLockOn Then ToggleCapsLock
    DoEvents
    
End Function



