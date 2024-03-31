VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFocusMaxAcqControl"
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

Private myFocusMax As Object
Private myFocuser As Object
Private FocusMaxEnabled As Boolean

Public Function Focus(clsAction As FocusAction) As Boolean
    Dim myHFD As Double
    Dim Counter As Integer
    Dim FocusResult As Integer
    Dim StartTime As Date
    
    If FocusMaxEnabled Then
        myFocusMax.Exposure = clsAction.ExposureTime
    
        Call myFocusMax.AcquireStarAsync
        Call Wait(4)
        StartTime = Now
        Do
            Call Wait(1)
            FocusResult = myFocusMax.AcquireStarAsyncStatus
        Loop While FocusResult = -1 And Not Aborted And DateDiff("s", StartTime, Now) < (Settings.FocusTimeOut * 60)
    
        If DateDiff("s", StartTime, Now) >= (Settings.FocusTimeOut * 60) Then
            Call AddToStatus("Focus Time Out reached!  Halting FocusMax.")
        
            Counter = 0
            On Error Resume Next
            Do
                Counter = Counter + 1
                myFocusMax.Halt
                Call Wait(1)
            Loop Until Err.Number = 0 Or Counter = 20
            On Error GoTo 0
            
            myHFD = 0
            FocusResult = 0
        ElseIf Aborted Then
            Call AddToStatus("Halting FocusMax")
            Counter = 0
            On Error Resume Next
            
            'need to do this so the wait doesn't immediately return
            Aborted = False
            Do
                Err.Clear
                Counter = Counter + 1
                myFocusMax.Halt
                Call Wait(1)
            Loop Until Err.Number = 0 Or Counter = 60
            Aborted = True
            On Error GoTo 0
        Else
            myHFD = myFocusMax.HalfFluxDiameter
        End If
    
        If myHFD = 0 Or FocusResult = 0 Or Aborted Then
            Call AddToStatus("Focus failed!")
            Focus = False
        Else
            Call AddToStatus("Focus succeeded! HFD = " & Format(myHFD, "0.00"))
            Call AddToStatus("Focus position @ " & myFocuser.Position)
            Focus = True
        End If
    Else
        Call AddToStatus("FocusMax unavailable!  Check the program settings.")
        Focus = True
    End If
End Function

Public Function OffsetFocuser(Offset As Integer) As Integer
    Dim SavedTempCompSetting As Boolean
    
    If myFocuser.TempCompAvailable Then
        SavedTempCompSetting = myFocuser.TempComp
        myFocuser.TempComp = False
    End If

    If myFocuser.Absolute Then
        If (myFocuser.Position + CLng(Offset)) > 0 And (myFocuser.Position + CLng(Offset)) < myFocuser.MaxStep Then
            Call AddToStatus("Moving focuser " & Offset & " steps.")
            Call myFocuser.Move(myFocuser.Position + CLng(Offset))
        Else
            Call AddToStatus("Error - cannot move " & Offset & " steps.")
        End If
    Else
        If Offset > -myFocuser.MaxIncrement And Offset < myFocuser.MaxIncrement Then
            Call AddToStatus("Moving focuser " & Offset & " steps.")
            Call myFocuser.Move(Offset)
        Else
            Call AddToStatus("Error - cannot move " & Offset & " steps.")
        End If
    End If

    Do While Not Aborted And myFocuser.IsMoving
        Call Wait(1)
    Loop

    If myFocuser.TempCompAvailable Then
        myFocuser.TempComp = SavedTempCompSetting
    End If
    
    OffsetFocuser = myFocuser.Position
End Function

Public Sub ConnectToFocuser()
    Dim Counter As Integer
    
    On Error GoTo FocusSetupError
    Set myFocusMax = CreateObject("FocusMax.FocusControl")
    Call Wait(1)    'Needed incase the above opens FocusMax.  Without the delay, the below will try to start another instance of FocusMax
    Set myFocuser = CreateObject("FocusMax.Focuser")
    FocusMaxEnabled = True
    On Error GoTo 0
    Exit Sub
    
FocusSetupError:
    Call AddToStatus("Problem connecting to FocusMax - either it doesn't exist on this system, or it is not functioning.")
    Call AddToStatus("Error = " & Err.Number & ", Description = " & Err.Description)
    FocusMaxEnabled = False
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set myFocusMax = Nothing
    Set myFocuser = Nothing
End Sub
