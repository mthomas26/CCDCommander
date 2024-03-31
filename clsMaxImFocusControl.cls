VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaxImFocusControl"
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

Private objMaxImApp As MaxIm.Application

Public Function Focus(clsAction As FocusAction) As Boolean
    Dim Test As Boolean
    Dim StartTime As Date
    
    If objMaxImApp.FocuserConnected Then
        Call AddToStatus("Setting MaxIm Focus Exposure time to " & Format(clsAction.ExposureTime, "0.000") & " seconds.")
        
        StartTime = Now
        Call objMaxImApp.AutoFocus(clsAction.ExposureTime)
        
        Do While objMaxImApp.AutofocusStatus = -1 And Not Aborted And DateDiff("s", StartTime, Now) < (Settings.FocusTimeOut * 60)
            Call Wait(1)
        Loop
        
        If Aborted Or DateDiff("s", StartTime, Now) >= (Settings.FocusTimeOut * 60) Then
            Call AddToStatus("Focusing timed out - aborting.")
            Camera.CameraAbort
            Focus = False
        Else
            Test = objMaxImApp.AutofocusStatus
            If Test = 0 Then
                Call AddToStatus("Focus Failed.")
                Focus = False
            Else
                Call AddToStatus("Focus Complete.")
                Focus = True
            End If
        End If
    Else
        Call AddToStatus("No focuser connected - skipping focus action")
        Focus = True
    End If
End Function

Public Function OffsetFocuser(Offset As Integer) As Integer
    Call AddToStatus("Cannot move focuser with MaxIm/DL")
    OffsetFocuser = 0
End Function

Public Sub ConnectToFocuser()
    Set objMaxImApp = New MaxIm.Application
    
    On Error Resume Next
    objMaxImApp.FocuserConnected = True
    On Error GoTo 0
    
    If Not objMaxImApp.FocuserConnected Then
        Call AddToStatus("Warning!  No focuser connected!")
    End If
End Sub

Private Sub Class_Terminate()
    Set objMaxImApp = Nothing
End Sub
