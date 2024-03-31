VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtFocus2Control"
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

Private objCCDSoftCam As CCDSoftLib.CCDSoftCamera

Public Function Focus(clsAction As FocusAction) As Boolean
    Dim Test As Boolean
    Dim StartTime As Date
        
    If objCCDSoftCam.focIsConnected = 1 Then
        If (objCCDSoftCam.PropLng("m_baf2AutoSlew") = 1 Or objCCDSoftCam.PropLng("m_baf2AutoExposure")) And InStr(frmOptions.lstMountControl.List(frmOptions.lstMountControl.ListIndex), "TheSky") = 0 Then
            Call AddToStatus("TheSky is required when @Focus2 AutoSlew or AutoExposure is enabled.")
        Else
            StartTime = Now
            objCCDSoftCam.PropLng("m_mAtFocus2") = 1
            Do While objCCDSoftCam.State <> cdStateNone And Not Aborted And DateDiff("s", StartTime, Now) < (Settings.FocusTimeOut * 60)
                Call Wait(1)
            Loop
            
            If Aborted Or DateDiff("s", StartTime, Now) >= (Settings.FocusTimeOut * 60) Then
                Call objCCDSoftCam.Abort
            End If
            
            On Error Resume Next
            Test = objCCDSoftCam.IsExposureComplete
            If Err.Number <> 0 Then
                Call AddToStatus("@Focus Failed.")
                Call AddToStatus(Err.Description)
                Focus = False
            Else
                Call AddToStatus("@Focus Complete.  Position = " & objCCDSoftCam.focPosition)
                Focus = True
            End If
            On Error GoTo 0
        End If
    Else
        Call AddToStatus("No focuser connected - skipping focus action")
        Focus = True
    End If
End Function

Public Function OffsetFocuser(Offset As Integer) As Integer
    If Offset < 0 Then
        Call AddToStatus("Moving focuser In " & Offset & " steps.")
        Call objCCDSoftCam.focMoveIn(Abs(Offset))
    Else
        Call AddToStatus("Moving focuser Out " & Offset & " steps.")
        Call objCCDSoftCam.focMoveOut(Abs(Offset))
    End If
    
    OffsetFocuser = objCCDSoftCam.focPosition
End Function

Public Sub ConnectToFocuser()
    Set objCCDSoftCam = New CCDSoftLib.CCDSoftCamera
    
    objCCDSoftCam.Asynchronous = 1
    
    Call objCCDSoftCam.focConnect
    
    If objCCDSoftCam.focIsConnected <> 1 Then
        Call AddToStatus("Warning!  No focuser connected!")
    End If
End Sub

Private Sub Class_Terminate()
    Set objCCDSoftCam = Nothing
End Sub
