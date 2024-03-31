VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTheSkyXAtFocus3Control"
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

Private objCCDSoftCam As Object 'theskyxLib.CCDSoftCamera

Public Function Focus(clsAction As FocusAction) As Boolean
    Dim Test As Boolean
    Dim StartTime As Date
    Dim ErrSource As String
    Dim ErrNumber As String
    Dim ErrDescription As String
        
    If objCCDSoftCam.focIsConnected = 1 Then
        Call AddToStatus("Setting @Focus3 exposure time to " & Format(clsAction.ExposureTime, "0.000") & " seconds.")
        objCCDSoftCam.FocusExposureTime = clsAction.ExposureTime
                                
        'Disable temp comp
        objCCDSoftCam.focTemperatureCompensationMode = cdfocTempCompMode_None
        
        
        On Error Resume Next
        Call objCCDSoftCam.AtFocus3(clsAction.FocusAverages, True)
        
        If Err.Number <> 0 Then
            ErrSource = Err.Source
            ErrNumber = Err.Number
            ErrDescription = Err.Description
            Call AddToStatus("@Focus Failed.")
            Call AddToStatus(ErrSource)
            Call AddToStatus(ErrNumber)
            Call AddToStatus(ErrDescription)
            Focus = False
        Else
            Call AddToStatus("@Focus Complete.  Position = " & objCCDSoftCam.focPosition)
            If (clsAction.TempComp) Then
                Call AddToStatus("Enabling temperature compensation B.")
                objCCDSoftCam.focTemperatureCompensationMode = cdfocTempCompMode_B
            End If
            Focus = True
        End If
        On Error GoTo 0
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
    Set objCCDSoftCam = CreateObject("theskyx.CCDSoftCamera")
    
    objCCDSoftCam.Asynchronous = 1
    
    Call objCCDSoftCam.focConnect
    
    If objCCDSoftCam.focIsConnected <> 1 Then
        Call AddToStatus("Warning!  No focuser connected!")
    End If
End Sub

Private Sub Class_Terminate()
    Set objCCDSoftCam = Nothing
End Sub
