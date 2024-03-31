Attribute VB_Name = "DomeControl"
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

Public Enum DomeControlTypes
    None = 0
    AutomaDome = 1
    DDW = 2
    ASCOM = 3
    AutomaDomeX = 4
End Enum

Private objDomeControl As Object

Public DomeEnabled As Boolean

Private Sub DomeError(ByVal FunctionName As String, ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal ErrDescription As String)
    Call AddToStatus("Dome Error in " + FunctionName + ":" + vbCrLf + vbTab + "Error Number = " + CStr(ErrNumber) + vbCrLf + vbTab + "Error Source = " + ErrSource + vbCrLf + vbTab + "Error Description = " + ErrDescription)

    If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
        'Send e-mail!
        Call EMail.SendEMail(frmMain, "CCD Commander Dome Operation Failed", "Error Number = " + CStr(ErrNumber) + vbCrLf + "Error Source = " + ErrSource + vbCrLf + "Error Description = " + ErrDescription & vbCrLf & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
    End If
End Sub

Public Function IsDomeOpen(Optional Force As Boolean = False, Optional State As Boolean = False) As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If DomeEnabled Then
IsDomeOpenRetry:
        IsDomeOpen = objDomeControl.IsDomeOpen(Force, State)
        If Err.Number = -1 Then
            On Error GoTo 0
            Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
            DomeEnabled = False
        ElseIf Err.Number = -2147220472 And Retry Then
            'Dome not connected error
            Retry = False
            Call AddToStatus("Recevied Dome not connected error - attempting to reconnect one time.")
            Call objDomeControl.DisconnectFromDome
            Call Wait(10)
            Call objDomeControl.ConnectToDome
            GoTo IsDomeOpenRetry
        ElseIf Err.Number = -2147220436 Then
            'Pass this error on to the calling routine - may be ignored if in the timer processing
            On Error GoTo 0
            Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        ElseIf Err.Number <> 0 Then
            Call DomeError("IsDomeOpen", Err.Number, Err.Source, Err.Description)
            On Error GoTo 0
            If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
        End If
    End If
End Function

Public Function IsDomeCoupled() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If DomeEnabled Then
IsDomeCoupledRetry:
        IsDomeCoupled = objDomeControl.IsDomeCoupled
    
        If Err.Number = -1 Then
            On Error GoTo 0
            Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
            DomeEnabled = False
        ElseIf Err.Number = -2147220472 And Retry Then
            'Dome not connected error
            Retry = False
            Call AddToStatus("Recevied Dome not connected error - attempting to reconnect one time.")
            Call objDomeControl.DisconnectFromDome
            Call Wait(10)
            Call objDomeControl.ConnectToDome
            GoTo IsDomeCoupledRetry
        ElseIf Err.Number = -2147220436 Then
            'Pass this error on to the calling routine - may be ignored if in the timer processing
            On Error GoTo 0
            Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        ElseIf Err.Number <> 0 Then
            Call DomeError("IsDomeCoupled", Err.Number, Err.Source, Err.Description)
            On Error GoTo 0
            If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
        End If
    End If
End Function

Public Function OpenDome() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
OpenDomeRetry:
    objDomeControl.DomeOpen

    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        OpenDome = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Recevied Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo OpenDomeRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("OpenDome", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        On Error GoTo 0
        OpenDome = True
    End If
End Function

Public Function CloseDome() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If Not DomeEnabled Then
        CloseDome = False
        Exit Function
    End If
    
    If (frmOptions.chkParkMountFirst.Value = vbChecked) Then
        'Check if the mount is connected - if so, park it!
        'If not connected, then I must have already parked it.
        If Mount.TelescopeConnected Then
            Call Mount.ParkMount
        End If
    End If
    
CloseDomeRetry:
    objDomeControl.DomeClose

    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        CloseDome = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Received Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo CloseDomeRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("CloseDome", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        CloseDome = True
    End If
End Function

Public Function CoupleDome() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If Not DomeEnabled Then
        CoupleDome = False
        Exit Function
    End If
    
CoupleDomeRetry:
    objDomeControl.DomeCouple
    
    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        CoupleDome = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Received Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo CoupleDomeRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("CoupleDome", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        CoupleDome = True
    End If
End Function

Public Function UnCoupleDome() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If Not DomeEnabled Then
        UnCoupleDome = False
        Exit Function
    End If
    
UnCoupleDomeRetry:
    objDomeControl.DomeUnCouple
    
    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        UnCoupleDome = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Received Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo UnCoupleDomeRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("UnCoupleDome", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        UnCoupleDome = True
    End If
End Function

Public Sub SetupDome()
    On Error Resume Next
    
    If frmOptions.lstDomeControl.ListIndex = DomeControlTypes.AutomaDome Then
        Set objDomeControl = New clsAutomaDomeControl
        
        DomeEnabled = True
    ElseIf frmOptions.lstDomeControl.ListIndex = DomeControlTypes.DDW Then
        Set objDomeControl = New clsDDWControl
        
        DomeEnabled = True
    ElseIf frmOptions.lstDomeControl.ListIndex = DomeControlTypes.ASCOM Then
        Set objDomeControl = New clsASCOMDomeControl
        
        DomeEnabled = True
    ElseIf frmOptions.lstDomeControl.ListIndex = DomeControlTypes.AutomaDomeX Then
        Set objDomeControl = New clsAutomaDomeXControl
        
        DomeEnabled = True
    End If
    
    If DomeEnabled Then
        Call objDomeControl.ConnectToDome
    
        If Err.Number = -1 Then
            On Error GoTo 0
            Call AddToStatus("Error communicating with Dome! Disabling dome.")
            DomeEnabled = False
        ElseIf Err.Number <> 0 Then
            Call DomeError("SetupDome", Err.Number, Err.Source, Err.Description)
            On Error GoTo 0
            If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
        End If
    End If
End Sub

Public Sub DomeUnload()
    Set objDomeControl = Nothing
    DomeEnabled = False
End Sub

Public Sub DisconnectDome()
    If DomeEnabled Then
        Call objDomeControl.DisconnectFromDome
    End If
End Sub

Public Sub DomeAction(clsAction As DomeAction)
    If Not DomeEnabled Then
        Call AddToStatus("Skipping Dome Action - no dome connected.")
        Exit Sub
    End If

    Select Case clsAction.ThisDomeAction
        Case DomeActionTypes.actCloseDome
            Call AddToStatus("Closing Dome.")
            If CloseDome() Then Call AddToStatus("Dome closed.")
        Case DomeActionTypes.actOpenDome
            Call AddToStatus("Opening Dome.")
            If OpenDome() Then Call AddToStatus("Dome open.")
        Case DomeActionTypes.actCoupleDome
            Call AddToStatus("Coupling dome to the mount.")
            If CoupleDome() Then Call AddToStatus("Coupling complete.")
        Case DomeActionTypes.actHomeDome
            Call AddToStatus("Moving dome to home position.")
            If HomeDome() Then Call AddToStatus("Dome at home position.")
        Case DomeActionTypes.actParkDome
            Call AddToStatus("Park dome.")
            If ParkDome() Then Call AddToStatus("Dome Park complete.")
        Case DomeActionTypes.actUnCoupleDome
            Call AddToStatus("Uncoupling dome from the mount.")
            If UnCoupleDome() Then Call AddToStatus("Uncoupling complete.")
        Case DomeActionTypes.actSlewDomeAzimuth
            Call AddToStatus("Slewing dome to azimuth = " & clsAction.Azimuth & " degrees.")
            If DomeSlewToAzimuth(clsAction.Azimuth) Then Call AddToStatus("Slew complete.")
    End Select
End Sub

Private Function DomeSlewToAzimuth(Azimuth As Integer) As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If Not DomeEnabled Then
        DomeSlewToAzimuth = False
        Exit Function
    End If
    
DomeSlewToAzimuthRetry:
    Call objDomeControl.SlewToAzimuth(Azimuth)
    
    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        DomeSlewToAzimuth = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Received Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo DomeSlewToAzimuthRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("DomeSlewToAzimuth", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        DomeSlewToAzimuth = True
    End If
End Function

Private Function HomeDome() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If Not DomeEnabled Then
        HomeDome = False
        Exit Function
    End If
    
HomeDomeRetry:
    objDomeControl.DomeHome
    
    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        HomeDome = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Received Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo HomeDomeRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("HomeDome", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        HomeDome = True
    End If
End Function

Private Function ParkDome() As Boolean
    Dim Retry As Boolean
    Retry = True
    On Error Resume Next
    
    If Not DomeEnabled Then
        ParkDome = False
        Exit Function
    End If
    
ParkDomeRetry:
    objDomeControl.DomePark

    If Err.Number = -1 Then
        On Error GoTo 0
        Call AddToStatus("Error communicating with Dome! Disabling dome and continuing execution...")
        DomeEnabled = False
        ParkDome = False
    ElseIf Err.Number = -2147220472 And Retry Then
        'Dome not connected error
        Retry = False
        Call AddToStatus("Received Dome not connected error - attempting to reconnect one time.")
        Call objDomeControl.DisconnectFromDome
        Call Wait(10)
        Call objDomeControl.ConnectToDome
        GoTo ParkDomeRetry
    ElseIf Err.Number = -2147220436 Then
        'Pass this error on to the calling routine - may be ignored if in the timer processing
        On Error GoTo 0
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    ElseIf Err.Number <> 0 Then
        Call DomeError("ParkDome", Err.Number, Err.Source, Err.Description)
        On Error GoTo 0
        If frmOptions.chkHaltOnDomeError Then Call Err.Raise(65530, "CCD Commander", "Dome Error")
    Else
        ParkDome = True
    End If
End Function
