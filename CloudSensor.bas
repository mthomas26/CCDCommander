Attribute VB_Name = "CloudSensor"
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

Public Enum WeatherMonitorControl
    None = 0
    ClarityI = 1
    ClarityII = 2
    ClarityIIRemote = 3
    AAG = 4
    AAGRemote = 5
End Enum

Public Enum CloudConditions
    NotImplemented = -1
    Unknown = 0
    Clear = 1
    Cloudy = 2
    VeryCloudy = 3
End Enum

Public Enum WindConditions
    NotImplemented = -1
    Unknown = 0
    Calm = 1
    Windy = 2
    VeryWindy = 3
End Enum

Public Enum RainConditions
    NotImplemented = -1
    Unknown = 0
    Dry = 1
    Wet = 2
    Rain = 3
End Enum

Public Enum DayLightConditions
    NotImplemented = -1
    Unknown = 0
    Dark = 1
    Light = 2
    VeryLight = 3
End Enum

Private objCloud As Object
Private ClearStartTime As Date
Private ClearCount As Long
Private LastRA As Double
Private LastDec As Double
Private TargetName As String

Private LastRotatorAngle As Double

Private CloudMonitoringEnabled As Boolean

Private DidICloseDome As Boolean

Private DomeUncoupled As Boolean

Private NeedToSetAbortFlag As Boolean

Public Sub CloudMonitorAction(clsAction As CloudMonitorAction)
    If clsAction.Enabled Then
        Call AddToStatus("Weather monitoring enabled.")
        
        'Check Cloud sensor status now!
        Call MyCheckCloudSensor
        
        'now enable the flag to continue checking
        CloudMonitoringEnabled = True
    Else
        CloudMonitoringEnabled = False
        
        Call AddToStatus("Weather monitoring disabled.")
        
        'need to clear the pause between actions flag so the list can continue
        PauseBetweenActions = False
        'Don't do anything else (move mount, open dome) that is up to the user now
    End If
End Sub

Public Sub CheckCloudSensor()
    If CloudMonitoringEnabled = False Then
        Exit Sub
    End If
    
    Call MyCheckCloudSensor
End Sub

Private Sub MyCheckCloudSensor()
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    Dim myString As String
        
    Dim CloudSensorStatus As CloudConditions
    Dim RainSensorStatus As RainConditions
    Dim WindSensorStatus As WindConditions
    Dim LightSensorStatus As DayLightConditions
    
    Dim objAction As MoveRADecAction
    
    On Error GoTo CloudSensorError
    
    NeedToSetAbortFlag = False
   
    CloudSensorStatus = objCloud.CloudStatus
    RainSensorStatus = objCloud.RainStatus
    WindSensorStatus = objCloud.WindStatus
    LightSensorStatus = objCloud.LightStatus
    
    'CloudSensorStatus = 1
   
    If frmOptions.chkParkMountFirst.Value = Checked And Not PauseBetweenActions Then
        If NeedToPauseActionList(CloudSensorStatus, RainSensorStatus, WindSensorStatus, LightSensorStatus, myString) Then
            Call PauseActionListAndPark(myString)
        End If
    End If
        
    If DomeControl.DomeEnabled And DomeControl.IsDomeOpen Then
        If NeedToCloseDome(CloudSensorStatus, RainSensorStatus, WindSensorStatus, LightSensorStatus, myString) Then
            'need to close the dome
            
            If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled Then
                DomeUncoupled = True
            End If
                        
            On Error Resume Next
            Call DomeControl.CloseDome
            If Err.Number <> 0 Then
                On Error GoTo CloudSensorError
                Call AddToStatus("Could not close dome - might already be closed.")
            Else
                On Error GoTo CloudSensorError
                Call AddToStatus("Dome closure complete.")
            End If
            
            If frmOptions.chkEnableWeatherMonitorScripts = vbChecked Then
                On Error Resume Next
                Call RunScriptDirect(frmOptions.txtAfterCloseScript.Text, True)
                On Error GoTo CloudSensorError
            End If
            
            If frmOptions.chkEMailAlert(EMailAlertIndexes.WeatherMonitorDomeClosed).Value = vbChecked Then
                'Send e-mail!
                Call EMail.SendEMail(frmMain, "CCD Commander Weather Monitor - Dome Closed", "Dome closed due to " & myString & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
            End If
            
            DidICloseDome = True
        ElseIf frmOptions.chkAutoDomeClose.Value = vbChecked And _
            ((CloudSensorStatus = CloudConditions.VeryCloudy) Or (RainSensorStatus = RainConditions.Wet) Or (RainSensorStatus = RainConditions.Rain) Or _
            (WindSensorStatus = WindConditions.VeryWindy) Or (LightSensorStatus = DayLightConditions.VeryLight)) Then
            
            Call AddToStatus("Autonomous closure condition detected!")
            If (CloudSensorStatus = CloudConditions.VeryCloudy) Then
                Call AddToStatus("Very cloudy condition detected.")
            End If
            If (RainSensorStatus = RainConditions.Wet) Then
                Call AddToStatus("Wet condition detected.")
            End If
            If (RainSensorStatus = RainConditions.Rain) Then
                Call AddToStatus("Rain detected.")
            End If
            If (WindSensorStatus = WindConditions.VeryWindy) Then
                Call AddToStatus("Very windy condition detected.")
            End If
            If (LightSensorStatus = DayLightConditions.VeryLight) Then
                Call AddToStatus("Very light condition detected.")
            End If
            Call AddToStatus("Dome should be closing automatically via the hardwired interface.")
            
            If frmOptions.lstDomeControl.ListIndex = DomeControlTypes.AutomaDome Then
                'AutomaDome can't query the dome slit state - so need to force the state to closed
                Call DomeControl.IsDomeOpen(True, False)
            End If
        
            If frmOptions.chkEnableWeatherMonitorScripts = vbChecked Then
                On Error Resume Next
                Call RunScriptDirect(frmOptions.txtAfterCloseScript.Text, True)
                On Error GoTo CloudSensorError
            End If
                    
            'need to set this here to allow an opening of the dome later
            DidICloseDome = True
        End If
    End If
    
    If frmOptions.chkParkMountFirst.Value = vbUnchecked And Not PauseBetweenActions Then
        If NeedToPauseActionList(CloudSensorStatus, RainSensorStatus, WindSensorStatus, LightSensorStatus, myString) Then
            Call PauseActionListAndPark(myString)
        End If
    End If
    
    If NeedToSetAbortFlag = True Then
        Aborted = True
        SoftSkip = False 'If SoftSkip was active, then abort due to clouds would fail.
        NeedToSetAbortFlag = False
    End If
        
    If CheckGoodConditions(CloudSensorStatus, RainSensorStatus, WindSensorStatus, LightSensorStatus) Then
        If PauseBetweenActions Then
            'clear - count intervals
            If ClearCount = 0 Then
                ClearStartTime = Now
            End If
            ClearCount = ClearCount + 1
            Call AddToStatus("Good weather condition detected for " & Format(CDbl(DateDiff("s", ClearStartTime, Now)) / 60, "0.0") & " minutes.")
        End If
    Else
        ClearCount = 0
    End If
    
    If (ClearCount > 0) And (CDbl(DateDiff("s", ClearStartTime, Now)) / 60 >= Settings.CloudMonitorClearTime) And PauseBetweenActions Then
        'We are clear enough to continue!
        Call AddToStatus("Good weather condition exists long enough!")
        
        'Run script first thing - could be opening a roof or other important process
        If frmOptions.chkEnableWeatherMonitorScripts = vbChecked Then
            On Error Resume Next
            Call RunScriptDirect(frmOptions.txtAfterGoodScript.Text, True)
            On Error GoTo CloudSensorError
        End If
        
        'did this routine close the dome?
        'if not, I don't want to open it as it wasn't open when the weather condition occurred
        If DidICloseDome Then
            'open the dome
            If Not DomeControl.IsDomeOpen Then
                Call AddToStatus("Opening Dome.")
                Call DomeControl.OpenDome
                Call AddToStatus("Dome open complete.")
            End If
        End If
        DidICloseDome = False
        
        If LastRA >= 0 Then
            'start the slew back to the object coordinates
            Call AddToStatus("Moving back to original location.")
            Set objAction = New MoveRADecAction
            
            objAction.RA = LastRA
            objAction.Dec = LastDec
            objAction.Name = TargetName
            
            Call Mount.MoveRADecAction(objAction)
            
            If Not Aborted And Rotator.RotatorConnected Then
                Call Rotator.Rotate(LastRotatorAngle)
            End If
            
            If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.DomeEnabled = True And DomeUncoupled Then
                Call DomeControl.CoupleDome
                DomeUncoupled = False
            End If
        End If

        If frmOptions.chkEMailAlert(EMailAlertIndexes.WeatherMonitorResuming).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander Weather Monitor - Resuming", "Action list resuming." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
        
        If (Settings.WeatherMonitorRestartActionList) Then
            Call AddToStatus("Restarting action list.")
            PauseBetweenActions = False
            AbortButton = True
            Aborted = True
            RestartActionList = True
        Else
            Call AddToStatus("Resuming action.")
            PauseBetweenActions = False
        End If
    End If
    
    Exit Sub
    
CloudSensorError:
    ErrorNum = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
        
    On Error GoTo 0
    
    If ErrorNum = &H80010005 Then
        'error is automation error - I can just ignore this and try again next time around
    ElseIf ErrorNum <> 0 Then
        Call AddToStatus("Error in Check Cloud Sensor routine.")
        Call AddToStatus("Error Number: " & ErrorNum)
        Call AddToStatus(ErrorDescription)
        Call AddToStatus(ErrorSource)
    End If
End Sub

Public Sub CloudSensorSetup()
    If frmOptions.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityI Then
        Set objCloud = New clsClairtyIControl
    ElseIf frmOptions.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityII Or _
        frmOptions.lstCloudSensor.ListIndex = WeatherMonitorControl.AAG Then
        Set objCloud = New clsClarityIIControl
    ElseIf frmOptions.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityIIRemote Or _
        frmOptions.lstCloudSensor.ListIndex = WeatherMonitorControl.AAGRemote Then
        Set objCloud = New clsClarityIIControlRemote
    End If
    
    Call objCloud.ConnectToWeatherMonitor
    
    LastRA = -1
    DidICloseDome = False
    CloudMonitoringEnabled = False
    DomeUncoupled = False
End Sub

Public Sub CloudSensorUnload()
    Set objCloud = Nothing
End Sub

Private Function NeedToCloseDome(CloudStatus As CloudConditions, RainStatus As RainConditions, WindStatus As WindConditions, LightStatus As DayLightConditions, ByRef CauseString As String) As Boolean
    Dim CloudSensorClose As Boolean
    Dim RainSensorClose As Boolean
    Dim WindSensorClose As Boolean
    Dim LightSensorClose As Boolean
    
    CauseString = ""
    
    CloudSensorClose = ((CloudStatus = CloudConditions.Unknown) And (frmOptions.lstCloudSensorCloseDomeWhen.Selected(0)))
    If (CloudSensorClose) And CauseString = "" Then
        CauseString = "unknown cloud sensor condition."
        Call AddToStatus("Unknown cloud sensor condition detected!  Closing Dome.")
    End If
    CloudSensorClose = CloudSensorClose Or ((CloudStatus = CloudConditions.Cloudy) And (frmOptions.lstCloudSensorCloseDomeWhen.Selected(1)))
    If (CloudSensorClose) And CauseString = "" Then
        CauseString = "cloudy condition."
        Call AddToStatus("Cloudy condition detected!  Closing Dome.")
    End If
    CloudSensorClose = CloudSensorClose Or ((CloudStatus = CloudConditions.VeryCloudy) And (frmOptions.lstCloudSensorCloseDomeWhen.Selected(2)))
    If (CloudSensorClose) And CauseString = "" Then
        CauseString = "very cloudy condition."
        Call AddToStatus("Very cloudy condition detected!  Closing Dome.")
    End If
    
    RainSensorClose = ((RainStatus = RainConditions.Unknown) And (frmOptions.lstRainSensorCloseDomeWhen.Selected(0)))
    If (RainSensorClose) And CauseString = "" Then
        CauseString = "unknown rain sensor condition."
        Call AddToStatus("Unknown rain sensor condition detected!  Closing Dome.")
    End If
    RainSensorClose = RainSensorClose Or ((RainStatus = RainConditions.Wet) And (frmOptions.lstRainSensorCloseDomeWhen.Selected(1)))
    If (RainSensorClose) And CauseString = "" Then
        CauseString = "wet condition."
        Call AddToStatus("Wet condition detected!  Closing Dome.")
    End If
    RainSensorClose = RainSensorClose Or ((RainStatus = RainConditions.Rain) And (frmOptions.lstRainSensorCloseDomeWhen.Selected(2)))
    If (RainSensorClose) And CauseString = "" Then
        CauseString = "rain."
        Call AddToStatus("Rain detected!  Closing Dome.")
    End If
    
    WindSensorClose = ((WindStatus = WindConditions.Unknown) And (frmOptions.lstWindSensorCloseDomeWhen.Selected(0)))
    If (WindSensorClose) And CauseString = "" Then
        CauseString = "unknown wind sensor condition."
        Call AddToStatus("Unknown wind sensor condition detected!  Closing Dome.")
    End If
    WindSensorClose = WindSensorClose Or ((WindStatus = WindConditions.Windy) And (frmOptions.lstWindSensorCloseDomeWhen.Selected(1)))
    If (WindSensorClose) And CauseString = "" Then
        CauseString = "windy condition."
        Call AddToStatus("Windy condition detected!  Closing Dome.")
    End If
    WindSensorClose = WindSensorClose Or ((WindStatus = WindConditions.VeryWindy) And (frmOptions.lstWindSensorCloseDomeWhen.Selected(2)))
    If (WindSensorClose) And CauseString = "" Then
        CauseString = "very windy condition."
        Call AddToStatus("Very windy condition detected!  Closing Dome.")
    End If
    
    LightSensorClose = ((LightStatus = DayLightConditions.Unknown) And (frmOptions.lstLightSensorCloseDomeWhen.Selected(0)))
    If (LightSensorClose) And CauseString = "" Then
        CauseString = "unknown light sensor condition."
        Call AddToStatus("Unknown light sensor condition detected!  Closing Dome.")
    End If
    LightSensorClose = LightSensorClose Or ((LightStatus = DayLightConditions.Light) And (frmOptions.lstLightSensorCloseDomeWhen.Selected(1)))
    If (LightSensorClose) And CauseString = "" Then
        CauseString = "light condition."
        Call AddToStatus("Light condition detected!  Closing Dome.")
    End If
    LightSensorClose = LightSensorClose Or ((LightStatus = DayLightConditions.VeryLight) And (frmOptions.lstLightSensorCloseDomeWhen.Selected(2)))
    If (LightSensorClose) And CauseString = "" Then
        CauseString = "very light condition."
        Call AddToStatus("Very light condition detected!  Closing Dome.")
    End If
            
    NeedToCloseDome = CloudSensorClose Or RainSensorClose Or WindSensorClose Or LightSensorClose
End Function

Private Function NeedToPauseActionList(CloudStatus As CloudConditions, RainStatus As RainConditions, WindStatus As WindConditions, LightStatus As DayLightConditions, ByRef CauseString As String) As Boolean
    Dim CloudSensorPause As Boolean
    Dim RainSensorPause As Boolean
    Dim WindSensorPause As Boolean
    Dim LightSensorPause As Boolean
    
    CauseString = ""
    
    CloudSensorPause = ((CloudStatus = CloudConditions.Unknown) And (frmOptions.lstCloudSensorPauseActionWhen.Selected(0)))
    If (CloudSensorPause) And CauseString = "" Then
        CauseString = "unknown cloud sensor condition."
        Call AddToStatus("Unknown cloud sensor condition detected!  Pausing Action List.")
    End If
    CloudSensorPause = CloudSensorPause Or ((CloudStatus = CloudConditions.Cloudy) And (frmOptions.lstCloudSensorPauseActionWhen.Selected(1)))
    If (CloudSensorPause) And CauseString = "" Then
        CauseString = "cloudy condition."
        Call AddToStatus("Cloudy condition detected!  Pausing Action List.")
    End If
    CloudSensorPause = CloudSensorPause Or ((CloudStatus = CloudConditions.VeryCloudy) And (frmOptions.lstCloudSensorPauseActionWhen.Selected(2)))
    If (CloudSensorPause) And CauseString = "" Then
        CauseString = "very cloudy condition."
        Call AddToStatus("Very cloudy condition detected!  Pausing Action List.")
    End If
    
    RainSensorPause = ((RainStatus = RainConditions.Unknown) And (frmOptions.lstRainSensorPauseActionWhen.Selected(0)))
    If (RainSensorPause) And CauseString = "" Then
        CauseString = "unknown rain sensor condition."
        Call AddToStatus("Unknown rain sensor condition detected!  Pausing Action List.")
    End If
    RainSensorPause = RainSensorPause Or ((RainStatus = RainConditions.Wet) And (frmOptions.lstRainSensorPauseActionWhen.Selected(1)))
    If (RainSensorPause) And CauseString = "" Then
        CauseString = "wet condition."
        Call AddToStatus("Wet condition detected!  Pausing Action List.")
    End If
    RainSensorPause = RainSensorPause Or ((RainStatus = RainConditions.Rain) And (frmOptions.lstRainSensorPauseActionWhen.Selected(2)))
    If (RainSensorPause) And CauseString = "" Then
        CauseString = "rain."
        Call AddToStatus("Rain detected!  Pausing Action List.")
    End If
    
    WindSensorPause = ((WindStatus = WindConditions.Unknown) And (frmOptions.lstWindSensorPauseActionWhen.Selected(0)))
    If (WindSensorPause) And CauseString = "" Then
        CauseString = "unknown wind sensor condition."
        Call AddToStatus("Unknown wind sensor condition detected!  Pausing Action List.")
    End If
    WindSensorPause = WindSensorPause Or ((WindStatus = WindConditions.Windy) And (frmOptions.lstWindSensorPauseActionWhen.Selected(1)))
    If (WindSensorPause) And CauseString = "" Then
        CauseString = "windy condition."
        Call AddToStatus("Windy condition detected!  Pausing Action List.")
    End If
    WindSensorPause = WindSensorPause Or ((WindStatus = WindConditions.VeryWindy) And (frmOptions.lstWindSensorPauseActionWhen.Selected(2)))
    If (WindSensorPause) And CauseString = "" Then
        CauseString = "very windy condition."
        Call AddToStatus("Very windy condition detected!  Pausing Action List.")
    End If
    
    LightSensorPause = ((LightStatus = DayLightConditions.Unknown) And (frmOptions.lstLightSensorPauseActionWhen.Selected(0)))
    If (LightSensorPause) And CauseString = "" Then
        CauseString = "unknown light sensor condition."
        Call AddToStatus("Unknown light sensor condition detected!  Pausing Action List.")
    End If
    LightSensorPause = LightSensorPause Or ((LightStatus = DayLightConditions.Light) And (frmOptions.lstLightSensorPauseActionWhen.Selected(1)))
    If (LightSensorPause) And CauseString = "" Then
        CauseString = "light condition."
        Call AddToStatus("Light condition detected!  Pausing Action List.")
    End If
    LightSensorPause = LightSensorPause Or ((LightStatus = DayLightConditions.VeryLight) And (frmOptions.lstLightSensorPauseActionWhen.Selected(2)))
    If (LightSensorPause) And CauseString = "" Then
        CauseString = "very light condition."
        Call AddToStatus("Very light condition detected!  Pausing Action List.")
    End If
            
    NeedToPauseActionList = CloudSensorPause Or RainSensorPause Or WindSensorPause Or LightSensorPause
End Function

Private Sub PauseActionListAndPark(PauseReason As String)
    Call CameraAbort
    
    If Mount.TelescopeConnected And Not Mount.SimulatedPark Then
        Call MountAbort
        
        Do While Not objTele.IsSlewComplete
            Call Wait(1)
        Loop
        
        LastRA = Mount.CurrentRA
        LastDec = Mount.CurrentDec
        TargetName = Mount.CurrentTargetName
    Else
        LastRA = -1
        LastDec = 0
    End If
          
    If Rotator.RotatorConnected Then
        LastRotatorAngle = Rotator.CurrentAngle
    Else
        LastRotatorAngle = 0
    End If
          
    If frmOptions.chkParkMountWhenCloudy.Value = vbChecked Then
        If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled Then
            DomeUncoupled = True
        End If
            
        Call Mount.ParkMount
    End If

    If frmOptions.chkEnableWeatherMonitorScripts = vbChecked Then
        On Error Resume Next
        Call RunScriptDirect(frmOptions.txtAfterPauseScript.Text, True)
        On Error GoTo 0
    End If

    If frmOptions.chkEMailAlert(EMailAlertIndexes.WeatherMonitorActionListPaused).Value = vbChecked Then
        'Send e-mail!
        Call EMail.SendEMail(frmMain, "CCD Commander Weather Monitor - Action List Paused", "Action list paused due to " & PauseReason & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
    End If
    
    'Do this here so the ParkMount function can execute normally
    'The actions will not begin to run until after the timer routine exits
    'set the PauseBetweenActions flag
    PauseBetweenActions = True
    
    'Aborted is set later so the dome closure can complete successfully
    NeedToSetAbortFlag = True
End Sub

Private Function CheckGoodConditions(CloudStatus As CloudConditions, RainStatus As RainConditions, WindStatus As WindConditions, LightStatus As DayLightConditions) As Boolean
    Dim CloudSensorGood As Boolean
    Dim RainSensorGood As Boolean
    Dim WindSensorGood As Boolean
    Dim LightSensorGood As Boolean
    Dim Counter As Integer

    'first check if any options for the sensor are enabled
    CloudSensorGood = True
    For Counter = 0 To frmOptions.lstCloudSensorResumeActionWhen.ListCount - 1
        CloudSensorGood = CloudSensorGood And (Not frmOptions.lstCloudSensorResumeActionWhen.Selected(Counter))
    Next Counter
    'At this point, the flag will be true if all the checkboxes are unchecked.  Otherwise it will be false and the conditions will determine the actual state
    CloudSensorGood = CloudSensorGood Or ((CloudStatus = CloudConditions.Unknown) And frmOptions.lstCloudSensorResumeActionWhen.Selected(0))
    CloudSensorGood = CloudSensorGood Or ((CloudStatus = CloudConditions.Clear) And frmOptions.lstCloudSensorResumeActionWhen.Selected(1))
    CloudSensorGood = CloudSensorGood Or ((CloudStatus = CloudConditions.Cloudy) And frmOptions.lstCloudSensorResumeActionWhen.Selected(2))

    'first check if any options for the sensor are enabled
    RainSensorGood = True
    For Counter = 0 To frmOptions.lstRainSensorResumeActionWhen.ListCount - 1
        RainSensorGood = RainSensorGood And (Not frmOptions.lstRainSensorResumeActionWhen.Selected(Counter))
    Next Counter
    'At this point, the flag will be true if all the checkboxes are unchecked.  Otherwise it will be false and the conditions will determine the actual state
    RainSensorGood = RainSensorGood Or ((RainStatus = RainConditions.Unknown) And frmOptions.lstRainSensorResumeActionWhen.Selected(0))
    RainSensorGood = RainSensorGood Or ((RainStatus = RainConditions.Dry) And frmOptions.lstRainSensorResumeActionWhen.Selected(1))

     'first check if any options for the sensor are enabled
    WindSensorGood = True
    For Counter = 0 To frmOptions.lstWindSensorResumeActionWhen.ListCount - 1
        WindSensorGood = WindSensorGood And (Not frmOptions.lstWindSensorResumeActionWhen.Selected(Counter))
    Next Counter
    'At this point, the flag will be true if all the checkboxes are unchecked.  Otherwise it will be false and the conditions will determine the actual state
    WindSensorGood = WindSensorGood Or ((WindStatus = WindConditions.Unknown) And frmOptions.lstWindSensorResumeActionWhen.Selected(0))
    WindSensorGood = WindSensorGood Or ((WindStatus = WindConditions.Calm) And frmOptions.lstWindSensorResumeActionWhen.Selected(1))
    WindSensorGood = WindSensorGood Or ((WindStatus = WindConditions.Windy) And frmOptions.lstWindSensorResumeActionWhen.Selected(2))

     'first check if any options for the sensor are enabled
    LightSensorGood = True
    For Counter = 0 To frmOptions.lstLightSensorResumeActionWhen.ListCount - 1
        LightSensorGood = LightSensorGood And (Not frmOptions.lstLightSensorResumeActionWhen.Selected(Counter))
    Next Counter
    'At this point, the flag will be true if all the checkboxes are unchecked.  Otherwise it will be false and the conditions will determine the actual state
    LightSensorGood = LightSensorGood Or ((LightStatus = DayLightConditions.Unknown) And frmOptions.lstLightSensorResumeActionWhen.Selected(0))
    LightSensorGood = LightSensorGood Or ((LightStatus = DayLightConditions.Dark) And frmOptions.lstLightSensorResumeActionWhen.Selected(1))
    LightSensorGood = LightSensorGood Or ((LightStatus = DayLightConditions.Light) And frmOptions.lstLightSensorResumeActionWhen.Selected(2))
    
    'All must be true to resume
    CheckGoodConditions = CloudSensorGood And RainSensorGood And WindSensorGood And LightSensorGood
End Function
