Attribute VB_Name = "Mount"
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

Public objTele As Object

Public CurrentRA As Double
Public CurrentDec As Double
Public CurrentTargetJ2000 As Boolean
Public CurrentTargetJ2000RA As Double
Public CurrentTargetJ2000Dec As Double
Public CurrentTargetName As String

Public Enum typMountSide
    EastSide
    WestSide
    Unknown
End Enum

Public MountSide As typMountSide

Public Enum MountControl
    TheSky6 = 0
    ASCOM = 1
    TheSkyX = 2
End Enum

Public TelescopeConnected As Boolean

Private RAOffset As Double
Private DecOffset As Double

Public SimulatedParkAzim As Double
Public SimulatedParkAlt As Double
Public SimulatedPark As Boolean

Public Sub MoveRADecAction(clsAction As MoveRADecAction)
    Dim Counter As Integer
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    Dim MaxPointingError As Double
    Dim PreviousSide As typMountSide
    Dim DomeUncoupled As Boolean
    Dim SlewError As Long
    Dim MoveCount As Integer
    Dim MoveTimeOut As Integer
    Dim Retry As Boolean
    Dim DidRetryOnce As Boolean
    
    On Error GoTo MoveRADecActionError
    
    If Not Aborted Then
        Call ConnectToTelescope
    End If
    
    Call AddToStatus("Starting move to action.")
    
    Call Camera.CheckAndStopAutoguider
    
    SimulatedPark = False
    
    PreviousSide = MountSide
    
    CurrentRA = clsAction.RA
    CurrentDec = clsAction.Dec
    CurrentTargetName = clsAction.Name
    
    'Run the optional script
    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
        Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
    End If
    
    If DomeControl.DomeEnabled = True Then
        If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled And Not Aborted Then
            Call AddToStatus("Uncoupling Dome.")
            Call DomeControl.UnCoupleDome
            DomeUncoupled = True
        Else
            DomeUncoupled = False
        End If
    End If
    
    If clsAction.Epoch = J2000 Then
        CurrentTargetJ2000 = True
        CurrentTargetJ2000RA = clsAction.RA
        CurrentTargetJ2000Dec = clsAction.Dec
        
        Call AddToStatus("Precessing coordinates.")
        Call AddToStatus("J2000 Coordinates: " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False))
        Call Mount.PrecessCoordinates(CurrentRA, CurrentDec)
        Call AddToStatus("JNow Coordinates: " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False))
    Else
        CurrentTargetJ2000 = False
    
        If clsAction.RecomputeObjectCoordinates = 1 Then
            Call AddToStatus("Updating RA/Dec coordinates.")
            Call Planetarium.GetObjectRADec(clsAction.Name, CurrentRA, CurrentDec)
            Call AddToStatus("Updated position: " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False))
        End If
    End If
    
    For Counter = 1 To 2
        If frmOptions.optMountType(0).Value Then
            If (MountSide = WestSide) And (Misc.DoubleModulus(CurrentRA - objTele.LocalSiderealTime, 24) > 12) And _
                (Misc.DoubleModulus(CurrentRA - RAOfWesternLimit() + 0.5, 24) < 12) Then
                
                'force mount to flip
                Call AddToStatus("Flipping mount...")
                
                If Not Aborted Then _
                    Call objTele.SlewToRADec(Misc.DoubleModulus(RAOfWesternLimit() - 0.5, 24), CurrentDec + DecOffset, "")
                    
                If Not Aborted Then
                    Do While Not objTele.IsSlewComplete And Not Aborted
                        Call Wait(1)
                        If Aborted Then Exit Do
                    Loop
                End If
                
                If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
                    Call AddToStatus("Starting post slew delay.")
                    Call Wait(Settings.DelayAfterSlew)
                    Call AddToStatus("Completed post slew delay.")
                End If
                    
                MountSide = EastSide
            ElseIf (MountSide = EastSide) And (Misc.DoubleModulus(CurrentRA - objTele.LocalSiderealTime, 24) < 12) And _
                (Misc.DoubleModulus(CurrentRA - RAOfEasternLimit() - 0.5, 24) > 12) Then
                
                'force mount to flip
                Call AddToStatus("Flipping mount...")
                
                If Not Aborted Then _
                    Call objTele.SlewToRADec(Misc.DoubleModulus(RAOfEasternLimit() + 0.5, 24), CurrentDec + DecOffset, "")
                
                If Not Aborted Then
                    Do While Not objTele.IsSlewComplete And Not Aborted
                        Call Wait(1)
                        If Aborted Then Exit Do
                    Loop
                End If
                
                If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
                    Call AddToStatus("Starting post slew delay.")
                    Call Wait(Settings.DelayAfterSlew)
                    Call AddToStatus("Completed post slew delay.")
                End If
                
                MountSide = WestSide
            End If
        End If
        
        If Not Aborted Then
            If clsAction.Name = "" Then
                Call AddToStatus("Slewing to " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False) & " ...")
            Else
                Call AddToStatus("Slewing to " & clsAction.Name & "...")
            End If
        End If
        
        If Not Aborted Then
            Retry = True
            DidRetryOnce = False
            
            Do While Retry
                Retry = False
                
                On Error Resume Next
                
                Call objTele.SlewToRADec(Misc.DoubleModulus(CurrentRA + RAOffset, 24), CurrentDec + DecOffset, CurrentTargetName)
                
                If Err.Number = 1121 Then
                    On Error GoTo MoveRADecActionError
                    If DidRetryOnce Then
                        Call AddToStatus("Received Error 1121. Already retried once, giving up.")
                        GoTo MoveRADecActionError
                    Else
                        Call AddToStatus("Received Error 1121. Attempting to abort.")
                        Call objTele.Abort
                        Call AddToStatus("Abort completed. Waiting 5 seconds before trying slew again.")
                        Call Wait(5)
                        Call AddToStatus("Retrying...")
                        Retry = True
                        DidRetryOnce = True
                    End If
                ElseIf Err.Number <> 0 Then
                    GoTo MoveRADecActionError
                End If
            Loop
        End If
        
        If Not Aborted Then
            MoveCount = 0
            MoveTimeOut = CInt(GetMySetting("CustomSetting", "MoveTimeOut", "300"))
            Do While Not objTele.IsSlewComplete And MoveCount < MoveTimeOut And Not Aborted
                Call Wait(1)
                MoveCount = MoveCount + 1
                If Aborted Then Exit Do
            Loop
        End If
        
        If Not Aborted Then
            If MoveCount >= MoveTimeOut Then
                Call AddToStatus("Move timed-out without completing.  Aborting slew.")
                objTele.Abort
            End If
        End If
        
        If Not Aborted Then
            SlewError = objTele.LastError
            If SlewError <> 0 Then
                Call AddToStatus("Mount reports a slew error. Error = " & SlewError)
            End If
        End If
        
        If Not Aborted Then
            Call AddToStatus("Done slewing!")
            
            If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
                Call AddToStatus("Starting post slew delay.")
                Call Wait(Settings.DelayAfterSlew)
                Call AddToStatus("Completed post slew delay.")
            End If
        End If
        
        If Not Aborted Then
            If frmOptions.optMountType(0).Value Then
                If (Misc.DoubleModulus(objTele.RA - objTele.LocalSiderealTime, 24) < 12) And Not WithinEasternLimit(objTele.RA) Then
                    MountSide = WestSide
                ElseIf (Misc.DoubleModulus(objTele.RA - objTele.LocalSiderealTime, 24) > 12) And Not WithinWesternLimit(objTele.RA) Then
                    MountSide = EastSide
                End If
            
                Call SaveMySetting("MountParameters", "MountSide", CStr(MountSide))
            End If
        
            If frmOptions.chkVerifyTeleCoords.Value = vbChecked And Not Aborted Then
                MaxPointingError = Settings.MaxPointingError
                
                If Abs(Misc.DoubleModulus(objTele.RA - RAOffset, 24) - CurrentRA) < ((MaxPointingError / 60) / Cos(CurrentDec * PI / 180) / 15) And Abs((objTele.Dec - DecOffset) - CurrentDec) < (MaxPointingError / 60) Then
                    'looks good, get out of the loop
                    Exit For
                Else
                    'Something went wrong - first get new RA/Dec again, just in case the first set was wrong.
                    Call Wait(1)
                    
                    If Not Aborted Then
                        If Abs(Misc.DoubleModulus(objTele.RA - RAOffset, 24) - CurrentRA) > ((MaxPointingError / 60) / Cos(CurrentDec * PI / 180) / 15) Then
                            Call AddToStatus("Slew failed - final RA more then " & MaxPointingError & " arcminutes different.")
                            Call AddToStatus("Target coordinates are: " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False))
                            Call AddToStatus("Telescope reported coordinates are: " & Misc.ConvertEquatorialToString(objTele.RA, objTele.Dec, False))
                            Call AddToStatus("RA Offset = " & RAOffset & ", DEC Offset = " & DecOffset)
                            If Counter = 1 Then
                                Call AddToStatus("Waiting 5s and trying again.")
                                Call Wait(5)
                            End If
                        ElseIf Abs((objTele.Dec - DecOffset) - CurrentDec) > (MaxPointingError / 60) Then
                            Call AddToStatus("Slew failed - final Dec more then " & MaxPointingError & " arcminutes different.")
                            Call AddToStatus("Target coordinates are: " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False))
                            Call AddToStatus("Telescope reported coordinates are: " & Misc.ConvertEquatorialToString(objTele.RA, objTele.Dec, False))
                            Call AddToStatus("RA Offset = " & RAOffset & ", DEC Offset = " & DecOffset)
                            If Counter = 1 Then
                                Call AddToStatus("Waiting 5s and trying again.")
                                Call Wait(5)
                            End If
                        Else
                            'New RA/Dec must be good - get out of the loop
                            Exit For
                        End If
                    End If
                End If
            Else
                Exit For
            End If
        Else
            'aborted
            Exit For
        End If
    Next Counter
    
    If frmOptions.chkVerifyTeleCoords.Value = vbChecked And Not Aborted Then
        If Abs(Misc.DoubleModulus(objTele.RA - RAOffset, 24) - CurrentRA) > ((MaxPointingError / 60) / Cos(CurrentDec * PI / 180) / 15) Then
            Call AddToStatus("Skipping to next move action.")
            MainMod.SkipToNextMoveAction = True
        ElseIf Abs((objTele.Dec - DecOffset) - CurrentDec) > (MaxPointingError / 60) Then
            Call AddToStatus("Skipping to next move action.")
            MainMod.SkipToNextMoveAction = True
        End If
    End If
    
    If Not Aborted Then
        'Run the optional script
        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
            Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
        End If
    End If
    
    If DomeControl.DomeEnabled = True Then
        If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeUncoupled Then
            Call AddToStatus("Coupling Dome to Mount...")
            Call DomeControl.CoupleDome
            Call AddToStatus("Done coupling dome.")
        End If
    End If
    
    If Not Aborted Then
        If Rotator.RotatorConnected And PreviousSide <> MountSide And frmOptions.optMountType(0).Value Then
            If frmOptions.optRotatorFlip(0).Value Then
                ' Spin the rotator 180 degrees by keeping the same current angle
                Call Rotator.Rotate(Rotator.CurrentAngle)
            Else
                ' Maintain the same rotation by adding 180 degrees to the current angle
                Call Rotator.Rotate(Misc.DoubleModulus(Rotator.CurrentAngle + 180, 360))
            End If
        End If
    End If
    
    If TelescopeConnected Then
        If Aborted And Not objTele.IsSlewComplete Then
            objTele.Abort
        End If
    End If
    
    On Error GoTo 0
    Exit Sub
    
MoveRADecActionError:
    ErrorNum = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
        
    On Error GoTo 0
    
    If ErrorNum = &H800404C2 Or ErrorNum = 1218 Then
        'error is "out of limit" error
        'Skip all actions until next move action
        Call AddToStatus("Coordinates beyond slew limits.  Skipping to next move action.")
        Call AddToMissedTargetList(clsAction.Name, clsAction.RA, clsAction.Dec)
        MainMod.SkipToNextMoveAction = True
    ElseIf ErrorNum = &H80040401 Then
        'error is some strange one from NOVA direct drive system
        'will cause a retry of the move action.
        Call AddToStatus("Received error 0x80040401.")
        If (MainMod.RetryMoveAction) Then
            Call AddToStatus("Already retried.  Skipping to next move action.")
            MainMod.SkipToNextMoveAction = True
        Else
            MainMod.RetryMoveAction = True
        End If
    Else
        Call Err.Raise(ErrorNum, ErrorSource, ErrorDescription)
    End If
End Sub

Public Sub WaitForAltAction(clsAction As WaitForAltAction)
    Dim MyAlt As Double
    Dim MyAz As Double
    Dim Counter As Integer
    Dim MyTime As Date
        
    If clsAction.Name = "Sun" Then
        If clsAction.Rising Then
            Call AddToStatus("Waiting for " & clsAction.Name & " to rise to an altitude of " & clsAction.Alt & " degrees...")
            
            Call Mount.ComputeTwilightStartTime(clsAction.Alt)
            
            If CInt(Format(Time, "hh")) > 12 And CInt(Format(Mount.TwilightStartTime, "hh")) < 12 Then
                MyTime = Format(Date + 1, "Short Date") & " " & Mount.TwilightStartTime
            Else
                MyTime = Format(Date, "Short Date") & " " & Mount.TwilightStartTime
            End If
        Else
            Call AddToStatus("Waiting for " & clsAction.Name & " to set to an altitude of " & clsAction.Alt & " degrees...")
            
            Call Mount.ComputeSunSetTime(clsAction.Alt)
            
            If CInt(Format(Time, "hh")) > 12 And CInt(Format(Mount.SunSetTime, "hh")) < 12 Then
                MyTime = Format(Date + 1, "Short Date") & " " & Mount.SunSetTime
            Else
                MyTime = Format(Date, "Short Date") & " " & Mount.SunSetTime
            End If
        End If
    
        'Add 30s to time to account for calculation errors
        MyTime = DateAdd("s", 30, MyTime)
        
        Call AddToStatus("Waiting until " & Format(MyTime, "Short Time"))
        
        Do Until DateDiff("s", Now, MyTime) <= 0 Or Aborted
            Call Wait(60)
        Loop
    
    ElseIf clsAction.Name = "Moon" Then
        If clsAction.Rising Then
            Call AddToStatus("Waiting for " & clsAction.Name & " to rise to an altitude of " & clsAction.Alt & " degrees...")
            
            Call AstroFunctions.MoonRise(clsAction.Alt)
            MyTime = AstroFunctions.MoonRiseTime
        Else
            Call AddToStatus("Waiting for " & clsAction.Name & " to set to an altitude of " & clsAction.Alt & " degrees...")
            
            Call AstroFunctions.Moonset(clsAction.Alt)
            MyTime = AstroFunctions.MoonSetTime
        End If
    
        'Add 30s to time to account for calculation errors
        MyTime = DateAdd("s", 30, MyTime)
        
        Call AddToStatus("Waiting until " & Format(MyTime, "Short Time"))
        
        Do Until DateDiff("s", Now, MyTime) <= 0 Or Aborted
            Call Wait(60)
        Loop
        
    Else
        If clsAction.Name <> "" Then
            If clsAction.Rising Then
                Call AddToStatus("Waiting for " & clsAction.Name & " to rise to an altitude of " & clsAction.Alt & " degrees...")
            Else
                Call AddToStatus("Waiting for " & clsAction.Name & " to set to an altitude of " & clsAction.Alt & " degrees...")
            End If
        Else
            If clsAction.Rising Then
                Call AddToStatus("Waiting to rise to an altitude of " & clsAction.Alt & " degrees...")
            Else
                Call AddToStatus("Waiting to set to an altitude of " & clsAction.Alt & " degrees...")
            End If
        End If
    
        Counter = 10
        Call Misc.ConvertRADecToAltAz(clsAction.RA, clsAction.Dec, objTele.LocalSiderealTime, objTele.Latitude, MyAlt, MyAz)
        Do Until ((MyAlt >= clsAction.Alt Or MyAz >= 180) And clsAction.Rising) _
            Or ((MyAlt <= clsAction.Alt And MyAz >= 180) And (Not clsAction.Rising)) _
            Or Aborted
            
            If Counter = 10 Then
                Call AddToStatus("Current altitude is " & Format(MyAlt, "0.0") & " degrees...")
                Counter = 1
            End If
            
            Call Wait(60)
            
            Call Misc.ConvertRADecToAltAz(clsAction.RA, clsAction.Dec, objTele.LocalSiderealTime, objTele.Latitude, MyAlt, MyAz)
            
            Counter = Counter + 1
        Loop
    
        Call AddToStatus("Current altitude is " & Format(MyAlt, "0.0") & " degrees...")
    End If
    
    Call AddToStatus("Done waiting!")
End Sub

Public Sub WaitForTimeAction(clsAction As WaitForTimeAction)
    Dim AltAz() As Variant
    Dim Counter As Integer
    Dim MyTime As Date
    
    If clsAction.AbsoluteTime Then
        If clsAction.Hour < 12 And CInt(Format(Time, "hh")) >= 12 Then
            MyTime = (Format(Date + 1, "Short Date")) & " " & clsAction.Hour & ":" & clsAction.Minute & ":" & clsAction.Second
        ElseIf clsAction.Hour >= 12 And CInt(Format(Time, "hh")) < 12 Then
            MyTime = (Format(Date - 1, "Short Date")) & " " & clsAction.Hour & ":" & clsAction.Minute & ":" & clsAction.Second
        Else
            MyTime = Format(Date, "Short Date") & " " & clsAction.Hour & ":" & clsAction.Minute & ":" & clsAction.Second
        End If
    
        Call AddToStatus("Waiting until " & Format(MyTime, "h:nn:ss") & "...")
    Else
        MyTime = DateAdd("s", (CLng(clsAction.Hour) * 60 * 60) + (CLng(clsAction.Minute) * 60) + CLng(clsAction.Second), Now)
    
        Call AddToStatus("Waiting until " & Format(MyTime, "hh:nn:ss") & "...")
    End If
        
    Do Until DateDiff("s", Now, MyTime) <= 0 Or Aborted
        Call Wait(DateDiff("s", Now, MyTime))
    Loop

    Call AddToStatus("Done waiting!")
End Sub

Public Sub MountSetup()
    If frmOptions.lstMountControl.ListIndex = MountControl.TheSky6 Then
        If (objTele Is Nothing) Or (TypeName(objTele) = "ASCOMTelescopeControl") Or (TypeName(objTele) = "TheSkyXTelescopeControl") Then
            Set objTele = New TheSky6TelescopeControl
        End If
    ElseIf frmOptions.lstMountControl.ListIndex = MountControl.ASCOM Then
        If (objTele Is Nothing) Or (TypeName(objTele) = "TheSky6TelescopeControl") Or (TypeName(objTele) = "TheSkyXTelescopeControl") Then
            Set objTele = New ASCOMTelescopeControl
            
            'need to connect to do anything with ASCOM...
            Call Mount.ConnectToTelescope
        End If
    ElseIf frmOptions.lstMountControl.ListIndex = MountControl.TheSkyX Then
        If (objTele Is Nothing) Or (TypeName(objTele) = "TheSky6TelescopeControl") Or (TypeName(objTele) = "ASCOMTelescopeControl") Then
            Set objTele = New TheSkyXTelescopeControl
        End If
    End If
    
    RAOffset = 0
    DecOffset = 0
    
    SimulatedParkAlt = 0
    SimulatedParkAzim = 0
    SimulatedPark = False
    
    CurrentTargetJ2000 = False
End Sub

Public Sub MountUnload()
    Set objTele = Nothing
End Sub

Public Function WithinEasternLimit(RA As Double) As Boolean
    If (Misc.DoubleModulus(RA - objTele.LocalSiderealTime, 24) < (Settings.EasternLimit / 60)) Or _
        (Misc.DoubleModulus(objTele.LocalSiderealTime + 12 - RA, 24) < (Settings.EasternLimit / 60)) Then
        
        WithinEasternLimit = True
    Else
        WithinEasternLimit = False
    End If
End Function

Public Function WithinWesternLimit(RA As Double) As Boolean
    If (Misc.DoubleModulus(objTele.LocalSiderealTime - RA, 24) < (Settings.WesternLimit / 60)) Or _
        (Misc.DoubleModulus(RA - (objTele.LocalSiderealTime + 12), 24) < (Settings.WesternLimit / 60)) Then
        
        WithinWesternLimit = True
    Else
        WithinWesternLimit = False
    End If
End Function

Public Sub GetMountPosition()
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    Dim AutoSelect As Boolean
    
    Dim myCurrentRA As Double
    Dim myCurrentDec As Double
    
    Dim DomeUncoupled As Boolean
    
    If Not TelescopeConnected Then
        MsgBox "Internal Error!", vbCritical
        End
    End If
        
    If frmOptions.optMountType(0).Value Then
        If WithinWesternLimit(CurrentRA) Or WithinEasternLimit(CurrentRA) Then
            If objTele.SideOfPier = typMountSide.Unknown Then
                Call AddToStatus("Cannot tell which side of mount telescope is on.")
                
                If frmOptions.chkAutoDetermineMountSide.Value = vbUnchecked Then
                    frmMountSideSelect.Show vbModal, frmMain
                    
                    If frmMountSideSelect.Tag = "East" Then
                        AutoSelect = False
                        MountSide = EastSide
                        Call AddToStatus("User specified telescope on East side")
                    ElseIf frmMountSideSelect.Tag = "West" Then
                        AutoSelect = False
                        MountSide = WestSide
                        Call AddToStatus("User specified telescope on West side")
                    ElseIf frmMountSideSelect.Tag = "Auto" Then
                        AutoSelect = True
                    End If
                    
                    Unload frmMountSideSelect
                Else
                    AutoSelect = True
                End If
        
                If AutoSelect Then
                    'Run the optional script
                    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
                        Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
                    End If
                    
                    If DomeControl.DomeEnabled = True Then
                        If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled And Not Aborted Then
                            Call DomeControl.UnCoupleDome
                            DomeUncoupled = True
                        Else
                            DomeUncoupled = False
                        End If
                    End If
                    
                    'cannot tell which side the mount is on
                    If objTele.Az >= 0 And objTele.Az < 180 Then
                        'Azimuth says scope should be on west side looking east
                        myCurrentRA = objTele.RA
                        myCurrentDec = objTele.Dec
                    
                        Call AddToStatus("Forcing to western side...")
                        
                        'force mount to flip
                        Call objTele.SlewToRADec(RAOfEasternLimit() + 0.5, myCurrentDec, "")
                        Do While Not objTele.IsSlewComplete And Not Aborted
                            Call Wait(1)
                        Loop
                        
                        If Settings.DelayAfterSlew > 0 Then
                            Call AddToStatus("Starting post slew delay.")
                            Call Wait(Settings.DelayAfterSlew)
                            Call AddToStatus("Completed post slew delay.")
                        End If
                        
                        MountSide = WestSide
                    ElseIf objTele.Az >= 180 And objTele.Az < 360 Then
                        'Azimuth says scope should be on east side looking west
                        myCurrentRA = objTele.RA
                        myCurrentDec = objTele.Dec
                        
                        Call AddToStatus("Forcing to eastern side...")
                        
                        'force mount to flip
                        Call objTele.SlewToRADec(RAOfWesternLimit() - 0.5, myCurrentDec, "")
                        Do While Not objTele.IsSlewComplete And Not Aborted
                            Call Wait(1)
                        Loop
                        
                        If Settings.DelayAfterSlew > 0 Then
                            Call AddToStatus("Starting post slew delay.")
                            Call Wait(Settings.DelayAfterSlew)
                            Call AddToStatus("Completed post slew delay.")
                        End If
                                
                        MountSide = EastSide
                    End If
                    
                    Call AddToStatus("Moving to original location...")
                    
                    On Error GoTo MoveOriginalLocationError
                    'slew back to original location
                    Call objTele.SlewToRADec(myCurrentRA, myCurrentDec, "")
                    Do While Not objTele.IsSlewComplete And Not Aborted
                        Call Wait(1)
                    Loop
                    
                    If Settings.DelayAfterSlew > 0 Then
                        Call AddToStatus("Starting post slew delay.")
                        Call Wait(Settings.DelayAfterSlew)
                        Call AddToStatus("Completed post slew delay.")
                    End If
                
                    'Run the optional script
                    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
                        Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
                    End If
                    
                    If DomeControl.DomeEnabled = True Then
                        If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeUncoupled Then
                            Call DomeControl.CoupleDome
                        End If
                    End If
                    
MoveOriginalLocationError:
                    ErrorNum = Err.Number
                    ErrorSource = Err.Source
                    ErrorDescription = Err.Description
                        
                    On Error GoTo 0
                    
                    If ErrorNum = &H800404C2 Then
                        'error is "out of limit" error
                        'Skip all actions until next move action
                        Call AddToStatus("Original location beyond slew limits.  Leaving mount where it is.")
                    ElseIf ErrorNum <> 0 Then
                        Call Err.Raise(ErrorNum, ErrorSource, ErrorDescription)
                    End If
                End If
            Else
                MountSide = objTele.SideOfPier
            End If
        ElseIf Misc.DoubleModulus(CurrentRA - objTele.LocalSiderealTime, 24) < 12 Then
            MountSide = WestSide
        ElseIf Misc.DoubleModulus(CurrentRA - objTele.LocalSiderealTime, 24) > 12 Then
            MountSide = EastSide
        End If
        
        Call SaveMySetting("MountParameters", "MountSide", CStr(MountSide))
    Else
        'fork mount - don't care
    End If
End Sub

Public Sub MountAbort()
    On Error Resume Next
    Call objTele.Abort
    On Error GoTo 0
End Sub

Public Sub Sync(ActualRA As Double, ActualDec As Double, Slew As Boolean)
    Dim DomeUncoupled As Boolean
    
    Call ConnectToTelescope
    
    Call AddToStatus("Syncing to " & Misc.ConvertEquatorialToString(ActualRA, ActualDec, False))
    Call objTele.Sync(ActualRA, ActualDec)
    
    RAOffset = 0
    DecOffset = 0
    
    If Slew Then
        Call Camera.CheckAndStopAutoguider
        
        'Run the optional script
        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
            Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
        End If
    
        If DomeControl.DomeEnabled = True Then
            If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled And Not Aborted Then
                Call DomeControl.UnCoupleDome
                DomeUncoupled = True
            Else
                DomeUncoupled = False
            End If
        End If
        
        'now go to the original RA/Dec coordinates again
        Call AddToStatus("Reslewing to original target coordinates...")
        Call objTele.SlewToRADec(CurrentRA, CurrentDec, CurrentTargetName)
        Do
            Call Wait(1)
        Loop While Not objTele.IsSlewComplete And Not Aborted
        If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
            Call AddToStatus("Starting post slew delay.")
            Call Wait(Settings.DelayAfterSlew)
            Call AddToStatus("Completed post slew delay.")
        End If
        
        'Run the optional script
        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
            Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
        End If
        
        If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.DomeEnabled = True And DomeUncoupled Then
            Call DomeControl.CoupleDome
        End If
        
    Else
        CurrentRA = ActualRA
        CurrentDec = ActualDec
    End If
End Sub

Public Sub OffsetCoordinates(ActualRA As Double, ActualDec As Double, Slew As Boolean)
    Dim NewRA As Double
    Dim NewDec As Double
    
    Dim DomeUncoupled As Boolean
    
    Call ConnectToTelescope
    
    RAOffset = (objTele.RA - ActualRA)
    DecOffset = (objTele.Dec - ActualDec)
    
    Call AddToStatus("Setting offset to RA " & Format(RAOffset, "0.0e+0") & ", Dec " & Format(DecOffset, "0.0e+0"))
    
    If Slew Then
        Call Camera.CheckAndStopAutoguider
        
        'Run the optional script
        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
            Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
        End If
        
        If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.DomeEnabled = True And DomeControl.IsDomeCoupled And Not Aborted Then
            Call DomeControl.UnCoupleDome
            DomeUncoupled = True
        Else
            DomeUncoupled = False
        End If
        
        'now go to the original RA/Dec coordinates again
        Call AddToStatus("Reslewing to original target coordinates...")
        Call objTele.SlewToRADec(CurrentRA + RAOffset, CurrentDec + DecOffset, CurrentTargetName)
        Do
            Call Wait(1)
        Loop While Not objTele.IsSlewComplete And Not Aborted
        If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
            Call AddToStatus("Starting post slew delay.")
            Call Wait(Settings.DelayAfterSlew)
            Call AddToStatus("Completed post slew delay.")
        End If
    
        'Run the optional script
        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
            Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
        End If
        
        If DomeControl.DomeEnabled = True Then
            If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeUncoupled Then
                Call DomeControl.CoupleDome
            End If
        End If
    End If
        
End Sub

Public Function FlipAtMeridian(ImageLink As Boolean, ImageLinkAction As ImageLinkSyncAction, RotateAfterFlip As Integer, DoubleImageLink As Boolean) As Boolean
    Dim lCurrentRA As Double
    Dim lCurrentDec As Double
    Dim ImageLinkStatus As Boolean
    Dim DomeUncoupled As Boolean
    
    Call ConnectToTelescope
    
    'ok, check to see if I can even flip yet
    Do While (Misc.DoubleModulus(CurrentRA - Mount.RAOfEasternLimit, 24) < 12) And Not Aborted
        'not even into the overlap zone yet
        'need to wait before I can flip
        'wait for 10% of the total overlap zone
        Call AddToStatus("Cannot flip the mount yet!")
        Call AddToStatus("Waiting for mount to enter the meridian zone...")
        Call Wait((Misc.DoubleModulus(CurrentRA - RAOfEasternLimit(), 24) * 3600) + 60)
    Loop
    
    'get the current RA/Dec of the mount
    lCurrentRA = CurrentRA + RAOffset
    lCurrentDec = CurrentDec + DecOffset
    
    Call AddToStatus("Flipping....")
    
    'Run the optional script
    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
        Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
    End If
    
    If DomeControl.DomeEnabled = True Then
        If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled And Not Aborted Then
            Call DomeControl.UnCoupleDome
            DomeUncoupled = True
        Else
            DomeUncoupled = False
        End If
    End If
    
    If Not Aborted Then _
        Call objTele.SlewToRADec(Misc.DoubleModulus(RAOfWesternLimit() - 0.5, 24), lCurrentDec, "")
    
    If Not Aborted Then
        Do
            Call Wait(1)
            If Aborted Then Exit Do
        Loop While Not objTele.IsSlewComplete And Not Aborted
    End If
    
    If Settings.DelayAfterSlew > 0 Then
        Call AddToStatus("Starting post slew delay.")
        Call Wait(Settings.DelayAfterSlew)
        Call AddToStatus("Completed post slew delay.")
    End If

    If Not Aborted Then
        Call AddToStatus("Recentering....")
        'now go to the original RA/Dec coordinates
        If Not Aborted Then _
            Call objTele.SlewToRADec(lCurrentRA, lCurrentDec, CurrentTargetName)
        
        If Not Aborted Then
            Do
                Call Wait(1)
                If Aborted Then Exit Do
            Loop While Not objTele.IsSlewComplete And Not Aborted
        End If
    End If
    
    If Settings.DelayAfterSlew > 0 Then
        Call AddToStatus("Starting post slew delay.")
        Call Wait(Settings.DelayAfterSlew)
        Call AddToStatus("Completed post slew delay.")
    End If
        
    'Flip complete - now on the east side of the monut
    MountSide = EastSide
    
    'Rotate camera if needed
    If Not Aborted And Rotator.RotatorConnected Then
        If RotateAfterFlip = 1 Then
            'Rotate 180 degrees by maintaining position angle
            Call Rotator.Rotate(Rotator.CurrentAngle)
        Else
            'Keep rotator where it is, but update values by changing current angle by 180 degrees
            Call Rotator.Rotate(Misc.DoubleModulus(Rotator.CurrentAngle + 180, 360))
        End If
    End If
    
    'Run the optional script
    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
        Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
    End If
    
    If DomeControl.DomeEnabled = True Then
        If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeUncoupled Then
            Call DomeControl.CoupleDome
        End If
    End If
    
    'check alignment
    If ImageLink And Not Aborted Then
        ImageLinkStatus = TakeImageAndLink(ImageLinkAction)
        
        If DoubleImageLink And Not Aborted Then
            ImageLinkStatus = TakeImageAndLink(ImageLinkAction) 'double link to fix backlash problems.
        End If
    Else
        ImageLinkStatus = True
    End If
    
    If Not Aborted Then
        Call AddToStatus("Flip complete.")
        Call SaveMySetting("MountParameters", "MountSide", CStr(MountSide))
    End If
    
    FlipAtMeridian = ImageLinkStatus
    
    'shouldn't have changed, but just in case
    CurrentRA = lCurrentRA
    CurrentDec = lCurrentDec
End Function

Public Function RAOfWesternLimit() As Double
    RAOfWesternLimit = Misc.DoubleModulus(objTele.LocalSiderealTime - (Settings.WesternLimit / 60), 24)
End Function

Public Function RAOfEasternLimit() As Double
    RAOfEasternLimit = Misc.DoubleModulus(objTele.LocalSiderealTime + (Settings.EasternLimit / 60), 24)
End Function

Public Sub ParkMount(Optional clsAction As ParkMountAction)
    Dim DomeUncoupled As Boolean
    Dim ParkRotator As Boolean
    Dim StoppedTracking As Boolean
    
    StoppedTracking = False
    
    Call ConnectToTelescope
    
    Call AddToStatus("Parking mount...")
    
    SimulatedPark = False
    
    Call Camera.CheckAndStopAutoguider
        
    If Not Exiting Then
        'Run the optional script
        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
            Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
        End If
                
        If DomeControl.DomeEnabled Then
            If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled And Not Aborted Then
                Call DomeControl.UnCoupleDome
                DomeUncoupled = True
            Else
                DomeUncoupled = False
            End If
        Else
            DomeUncoupled = False
        End If
        
        If clsAction Is Nothing Then
            Call objTele.Park
            'ParkRotator = True
            ParkRotator = False
        Else
            If clsAction.DoSimulatedPark Then
                If (clsAction.AltD < 0) Then
                    SimulatedParkAlt = clsAction.AltD - (clsAction.AltM / 60) - (clsAction.AltS / 3600)
                Else
                    SimulatedParkAlt = clsAction.AltD + (clsAction.AltM / 60) + (clsAction.AltS / 3600)
                End If
                SimulatedParkAzim = clsAction.AzimD + (clsAction.AzimM / 60) + (clsAction.AzimS / 3600)
                Call Mount.MoveToAltAz(SimulatedParkAlt, SimulatedParkAzim)
            ElseIf clsAction.DoHomePark Then
                Call objTele.Home
            ElseIf Not clsAction.DoTrackingOff Then
                Call objTele.Park
            End If
            
            ParkRotator = clsAction.ParkRotator
        End If
        
        On Error Resume Next
        Do
            Call Wait(5)
            If Aborted Then Exit Do
        Loop While Not objTele.IsSlewComplete And Not Aborted
            'IsSlewComplete will return an error when the park is finished since TheSky will disconnect.
            'Simply trap the error and continue
        On Error GoTo 0
        
        If Not (clsAction Is Nothing) Then
            If clsAction.DoSimulatedPark Then
                If Mount.objTele.CanSetTracking Then
                    'Slew one more time to get closer to the desired position
                    Call AddToStatus("Starting second slew to Alt/Az to get closer to desired position.")
                    Call Mount.MoveToAltAz(SimulatedParkAlt, SimulatedParkAzim)
                    Do
                        Call Wait(1)
                        If Aborted Then Exit Do
                    Loop While Not objTele.IsSlewComplete And Not Aborted
                    
                    Mount.objTele.Tracking = False
                    Call AddToStatus("Tracking disabled.")
                    
                    StoppedTracking = True   'Am able to do a real park at this location
                End If
            ElseIf clsAction.DoHomePark Or clsAction.DoTrackingOff Then
                If Mount.objTele.CanSetTracking Then
                    Mount.objTele.Tracking = False
                    Call AddToStatus("Tracking disabled.")
                    
                    StoppedTracking = True   'Am able to do a real park at this location
                End If
            End If
        End If
                
        'Scope is done parking, park the Rotator if enabled
        If frmOptions.lstRotator.ListIndex > 0 And ParkRotator Then
            Call Rotator.Rotate(Settings.HomeRotationAngle)
        End If
                
        'Run the optional script
        'Disable this for now.  Might make sense to not do this in the Park function
'        If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
'            Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
'        End If
'        If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.DomeEnabled = True And DomeUncoupled Then
'            Call DomeControl.CoupleDome
'        End If
               
        Call AddToStatus("Done parking!")
    End If
    
    If clsAction Is Nothing Then
        Call Camera.objCameraControl.DisconnectFromTelescope
        
        TelescopeConnected = False
    Else
        If Not clsAction.DoSimulatedPark Or StoppedTracking Then
            Call Camera.objCameraControl.DisconnectFromTelescope
            TelescopeConnected = False
            SimulatedPark = False
        Else
            SimulatedPark = True
        End If
    End If
End Sub

Public Sub MoveToAltAz(Alt As Double, Az As Double)
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    Dim DomeUncoupled As Boolean
    Dim RA As Double
    Dim Dec As Double
    
    On Error GoTo MoveAltAzError
    
    Call ConnectToTelescope
    
    SimulatedPark = False
    
    Call AddToStatus("Moving to Alt-Az position: " & Misc.ConvertAltAzToString(Alt, Az, False))
    
    Call Camera.CheckAndStopAutoguider
        
    'Run the optional script
    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
        Call RunScriptDirect(frmOptions.txtBeforeScript.Text, True)
    End If
            
    If DomeControl.DomeEnabled = True Then
        If frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeControl.IsDomeCoupled And Not Aborted Then
            Call DomeControl.UnCoupleDome
            DomeUncoupled = True
        Else
            DomeUncoupled = False
        End If
    Else
        DomeUncoupled = False
    End If
               
    Call Misc.ConvertAltAzToRADec(Alt, Az, objTele.LocalSiderealTime, objTele.Latitude, RA, Dec)
               
    If Not Aborted Then _
        Call objTele.SlewToRADec(RA, Dec, "")
    
    If Not Aborted Then
        Do
            Call Wait(1)
            If Aborted Then Exit Do
        Loop While Not objTele.IsSlewComplete And Not Aborted
    End If
    
    If frmOptions.optMountType(0).Value Then
        If (Misc.DoubleModulus(objTele.RA - objTele.LocalSiderealTime, 24) < 12) And Not WithinEasternLimit(objTele.RA) Then
            MountSide = WestSide
        ElseIf (Misc.DoubleModulus(objTele.RA - objTele.LocalSiderealTime, 24) > 12) And Not WithinWesternLimit(objTele.RA) Then
            MountSide = EastSide
        End If
    
        Call SaveMySetting("MountParameters", "MountSide", CStr(MountSide))
    End If
    
    If Not Aborted Then
        Call AddToStatus("Done slewing!")
        
        If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
            Call AddToStatus("Starting post slew delay.")
            Call Wait(Settings.DelayAfterSlew)
            Call AddToStatus("Completed post slew delay.")
        End If
    End If
            
    'Run the optional script
    If frmOptions.chkEnableScripts = vbChecked And Not Aborted Then
        Call RunScriptDirect(frmOptions.txtAfterScript.Text, True)
    End If
                
    If DomeControl.DomeEnabled = True Then
        If Not Aborted And frmOptions.chkUncoupleDomeDuringSlews = vbChecked And DomeUncoupled Then
            Call DomeControl.CoupleDome
        End If
    End If
    
MoveAltAzError:
    ErrorNum = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
        
    On Error GoTo 0
    
    If ErrorNum = &H800404C2 Or ErrorNum = 1218 Then
        'error is "out of limit" error
        'Skip all actions until next move action
        Call AddToStatus("Coordinates beyond slew limits.  Aborting.")
        Aborted = True
    ElseIf ErrorNum <> 0 Then
        Call Err.Raise(ErrorNum, ErrorSource, ErrorDescription)
    End If
End Sub

Public Sub RecenterAltAz(Alt As Double, Az As Double, Optional WaitForComplete As Boolean = True)
    Dim RA As Double
    Dim Dec As Double
    
    'Not a full featured as the Move routine above.  This should just be a short jog, so don't need to do as much
    
    If Not objTele.IsSlewComplete Then
        'cannot move just yet - mount move is in progress
        'just exit out - will get the move in the next iteration
        Exit Sub
    End If
    
    Call Misc.ConvertAltAzToRADec(Alt, Az, objTele.LocalSiderealTime, objTele.Latitude, RA, Dec)
               
    If Not Aborted Then _
        Call objTele.SlewToRADec(RA, Dec, "")
                
    If Not Aborted And WaitForComplete Then
        Do While Not objTele.IsSlewComplete And Not Aborted
            Call Wait(1)
            If Aborted Then Exit Do
        Loop
    End If
End Sub

Public Sub GetTelescopeRADec(RA As Double, Dec As Double)
    'ensure I'm connected
    If objTele Is Nothing Then
        Call Mount.MountSetup
    End If
    
    objTele.ConnectToMount
    
    RA = objTele.RA
    Dec = objTele.Dec
End Sub

Public Sub GetTelescopeAltAz(Alt As Double, Az As Double)
    Call Misc.ConvertRADecToAltAz(Mount.CurrentRA, Mount.CurrentDec, objTele.LocalSiderealTime, objTele.Latitude, Alt, Az)
    
'    On Error Resume Next
'    Alt = objTele.Alt
'    Az = objTele.Az
'    On Error GoTo 0
End Sub

Public Sub ConnectToTelescope()
    If Not TelescopeConnected And Not Aborted Then
    
        objTele.ConnectToMount
        TelescopeConnected = True
        SimulatedPark = False
    
        Do
            Call Wait(1)
        
            If Not Aborted Then
                CurrentRA = objTele.RA
                CurrentDec = objTele.Dec
                CurrentTargetName = ""
            End If
        Loop While CurrentRA = 0 And CurrentDec = 0 And Not Aborted
        
        If Not Aborted Then _
            Call GetMountPosition
    End If
End Sub

Public Sub PrecessCoordinates(RA As Double, Dec As Double)
    Dim RANow As Double
    Dim DecNow As Double
    
    Call Misc.PrecessCoordinates(RA, Dec, RANow, DecNow)
    
    RA = RANow
    Dec = DecNow
End Sub

Public Function GetTelescopeType() As String
    GetTelescopeType = objTele.TelescopeType
End Function

Public Sub JogTelescope(RAJog As Double, DecJog As Double)
    CurrentRA = Misc.DoubleModulus(CurrentRA + RAJog, 24)
    CurrentDec = CurrentDec + DecJog
    
    Call AddToStatus("Moving mount to " & Misc.ConvertEquatorialToString(CurrentRA, CurrentDec, False))
    
    If Not Aborted Then
        Call objTele.SlewToRADec(Misc.DoubleModulus(CurrentRA + RAOffset, 24), CurrentDec + DecOffset, "")
    End If
    
    If Not Aborted Then
        Do While Not objTele.IsSlewComplete And Not Aborted
            Call Wait(1)
            If Aborted Then Exit Do
        Loop
    End If
    
    If Not Aborted Then
        Call AddToStatus("Done slewing!")
        
        If Settings.DelayAfterSlew > 0 And frmOptions.chkOnlyDelayAfterMeridianFlip = vbUnchecked And Not Aborted Then
            Call AddToStatus("Starting post slew delay.")
            Call Wait(Settings.DelayAfterSlew)
            Call AddToStatus("Completed post slew delay.")
        End If
    End If
End Sub

Public Function GetLatitude() As Double
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                
    GetLatitude = objTele.Latitude
End Function

Public Function GetLongitude() As Double
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                
    GetLongitude = objTele.Longitude
End Function

Public Function GetSiderealTime() As Double
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                
    GetSiderealTime = objTele.LocalSiderealTime
End Function

Public Sub ComputeSunSetTime(Altitude As Double)
    'Moved to AstroFunctions
    Call AstroFunctions.ComputeSunSetTime(Altitude)
End Sub

Public Sub ComputeTwilightStartTime(Altitude As Double)
    'Moved to AstroFunctions
    Call AstroFunctions.ComputeTwilightStartTime(Altitude)
End Sub

Public Property Get SunSetTime() As Date
    'Moved to AstroFunctions
    SunSetTime = AstroFunctions.SunSetTime
End Property

Public Property Get TwilightStartTime() As Date
    'Moved to AstroFunctions
    TwilightStartTime = AstroFunctions.TwilightStartTime
End Property
