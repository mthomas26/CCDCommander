Attribute VB_Name = "Camera"
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

Public AutoguiderGuideErrorEvent As Boolean
Public AutoguiderGuideFailedEvent As Boolean
Public Autoguiding As Boolean
Public UseAGBeforeTakeImageEvent As Boolean

Public objCameraControl As Object
'Public objCameraControl As clsCCDSoftControl

Public Enum CameraControl
    CCDSoft = 0
    MaxIm = 1
    CCDSoftAO = 2
    TheSkyX = 3
End Enum

Public Enum ReductionType
    NoReduction
    AutoDark
    FullReduction
End Enum

Public Const MaxAutoFlatFilters = 20
Public Const MaxAutoFlatRotations = 20

Public ImageTypes() As Variant

Private LastGuideStarX As Double
Private LastGuideStarY As Double
Private StartRA As Double
Private StartDec As Double
Private StartMountSide As typMountSide
    
Public PixelSize As Double
    
Public Sub ImagerAction(clsAction As ImagerAction)
    Dim TotalExposures As Long
    Dim SeriesLength As Long
    Dim ImageSeriesRemaining As Long
    Dim DitherCounter As Long
    Dim AbortThisAction As Boolean
    Dim Counter As Integer
    Dim GuideStarX As Double
    Dim GuideStarY As Double
    Dim TimeToWesternLimit As Double
    Dim frmMainActionText As String
    Dim UnguidedDitherRA As Double
    Dim UnguidedDitherDec As Double
    Dim GuideStarRecoveryStatus As Boolean
    Dim NeedToFlip As Boolean
    Dim StarFadedCount As Integer
    Dim StartedAutoguiderAlready As Boolean
    Dim SaveToPath As String
    Static DitherIndexX As Integer
    Static DitherIndexY As Integer
    Static DitherIncLimit As Integer
    Static DitherIncAxis As Integer
    Static DitherIncCount As Integer
    Static DitherIncDir As Integer
    Static LastFilterIndex As Integer
    Static LastAGExposure As Double
    Static LastAGBinning As Integer
    Static LastDitherCounter As Integer
    
    Dim AutoguiderEnabled As Boolean
    
    Dim GuiderErrorExceededLimitCount As Integer
    Dim LastGuideErrorEventTime As Date
    
    If (clsAction.AutoguiderEnabled = vbChecked) Then
        AutoguiderEnabled = True
    Else
        AutoguiderEnabled = False
    End If
    
    AbortThisAction = False
    StartedAutoguiderAlready = False
    
    If ((StartRA <> Mount.CurrentRA Or StartDec <> Mount.CurrentDec) And Not Mount.CurrentTargetJ2000) Or _
        ((StartRA <> Mount.CurrentTargetJ2000RA Or StartDec <> Mount.CurrentTargetJ2000Dec) And Mount.CurrentTargetJ2000) Or _
        ((StartMountSide <> Mount.MountSide) And (frmOptions.chkInternalGuider.Value <> vbChecked)) Or _
        (LastAGBinning <> clsAction.AutoguiderBin) Or _
        ((LastFilterIndex <> clsAction.ImagerFilter) And (frmOptions.chkInternalGuider.Value = vbChecked) And (frmOptions.chkEnable.Value = vbChecked)) Or _
        (((LastGuideStarX <> clsAction.AutoguiderXPos) Or (LastGuideStarY <> clsAction.AutoguiderYPos)) And (frmOptions.chkEnable.Value = vbUnchecked)) Then
        
        If Mount.CurrentTargetJ2000 Then
            StartRA = Mount.CurrentTargetJ2000RA
            StartDec = Mount.CurrentTargetJ2000Dec
        Else
            StartRA = Mount.CurrentRA
            StartDec = Mount.CurrentDec
        End If
        
        StartMountSide = Mount.MountSide
        
        LastAGBinning = clsAction.AutoguiderBin
        
        LastFilterIndex = clsAction.ImagerFilter
        DitherIndexX = 0
        DitherIndexY = 0
        DitherIncLimit = 1
        DitherIncAxis = 0   '0=X, 1=Y
        DitherIncCount = 0
        DitherIncDir = 1
        LastGuideStarX = 0
        LastGuideStarY = 0
        LastAGExposure = 0
        LastDitherCounter = 0
    ElseIf (LastFilterIndex <> clsAction.ImagerFilter) And (clsAction.MaintainDitherOnFilterChange <> vbChecked) Then
        'reset dithering for the new filter
        'but still try for the same guide star since we aren't using an internal guider
        LastFilterIndex = clsAction.ImagerFilter
        DitherIndexX = 0
        DitherIndexY = 0
        DitherIncLimit = 1
        DitherIncAxis = 0   '0=X, 1=Y
        DitherIncCount = 0
        DitherIncDir = 1
        LastDitherCounter = 0
    ElseIf (LastFilterIndex <> clsAction.ImagerFilter) And (clsAction.MaintainDitherOnFilterChange = vbChecked) Then
        'just reset last filter index
        LastFilterIndex = clsAction.ImagerFilter
    End If
    
    objCameraControl.ImageType = clsAction.ImagerType
    Call AddToStatus("Setting image type to " & ImageTypes(clsAction.ImagerType - 1) & ".")
    
    objCameraControl.BinXY = clsAction.ImagerBin + 1
    Call AddToStatus("Setting imager bin mode to " & clsAction.ImagerBin + 1 & "x" & clsAction.ImagerBin + 1 & ".")
    
    If clsAction.ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
        Call AddToStatus("Setting filter to " & frmOptions.lstFilters.List(clsAction.ImagerFilter) & ".")
        
        'take a dummy exposure to get things setup
        Call Camera.ForceFilterChange(clsAction.ImagerFilter)
    End If
    
    If Not Aborted Then
        If clsAction.FrameSize = FullFrame Then
            Call AddToStatus("Setting imager to full frame.")
            Call objCameraControl.SubFrame(False)
        ElseIf clsAction.FrameSize = HalfFrame Then
            Call AddToStatus("Setting imager to half frame.")
            With objCameraControl
                Call .SubFrame(True, .WidthInPixels / 4, .WidthInPixels * 3 / 4, .HeightInPixels / 4, .HeightInPixels * 3 / 4)
            End With
        ElseIf clsAction.FrameSize = QuarterFrame Then
            With objCameraControl
                Call AddToStatus("Setting imager to quarter frame.")
                Call .SubFrame(True, .WidthInPixels * 3 / 8, .WidthInPixels * 5 / 8, .HeightInPixels * 3 / 8, .HeightInPixels * 5 / 8)
            End With
        ElseIf clsAction.FrameSize = CustomFrame Then
            With objCameraControl
                If (.WidthInPixels < clsAction.CustomFrameWidth) Then
                    Call AddToStatus("Custom Frame Width larger than imager width.  Resetting frame width to maximum.")
                    clsAction.CustomFrameWidth = .WidthInPixels
                End If
                
                If (.HeightInPixels < clsAction.CustomFrameHeight) Then
                    Call AddToStatus("Custom Frame Height larger than imager height.  Resetting frame height to maximum.")
                    clsAction.CustomFrameHeight = .HeightInPixels
                End If
                
                Call AddToStatus("Setting imager to custom frame size of " & clsAction.CustomFrameWidth & " by " & clsAction.CustomFrameHeight & ".")
                
                Call .SubFrame(True, (.WidthInPixels - clsAction.CustomFrameWidth) / 2, .WidthInPixels - (.WidthInPixels - clsAction.CustomFrameWidth) / 2, _
                                (.HeightInPixels - clsAction.CustomFrameHeight) / 2, .HeightInPixels - (.HeightInPixels - clsAction.CustomFrameHeight) / 2)
            End With
        End If
    End If
    
    If Not Aborted Then
        If ImageTypes(clsAction.ImagerType - 1) <> "Bias" Then
            objCameraControl.ExposureTime = clsAction.ImagerExpTime
            Call AddToStatus("Setting imager exposure time to " & clsAction.ImagerExpTime & " seconds.")
        Else
            clsAction.ImagerExpTime = 1
        End If
        
        If clsAction.CalibrateImages = vbChecked Then
            objCameraControl.ImageReduction = FullReduction
        Else
            objCameraControl.ImageReduction = NoReduction
        End If
        
        SeriesLength = 1
        TotalExposures = clsAction.ImagerNumExp
    End If
    
    If Not Aborted And Not (Autoguiding And frmOptions.chkContinuousAutoguide.Value = vbChecked) Or (Not AutoguiderEnabled) Then
        Autoguiding = False
        Call objCameraControl.StopAutoguider
        Call Wait(1)
    End If
    
    If Not Aborted And AutoguiderEnabled Then
        'Should leave the Autoguider filter alone - AG could be a seperate camera and don't want to change it!
        If frmOptions.chkInternalGuider.Value = vbChecked Then
            If clsAction.ImagerFilter >= 0 Then
                objCameraControl.AGFilterNumber = clsAction.ImagerFilter
            End If
        End If
        
        objCameraControl.AGBinXY = clsAction.AutoguiderBin + 1
        Call AddToStatus("Setting autoguider bin mode to " & clsAction.AutoguiderBin + 1 & "x" & clsAction.AutoguiderBin + 1 & ".")
    End If
            
    'only need to check this if I have a GEM
    If Not Aborted And AutoguiderEnabled And frmOptions.optMountType(0).Value And objCameraControl.ReverseXNecessary Then
        If Mount.MountSide = EastSide And frmOptions.optGuiderCal(0).Value And Not (Rotator.RotatorConnected And (frmOptions.chkGuiderRotates.Value = vbChecked)) Then
            objCameraControl.ReverseX = True
        ElseIf Mount.MountSide = WestSide And frmOptions.optGuiderCal(1).Value And Not (Rotator.RotatorConnected And (frmOptions.chkGuiderRotates.Value = vbChecked)) Then
            objCameraControl.ReverseX = True
        Else
            objCameraControl.ReverseX = False
        End If
    End If
    
    NeedToFlip = False
    'If connected to telescope and I have a GEM
    If Mount.TelescopeConnected And frmOptions.optMountType(0).Value Then
        'maybe need to do a flip...first check if I can do another exposre before the western tracking limit
        TimeToWesternLimit = Misc.DoubleModulus(Mount.CurrentRA - Mount.RAOfWesternLimit, 24)
        If ((TimeToWesternLimit < (clsAction.ImagerExpTime / 3600)) Or _
            (TimeToWesternLimit > 12)) And _
            (Mount.MountSide = WestSide) Then
            'Need to do a meridian flip!!!
            'skip autoguiding setup stuff here
            'flip will occur below and take care of getting the new guidestar
            NeedToFlip = True
        End If
    End If
    
    
    If AutoguiderEnabled And Not (Autoguiding And frmOptions.chkContinuousAutoguide.Value = vbChecked) And _
        (frmOptions.chkDisableGuideStarRecovery.Value <> vbChecked) And LastGuideStarX <> 0 And LastGuideStarY <> 0 And Not NeedToFlip And Not Aborted Then
        
        If clsAction.AutoguiderDitherFreq > 0 And frmOptions.chkContinuousAutoguide.Value <> vbChecked Then
            'try to recover last guide star + dithering
            'compute dithered guide star pos
            GuideStarX = LastGuideStarX + (CDbl(DitherIndexX) * clsAction.AutoguiderDitherStep)
            GuideStarY = LastGuideStarY + (CDbl(DitherIndexY) * clsAction.AutoguiderDitherStep)
            
            If objCameraControl.MultiStarGuidingEnabled Then
                Call AddToStatus("Disabling Multi-Star guiding to use guider to recenter previous guide star and to dither.")
                objCameraControl.MultiStarGuidingEnabled = False
                
                Call objCameraControl.UseGuideStarAt(GuideStarX, GuideStarY)
                Call AddToStatus("Attempting to recover last guide star at " & Format(GuideStarX, "0.0") & "," & Format(GuideStarY, "0.0") & ".")
            
                Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                objCameraControl.AGExposureTime = LastAGExposure
                
                GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                Call AddToStatus("Re-enabling Multi-Star guiding.")
                objCameraControl.MultiStarGuidingEnabled = True
                If GuideStarRecoveryStatus Then
                    GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                End If
            Else
                Call objCameraControl.UseGuideStarAt(GuideStarX, GuideStarY)
                Call AddToStatus("Attempting to recover last guide star at " & Format(GuideStarX, "0.0") & "," & Format(GuideStarY, "0.0") & ".")
            
                Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                objCameraControl.AGExposureTime = LastAGExposure
            
                GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
            End If
                        
            StartedAutoguiderAlready = GuideStarRecoveryStatus
        Else
            GuideStarRecoveryStatus = False
        End If
        
        If Not GuideStarRecoveryStatus And Not Aborted Then
            'either haven't tried, or wasnt able to guide at the dithered position
            'use the last centered position
            If objCameraControl.MultiStarGuidingEnabled Then
                Call AddToStatus("Disabling Multi-Star guiding to use guider to recenter previous guide star.")
                objCameraControl.MultiStarGuidingEnabled = False
            
                Call AddToStatus("Attempting to recover last guide star at " & Format(LastGuideStarX, "0.0") & "," & Format(LastGuideStarY, "0.0") & ".")
                Call objCameraControl.UseGuideStarAt(LastGuideStarX, LastGuideStarY)
                Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                objCameraControl.AGExposureTime = LastAGExposure
                GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
            
                Call AddToStatus("Re-enabling Multi-Star guiding.")
                objCameraControl.MultiStarGuidingEnabled = True
                If GuideStarRecoveryStatus Then
                    GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                End If
            Else
                Call AddToStatus("Attempting to recover last guide star at " & Format(LastGuideStarX, "0.0") & "," & Format(LastGuideStarY, "0.0") & ".")
                Call objCameraControl.UseGuideStarAt(LastGuideStarX, LastGuideStarY)
                Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                objCameraControl.AGExposureTime = LastAGExposure
                GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
            End If
            
            If clsAction.AutoguiderDitherFreq > 0 And frmOptions.chkContinuousAutoguide.Value <> vbChecked Then
                'set this false now since I want the dither operation to complete later
                StartedAutoguiderAlready = False
            Else
                'set this true since I should be good to go now!
                StartedAutoguiderAlready = True
            End If
        End If
        
        If GuideStarRecoveryStatus Then
            Call AddToStatus("Guide star recovery successful!")
            clsAction.AutoguiderXPos = LastGuideStarX
            clsAction.AutoguiderYPos = LastGuideStarY
            clsAction.AutoguiderExpTime = LastAGExposure
        Else
            Call AddToStatus("Guide star recovery unsuccessful.")
        End If
    Else
        GuideStarRecoveryStatus = False
    End If
    
    'Setup Autoguider Automatic Expsosure time
    If (Not Aborted) And (frmOptions.chkEnable = vbChecked) And AutoguiderEnabled And _
        (Not (Autoguiding And frmOptions.chkContinuousAutoguide.Value = vbChecked)) And (Not GuideStarRecoveryStatus) And (Not NeedToFlip) Then
        
        'reset dithering parameters since I will be using a new guide star
        DitherIndexX = 0
        DitherIndexY = 0
        DitherIncLimit = 1
        DitherIncAxis = 0   '0=X, 1=Y
        DitherIncCount = 0
        DitherIncDir = 1
        
        Counter = 0
        Do
            AbortThisAction = Not SetupAutoguiderAutomaticExposureTime(clsAction)
            If AbortThisAction Then Exit Do 'just bail out - could not find a guide star
            
            Counter = Counter + 1
            Call AddToStatus("Trying to autoguide on star at " & Format(clsAction.AutoguiderXPos, "0.0") & "," & Format(clsAction.AutoguiderYPos, "0.0") & ".")
            
            Call objCameraControl.UseGuideStarAt(clsAction.AutoguiderXPos, clsAction.AutoguiderYPos)
                        
            'Set the guide exposure time again here to work around a bug in CCDSoft
            objCameraControl.AGExposureTime = clsAction.AutoguiderExpTime
        Loop Until SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay) Or Aborted Or Counter > 2
            
        If Counter > 2 Then
            Call AddToStatus("Stopping autoguider...")
            Autoguiding = False
            Call objCameraControl.StopAutoguider
            Call Wait(1)
                
            If Not (CBool(GetMySetting("CustomSetting", "ContinueImageWhenAutoguidingFailed", "0"))) Then
                AbortThisAction = True
                Call AddToStatus("Unable to guide on the chosen guide star - stoping this action.")
            Else
                Call AddToStatus("Unable to guide on the chosen guide star - continuing unguided.")
                AutoguiderEnabled = False
            End If
        
            If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                'Send e-mail!
                Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
            End If
        ElseIf AbortThisAction Then
            Call AddToStatus("Stopping autoguider...")
            Autoguiding = False
            Call objCameraControl.StopAutoguider
            Call Wait(1)
                
            Call AddToStatus("Unable to find a guide star.")
                
            If (CBool(GetMySetting("CustomSetting", "ContinueImageWhenAutoguidingFailed", "0"))) And Not Aborted Then
                AbortThisAction = False
                Call AddToStatus("Unable to guide - continuing unguided.")
                AutoguiderEnabled = False
            End If
                
            If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarAcquisitionFailed).Value = vbChecked And Not MainMod.AbortButton Then
                'Send e-mail!
                Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Acquisition Failed", "CCD Commander was not able to find a suitable guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
            End If
        Else
            Call AddToStatus("Automatic guide star acquisition successful!")
            
            LastGuideStarX = clsAction.AutoguiderXPos
            LastGuideStarY = clsAction.AutoguiderYPos
            LastAGExposure = clsAction.AutoguiderExpTime
            
            StartedAutoguiderAlready = True
        End If
    End If
        
    If (LastDitherCounter > clsAction.AutoguiderDitherFreq) Then
        LastDitherCounter = clsAction.AutoguiderDitherFreq
    End If
        
    DitherCounter = LastDitherCounter
    
    frmMainActionText = frmMain.txtCurrentAction.Text
    
    Do While TotalExposures > 0 And Not Aborted And Not AbortThisAction
        frmMain.txtCurrentAction.Text = frmMainActionText & " - " & (clsAction.ImagerNumExp - TotalExposures + 1) & " of " & clsAction.ImagerNumExp
    
CameraMeridianFlip:
        'If connected to telescope and I have a GEM
        If Mount.TelescopeConnected And frmOptions.optMountType(0).Value Then
            'maybe do a flip...first check if I can do another exposre before the western tracking limit
            TimeToWesternLimit = Misc.DoubleModulus(Mount.CurrentRA - Mount.RAOfWesternLimit, 24)
            If ((TimeToWesternLimit < (clsAction.ImagerExpTime / 3600)) Or _
                (TimeToWesternLimit > 12)) And _
                (Mount.MountSide = WestSide) Then
                
                'Need to do a meridian flip!!!
                Call AddToStatus("Need to do a meridian flip!")
                'First Stop the autoguider (if running)
                If Autoguiding Then
                    Call AddToStatus("Stopping autoguider...")
                    Autoguiding = False
                    Call objCameraControl.StopAutoguider
                    
                    'force dithercounter to 0 to get autoguiding to start again
                    DitherCounter = 0
                End If
                
                'Reset dither stuff since I've got to get new guide star - unless the rotate option is enabled
                'otherwise try to recover old star
                If (clsAction.RotateAfterFlip <> vbChecked) Or (frmOptions.chkInternalGuider.Value <> vbChecked) Then
                    LastGuideStarX = 0
                    LastGuideStarY = 0
                End If
                
                If Not Aborted Then
                    AbortThisAction = Not Mount.FlipAtMeridian(clsAction.ImageLinkAfterMeridianFlip, clsAction.clsImageLinkAction, clsAction.RotateAfterFlip, clsAction.DoubleImageLink)
                End If
                
                'Flip may have included a plate solve that changed the sub-frame settings - reset the sub-frame
                If Not Aborted Then
                    If clsAction.FrameSize = FullFrame Then
                        Call AddToStatus("Setting imager to full frame.")
                        Call objCameraControl.SubFrame(False)
                    ElseIf clsAction.FrameSize = HalfFrame Then
                        Call AddToStatus("Setting imager to half frame.")
                        With objCameraControl
                            Call .SubFrame(True, .WidthInPixels / 4, .WidthInPixels * 3 / 4, .HeightInPixels / 4, .HeightInPixels * 3 / 4)
                        End With
                    ElseIf clsAction.FrameSize = QuarterFrame Then
                        With objCameraControl
                            Call AddToStatus("Setting imager to quarter frame.")
                            Call .SubFrame(True, .WidthInPixels * 3 / 8, .WidthInPixels * 5 / 8, .HeightInPixels * 3 / 8, .HeightInPixels * 5 / 8)
                        End With
                    ElseIf clsAction.FrameSize = CustomFrame Then
                        With objCameraControl
                            If (.WidthInPixels < clsAction.CustomFrameWidth) Then
                                Call AddToStatus("Custom Frame Width larger than imager width.  Resetting frame width to maximum.")
                                clsAction.CustomFrameWidth = .WidthInPixels
                            End If
                            
                            If (.HeightInPixels < clsAction.CustomFrameHeight) Then
                                Call AddToStatus("Custom Frame Height larger than imager height.  Resetting frame height to maximum.")
                                clsAction.CustomFrameHeight = .HeightInPixels
                            End If
                            
                            Call AddToStatus("Setting imager to custom frame size of " & clsAction.CustomFrameWidth & " by " & clsAction.CustomFrameHeight & ".")
                            
                            Call .SubFrame(True, (.WidthInPixels - clsAction.CustomFrameWidth) / 2, .WidthInPixels - (.WidthInPixels - clsAction.CustomFrameWidth) / 2, _
                                            (.HeightInPixels - clsAction.CustomFrameHeight) / 2, .HeightInPixels - (.HeightInPixels - clsAction.CustomFrameHeight) / 2)
                        End With
                    End If
                End If
                                               
                'filter might have changed for image link - take dummy image to force change back.
                If clsAction.ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 And Not Aborted Then
                    Call AddToStatus("Setting filter to " & frmOptions.lstFilters.List(clsAction.ImagerFilter) & ".")
                    
                    'take a dummy exposure to get things setup
                    Call Camera.ForceFilterChange(clsAction.ImagerFilter)
                End If
                
                If Not (Rotator.RotatorConnected And (frmOptions.chkGuiderRotates.Value = vbChecked)) And objCameraControl.ReverseXNecessary Then
                    objCameraControl.ReverseX = Not objCameraControl.ReverseX
                End If
                            
                If AutoguiderEnabled And (frmOptions.chkDisableGuideStarRecovery.Value <> vbChecked) And (LastGuideStarX <> 0) And (LastGuideStarY <> 0) And Not Aborted Then
                    'try to recover last guide star
                    Call AddToStatus("Trying to recover guide star after meridian flip.")
                    
                    If clsAction.AutoguiderDitherFreq > 0 And frmOptions.chkContinuousAutoguide.Value <> vbChecked Then
                        'try to recover last guide star + dithering
                        'compute dithered guide star pos
                        GuideStarX = LastGuideStarX + (CDbl(DitherIndexX) * clsAction.AutoguiderDitherStep)
                        GuideStarY = LastGuideStarY + (CDbl(DitherIndexY) * clsAction.AutoguiderDitherStep)
                        
                        If objCameraControl.MultiStarGuidingEnabled Then
                            Call AddToStatus("Disabling Multi-Star guiding to use guider to recenter and dither previous guide star.")
                            objCameraControl.MultiStarGuidingEnabled = False
                            
                            Call objCameraControl.UseGuideStarAt(GuideStarX, GuideStarY)
                            Call AddToStatus("Attempting to recover last guide star at " & Format(GuideStarX, "0.0") & "," & Format(GuideStarY, "0.0") & ".")
                        
                            Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                            objCameraControl.AGExposureTime = LastAGExposure
                            
                            GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                            Call AddToStatus("Re-enabling Multi-Star guiding.")
                            objCameraControl.MultiStarGuidingEnabled = True
                            If GuideStarRecoveryStatus Then
                                GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                            End If
                        Else
                            Call objCameraControl.UseGuideStarAt(GuideStarX, GuideStarY)
                            Call AddToStatus("Attempting to recover last guide star at " & Format(GuideStarX, "0.0") & "," & Format(GuideStarY, "0.0") & ".")
                        
                            Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                            objCameraControl.AGExposureTime = LastAGExposure
                            
                            GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                        End If
                        
                        
                        StartedAutoguiderAlready = GuideStarRecoveryStatus
                    Else
                        GuideStarRecoveryStatus = False
                    End If
                    
                    If Not GuideStarRecoveryStatus Then
                        'either haven't tried, or wasnt able to guide at the dithered position
                        'use the last centered position
                        If objCameraControl.MultiStarGuidingEnabled Then
                            Call AddToStatus("Disabling Multi-Star guiding to use guider to recenter previous guide star.")
                            objCameraControl.MultiStarGuidingEnabled = False
                        
                            Call AddToStatus("Attempting to recover last guide star at " & Format(LastGuideStarX, "0.0") & "," & Format(LastGuideStarY, "0.0") & ".")
                            Call objCameraControl.UseGuideStarAt(LastGuideStarX, LastGuideStarY)
                            Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                            objCameraControl.AGExposureTime = LastAGExposure
                            GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                        
                            Call AddToStatus("Re-enabling Multi-Star guiding.")
                            objCameraControl.MultiStarGuidingEnabled = True
                            If GuideStarRecoveryStatus Then
                                GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                            End If
                        Else
                            Call AddToStatus("Attempting to recover last guide star at " & Format(LastGuideStarX, "0.0") & "," & Format(LastGuideStarY, "0.0") & ".")
                            Call objCameraControl.UseGuideStarAt(LastGuideStarX, LastGuideStarY)
                            Call AddToStatus("Setting autoguider exposure time to " & LastAGExposure & " seconds.")
                            objCameraControl.AGExposureTime = LastAGExposure
                            GuideStarRecoveryStatus = SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay)
                        End If
                        
                        If clsAction.AutoguiderDitherFreq > 0 And frmOptions.chkContinuousAutoguide.Value <> vbChecked Then
                            'set this false now since I want the dither operation to complete later
                            StartedAutoguiderAlready = False
                        Else
                            'set this true since I should be good to go now!
                            StartedAutoguiderAlready = True
                        End If
                    End If
                    
                    If GuideStarRecoveryStatus Then
                        Call AddToStatus("Guide star recovery successful!")
                        clsAction.AutoguiderXPos = LastGuideStarX
                        clsAction.AutoguiderYPos = LastGuideStarY
                        clsAction.AutoguiderExpTime = LastAGExposure
                        
                        StartedAutoguiderAlready = True
                    Else
                        Call AddToStatus("Guide star recovery unsuccessful.")
                    End If
                Else
                    GuideStarRecoveryStatus = False
                End If
                                        
                'Setup Autoguider Automatic Expsosure time
                If frmOptions.chkEnable = vbChecked And AutoguiderEnabled And Not Aborted And Not AbortThisAction And Not GuideStarRecoveryStatus Then
                    
                    DitherIndexX = 0
                    DitherIndexY = 0
                    DitherIncLimit = 1
                    DitherIncAxis = 0   '0=X, 1=Y
                    DitherIncCount = 0
                    DitherIncDir = 1
                    
                    Counter = 0
                    Do
                        AbortThisAction = Not SetupAutoguiderAutomaticExposureTime(clsAction)
                        If AbortThisAction Then Exit Do 'just bail out - could not find a guide star
                        
                        Counter = Counter + 1
                        Call AddToStatus("Trying to autoguide on star at " & Format(clsAction.AutoguiderXPos, "0.0") & "," & Format(clsAction.AutoguiderYPos, "0.0") & ".")
                        Call objCameraControl.UseGuideStarAt(clsAction.AutoguiderXPos, clsAction.AutoguiderYPos)
                    Loop Until SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, clsAction.AutoguiderDelay) Or AbortThisAction Or Aborted Or Counter > 2
                
                    If Counter > 2 Then
                        AbortThisAction = True
                        Call AddToStatus("Unable to guide on the chosen guide star - stoping this action.")
                    
                        If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                            'Send e-mail!
                            Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                        End If
                    ElseIf AbortThisAction Then
                        Call AddToStatus("Unable to find a guide star.")
                            
                        If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarAcquisitionFailed).Value = vbChecked Then
                            'Send e-mail!
                            Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Acquisition Failed", "CCD Commander was not able to find a suitable guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                        End If
                    Else
                        Call AddToStatus("Automatic guide star acquisition successful!")
                        
                        LastGuideStarX = clsAction.AutoguiderXPos
                        LastGuideStarY = clsAction.AutoguiderYPos
                        LastAGExposure = clsAction.AutoguiderExpTime
                    
                        StartedAutoguiderAlready = True
                    End If
                End If
            End If
        End If
        
        If DitherCounter = 0 And Not Aborted And Not AbortThisAction Then
            If AutoguiderEnabled And clsAction.AutoguiderDitherFreq > 0 And Autoguiding And _
                frmOptions.chkContinuousAutoguide.Value = vbUnchecked And Not StartedAutoguiderAlready Then
                
                Call AddToStatus("Stopping autoguider...")
                Autoguiding = False
                Call objCameraControl.StopAutoguider
                Call Wait(1)
            End If
            
            StartedAutoguiderAlready = False
            
            DitherCounter = clsAction.AutoguiderDitherFreq

            If Not Autoguiding And AutoguiderEnabled Then
                objCameraControl.AGExposureTime = clsAction.AutoguiderExpTime
                Call AddToStatus("Setting autoguider exposure time to " & clsAction.AutoguiderExpTime & " seconds.")
                
                If frmOptions.chkEnable.Value = vbUnchecked And frmOptions.lstCameraControl.ListIndex = CameraControl.MaxIm And _
                    TotalExposures = clsAction.ImagerNumExp Then
                    'Need to take an exposure before starting autoguider
                    objCameraControl.AGTakeImage
                    
                    'wait for image
                    Do Until objCameraControl.AGTakeImageComplete Or Aborted
                        Call Wait(1)
                    Loop
                End If
            
                If clsAction.AutoguiderDitherFreq = 0 Or frmOptions.chkContinuousAutoguide.Value = vbChecked Then
                    Call objCameraControl.UseGuideStarAt(clsAction.AutoguiderXPos, clsAction.AutoguiderYPos)
                    Call AddToStatus("Setting guide star position to " & Format(clsAction.AutoguiderXPos, "0.0") & "," & Format(clsAction.AutoguiderYPos, "0.0") & ".")
                
                    'don't care about multi-star guiding here
                
                    If Not Aborted And Not AbortThisAction Then
                        ' No Guider Delay here - haven't moved the mount prior to being here, only guider movements
                        AbortThisAction = Not SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, 0)
                        
                        If Not AbortThisAction And frmOptions.chkEnable.Value = vbUnchecked Then
                            'Save these now since automatic guide star is disabled
                            LastGuideStarX = clsAction.AutoguiderXPos
                            LastGuideStarY = clsAction.AutoguiderYPos
                            LastAGExposure = clsAction.AutoguiderExpTime
                        ElseIf AbortThisAction Then
                            If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                                'Send e-mail!
                                Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                            End If
                        End If
                    End If
                Else
                    'compute new guide star pos
                    GuideStarX = clsAction.AutoguiderXPos + (CDbl(DitherIndexX) * clsAction.AutoguiderDitherStep)
                    GuideStarY = clsAction.AutoguiderYPos + (CDbl(DitherIndexY) * clsAction.AutoguiderDitherStep)
                    
                    If objCameraControl.MultiStarGuidingEnabled Then
                        Call AddToStatus("Disabling Multi-Star guiding to use guider to dither guide star.")
                        objCameraControl.MultiStarGuidingEnabled = False
                        
                        Call objCameraControl.UseGuideStarAt(GuideStarX, GuideStarY)
                        Call AddToStatus("Setting guide star position to " & Format(GuideStarX, "0.0") & "," & Format(GuideStarY, "0.0") & ".")
                        
                        If Not Aborted And Not AbortThisAction Then
                            ' No Guider Delay here - haven't moved the mount prior to being here, only guider movements
                            AbortThisAction = Not SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, 0)
                            
                            Call AddToStatus("Re-enabling Multi-Star guiding.")
                            objCameraControl.MultiStarGuidingEnabled = True
                            If Not AbortThisAction Then
                                AbortThisAction = Not SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, 0)
                            End If
                            
                            If Not AbortThisAction And frmOptions.chkEnable.Value = vbUnchecked Then
                                'Save these now since automatic guide star is disabled
                                LastGuideStarX = clsAction.AutoguiderXPos
                                LastGuideStarY = clsAction.AutoguiderYPos
                                LastAGExposure = clsAction.AutoguiderExpTime
                            ElseIf AbortThisAction Then
                                If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                                    'Send e-mail!
                                    Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                                End If
                            End If
                        End If
                    Else
                        Call objCameraControl.UseGuideStarAt(GuideStarX, GuideStarY)
                        Call AddToStatus("Setting guide star position to " & Format(GuideStarX, "0.0") & "," & Format(GuideStarY, "0.0") & ".")
                        
                        If Not Aborted And Not AbortThisAction Then
                            ' No Guider Delay here - haven't moved the mount prior to being here, only guider movements
                            AbortThisAction = Not SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, 0)
                            
                            If Not AbortThisAction And frmOptions.chkEnable.Value = vbUnchecked Then
                                'Save these now since automatic guide star is disabled
                                LastGuideStarX = clsAction.AutoguiderXPos
                                LastGuideStarY = clsAction.AutoguiderYPos
                                LastAGExposure = clsAction.AutoguiderExpTime
                            ElseIf AbortThisAction Then
                                If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                                    'Send e-mail!
                                    Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                                End If
                            End If
                        End If
                    End If
                End If
            ElseIf AutoguiderEnabled And Not AbortThisAction Then
                ' Need to monitor guide error here.
                ' Guider could move off center during the download
                ' No Guider Delay here - haven't moved the mount prior to being here, only guider movements
                AbortThisAction = Not SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, 0, True)
                
                If AbortThisAction Then
                    If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                        'Send e-mail!
                        Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                    End If
                End If
                
            ElseIf clsAction.UnguidedDither = vbChecked And clsAction.AutoguiderDitherFreq > 0 Then
                UnguidedDitherRA = 0
                UnguidedDitherDec = 0
                
                'compute new telescope pos
                If clsAction.XAxisDither And (DitherIncAxis = 0 Or (Not clsAction.YAxisDither)) Then  'X
                    UnguidedDitherRA = (CDbl(DitherIncDir) * clsAction.AutoguiderDitherStep / 3600) / Cos(StartDec * PI / 180) / 15
                ElseIf (clsAction.YAxisDither) Then 'Y
                    UnguidedDitherDec = CDbl(DitherIncDir) * clsAction.AutoguiderDitherStep / 3600
                End If
                                
                If (UnguidedDitherRA <> 0) Or (UnguidedDitherDec <> 0) Then
                    Call Mount.JogTelescope(UnguidedDitherRA, UnguidedDitherDec)
                    
                    If Not Mount.CurrentTargetJ2000 Then
                        'Reset these so that the maintain dither will work correctly
                        StartRA = Mount.CurrentRA
                        StartDec = Mount.CurrentDec
                    End If
                End If
            End If
        ElseIf StartedAutoguiderAlready Then
            ' Don't need to do anything else, I just started autoguiding, already checked the error - so go direct to the image
            StartedAutoguiderAlready = False
        Else
            ' Need to monitor guide error here.
            ' Guider could move off center during the download
            ' No Guider Delay here - haven't moved the mount prior to being here, only guider movements
            If AutoguiderEnabled And Not AbortThisAction Then
                AbortThisAction = Not SetupAutoguider(clsAction.AutoguiderMinError, clsAction.AutoguiderMaxGuideCycles, 0, True)
                
                If AbortThisAction Then
                    If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFailedToCenter).Value = vbChecked Then
                        'Send e-mail!
                        Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Failed To Center", "CCD Commander was not able to guide on the chosen guide star." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                    End If
                End If
            End If
        End If
        
        If Not Aborted And Not AbortThisAction And clsAction.ImagerDelayTime > 0 Then
            Call AddToStatus("Starting imager delay...")
            Call Wait(clsAction.ImagerDelayTime)
        End If
        
        AutoguiderGuideFailedEvent = False
        
        'Set file name here
        If clsAction.AutosaveExposure = vbChecked Then
            Call AddToStatus("Setting file name prefix to: " & Misc.FixFileName(clsAction.FileNamePrefix))
            If clsAction.UseGlobalSaveToLocation Then
                SaveToPath = frmOptions.txtSaveTo
            Else
                SaveToPath = Misc.FixFileName(clsAction.FileSavePath, True)
            End If
            
            'Make sure SaveToPath exists
            On Error Resume Next
            Call ChDir(SaveToPath)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Call Misc.CreatePath(SaveToPath)
            End If
            On Error GoTo 0
                        
            Call objCameraControl.AutoSave(True, SaveToPath, Misc.FixFileName(clsAction.FileNamePrefix))
        Else
            Call objCameraControl.AutoSave(False)
        End If
                
        'About to start an exposure - the search for a guide star might have taken enough time, that I need to do a meridian flip now
        '(but didn't before).  Check the time again now and just jump back above if necessary.
        If Mount.TelescopeConnected And frmOptions.optMountType(0).Value And Not Aborted Then
            'maybe do a flip...first check if I can do another exposre before the western tracking limit
            TimeToWesternLimit = Misc.DoubleModulus(Mount.CurrentRA - Mount.RAOfWesternLimit, 24)
            If ((TimeToWesternLimit < (clsAction.ImagerExpTime / 3600)) Or _
                (TimeToWesternLimit > 12)) And _
                (Mount.MountSide = WestSide) Then
                
                GoTo CameraMeridianFlip
            End If
        End If
        
        If Not Aborted And Not AbortThisAction Then
            Call AddToStatus("Starting imager exposure (" & (clsAction.ImagerNumExp - TotalExposures + 1) & " of " & clsAction.ImagerNumExp & ").")
            
            Call objCameraControl.TakeImage
        End If
        
        StarFadedCount = 0
        GuiderErrorExceededLimitCount = 0
        AutoguiderGuideErrorEvent = False
        LastGuideErrorEventTime = Now
        Do Until objCameraControl.TakeImageComplete Or (Aborted And Not SoftSkip) Or AbortThisAction
            If frmOptions.chkRestartGuidingWhenLargeError = vbChecked Then
                If AutoguiderGuideErrorEvent Or _
                    (DateDiff("s", LastGuideErrorEventTime, Now) > (clsAction.AutoguiderExpTime * 1.2) And (clsAction.AutoguiderExpTime >= 0.5)) Or _
                    (DateDiff("s", LastGuideErrorEventTime, Now) > (clsAction.AutoguiderExpTime * 2) And (clsAction.AutoguiderExpTime < 0.5)) Then
                    
                    AutoguiderGuideErrorEvent = False
                    LastGuideErrorEventTime = Now
                    
                    If (Abs(objCameraControl.GuideErrorX) > Settings.GuiderRestartError) Or (Abs(objCameraControl.GuideErrorY) > Settings.GuiderRestartError) Then
                        GuiderErrorExceededLimitCount = GuiderErrorExceededLimitCount + 1
                        
                        Call AddToStatus("Guider error beyond limit.  Count = " & GuiderErrorExceededLimitCount)
                        
                        If (GuiderErrorExceededLimitCount >= Settings.GuiderRestartCycles) Then
                            ' too many consecutive guide errors above the limit
                            ' Abort this exposure
                            
                            Call AddToStatus("Too many consecutive guide errors beyond limit. Aborting image.")
                            Call objCameraControl.Abort
                            Call objCameraControl.StopAutoguider
                            Autoguiding = False
                            
                            Call AddToStatus("Restarting image #" & (clsAction.ImagerNumExp - TotalExposures + 1) & ".")
                            
                            TotalExposures = TotalExposures + 1
                            
                            Exit Do
                        End If
                    Else
                        GuiderErrorExceededLimitCount = 0
                    End If
                End If
            End If
            
            If AutoguiderGuideFailedEvent Then
                AutoguiderGuideFailedEvent = False
                
                Call AddToStatus("Guide star faded...")
                StarFadedCount = StarFadedCount + 1
                
                If StarFadedCount > Settings.MaximumStarFadedErrors Then
                    Call AddToStatus("Too many guide star faded events, giving up.")
                    objCameraControl.Abort
                    AbortThisAction = True
                    
                    If frmOptions.chkEMailAlert(EMailAlertIndexes.GuideStarFaded).Value = vbChecked Then
                        'Send e-mail!
                        Call EMail.SendEMail(frmMain, "CCD Commander Guide Star Faded", "The guide star has faded more than 5 times - take image action is giving up." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                    End If
                    
                    Exit Do
                End If
            End If
                
            Call Wait(0.1)
        Loop
        
        If Not Aborted And Not AbortThisAction Then
            Call AddToStatus("Imager exposure complete.")
                    
            If clsAction.AutoguiderDitherFreq > 0 Then
                DitherCounter = DitherCounter - 1
                LastDitherCounter = DitherCounter
            End If
            
            TotalExposures = TotalExposures - 1
        End If
        
        If DitherCounter = 0 And clsAction.AutoguiderDitherFreq > 0 Then
            'Compute next dither index
            
            If Abs(DitherIndexX + DitherIncDir) > clsAction.AutoguiderDitherAmount And DitherIncAxis = 0 Then
                DitherIncDir = -1
            ElseIf DitherIndexX = 0 And DitherIndexY = 0 Then
                DitherIncLimit = 1
                DitherIncAxis = 0   '0=X, 1=Y
                DitherIncCount = 0
                DitherIncDir = 1
            End If
            
            'increment dither location
            If clsAction.XAxisDither And (DitherIncAxis = 0 Or (Not clsAction.YAxisDither)) Then  'X
                DitherIndexX = DitherIndexX + DitherIncDir
            ElseIf (clsAction.YAxisDither) Then 'Y
                DitherIndexY = DitherIndexY + DitherIncDir
            End If
            
            If DitherIncLimit <> 0 Then
                DitherIncCount = DitherIncCount + DitherIncDir
            End If
            
            If clsAction.XAxisDither And clsAction.YAxisDither Then
                If DitherIncCount = DitherIncLimit And DitherIncLimit <> 0 Then
                    If DitherIncAxis = 0 Then
                        DitherIncAxis = 1
                    Else
                        DitherIncAxis = 0
                        DitherIncDir = DitherIncDir * -1
                        DitherIncLimit = (Abs(DitherIncLimit) + 1) * DitherIncDir
                    End If
                    DitherIncCount = 0
                ElseIf DitherIncCount = 0 Then
                    If DitherIncAxis = 1 Then
                        DitherIncAxis = 0
                    Else
                        DitherIncAxis = 1
                        DitherIncDir = DitherIncDir * -1
                        DitherIncLimit = (Abs(DitherIncLimit) - 1) * DitherIncDir * -1
                    End If
                    DitherIncCount = DitherIncLimit
                End If
            End If
        End If
    Loop
    
    If Autoguiding And frmOptions.chkContinuousAutoguide.Value = vbUnchecked Then
        Call AddToStatus("Stopping autoguider...")
        Autoguiding = False
        objCameraControl.StopAutoguider
        Call Wait(1)
        
        If clsAction.CenterAO = 1 And Not Aborted Then
            Call AddToStatus("Centering AO...")
            Call objCameraControl.CenterAO
        End If
    End If
    
    If clsAction.SyncToCurrentAtEnd = vbChecked And Not Aborted Then
        Call Mount.Sync(Mount.CurrentRA, Mount.CurrentDec, False)
    End If
    
    'Disable autosave
    Call objCameraControl.AutoSave(False)
    
    If Not Aborted And Not AbortThisAction Then
        Call AddToStatus("Take Images Action complete.")
    ElseIf AbortThisAction Then
        Call AddToStatus("Take Images Action failed.")
    End If
End Sub

Private Function SetupAutoguider(MaxAllowedError As Double, MaxNumCycles As Integer, GuiderDelay As Integer, Optional DoNotStartAutoguider As Boolean = False) As Boolean
    Dim MaxError As Single
    Dim Counter As Integer
    Dim Counter2 As Integer
    Dim Time1 As Date
    Dim Time2 As Date
    Dim DelayTime As Double
    Dim RetriedGuiding As Boolean
    Dim DoNotUseEvents As Boolean
    Dim StarFadedCount As Integer
    
    RetriedGuiding = False
    
    If frmOptions.lstCameraControl.ListIndex = CameraControl.CCDSoftAO Or frmOptions.lstCameraControl.ListIndex = CameraControl.TheSkyX Then
        DoNotUseEvents = True
    Else
        DoNotUseEvents = False
    End If
    
    If Not DoNotStartAutoguider Then
        Call AddToStatus("Starting to autoguide.")
        'This starts autoguiding and waits until the autoguider error is less then 0.3pix
        objCameraControl.StartAutoguider
        
        If objCameraControl.AutoguiderInterval < 1 Then
            DelayTime = 4
        Else
            DelayTime = objCameraControl.AutoguiderInterval * 4
        End If
        
        Call Wait(DelayTime + GuiderDelay)
    Else
        If objCameraControl.AutoguiderInterval < 1 Then
            DelayTime = 2
        Else
            DelayTime = objCameraControl.AutoguiderInterval * 2
        End If
    End If
        
    MaxError = MaxAllowedError  'Set this here to ensure I get a valid reading before exiting my loop below
    Counter = 0
    StarFadedCount = 0
    AutoguiderGuideErrorEvent = False
    AutoguiderGuideFailedEvent = False
    Do
        Call AddToStatus("Checking autoguider max error...")
        
        If DoNotUseEvents And Not DoNotStartAutoguider Then
            Call Wait(objCameraControl.AutoguiderInterval * 1.1)
        End If
        
        If UseAGBeforeTakeImageEvent Then
            'clear AutoguiderGuideErrorEvent to ensure I get a new one before checking the guide error
            AutoguiderGuideErrorEvent = False
        End If
        
        Counter2 = 0
        Time1 = Now
        Do Until AutoguiderGuideErrorEvent Or Aborted Or DoNotUseEvents Or AutoguiderGuideFailedEvent Or Counter2 > 5
            Call Wait(0.05)
            Time2 = Now
            If UseAGBeforeTakeImageEvent = False Then
                If DateDiff("s", Time1, Time2) > (DelayTime * 4) Then
                    'double check that guider is running
                    If objCameraControl.GuiderRunning Then
                        'been waiting for more then 16 guider exposures.  Maybe try to use the BeforeTakeImage event.
                        If frmOptions.lstCameraControl.ListIndex = CameraControl.CCDSoft Then
                            Call AddToStatus("Never received GuideError event - switching to another event.")
                            UseAGBeforeTakeImageEvent = True
                        ElseIf frmOptions.lstCameraControl.ListIndex = CameraControl.MaxIm Then
                            Call AddToStatus("Never received GuideError event - stopping event checking.")
                            DoNotUseEvents = True
                        End If
                    Else
                        Call AddToStatus("Guider is not running!?!?!?!  Restarting guider.")
                        Counter2 = Counter2 + 1
                        objCameraControl.StartAutoguider
                        
                        If objCameraControl.AutoguiderInterval < 1 Then
                            DelayTime = 4
                        Else
                            DelayTime = objCameraControl.AutoguiderInterval * 4
                        End If
                        
                        Call Wait(DelayTime)
                        
                        Time1 = Now
                    End If
                End If
            ElseIf DateDiff("s", Time1, Time2) > (DelayTime * 8) Then
                'been waiting waaay to long...try to restart guiding
                If RetriedGuiding Then
                    Call AddToStatus("Never received guide error event - aborting.")
                    Aborted = True
                Else
                    Call AddToStatus("Never received guide error event - restarting guiding.")
                    UseAGBeforeTakeImageEvent = False
                    RetriedGuiding = True
                    Call objCameraControl.StopAutoguider
                    Call Wait(1)
                    Call objCameraControl.StartAutoguider
                    Call Wait(DelayTime)
                    Call AddToStatus("Checking autoguider max error...")
                    Time1 = Now
                End If
            End If
        Loop
        
        If Counter2 > 5 Then
            'major poroblem
            Call AddToStatus("Serious failure of the autoguider - giving up on this attempt.")
            SetupAutoguider = False
            'Aborted = True
            Call objCameraControl.StopAutoguider
            Exit Function
        ElseIf AutoguiderGuideFailedEvent = True Then
            'Something went wrong, give up
            Call AddToStatus("Star faded...")
            StarFadedCount = StarFadedCount + 1
            
            If StarFadedCount > Settings.MaximumStarFadedErrors Then
                Call AddToStatus("Too many star faded messages, giving up.")
                SetupAutoguider = False
                Call objCameraControl.StopAutoguider
                Exit Function
            End If
            
        ElseIf Not Aborted Then
            'autoguider finished exposure, check error
            MaxError = Abs(objCameraControl.GuideErrorX)
            If MaxError < Abs(objCameraControl.GuideErrorY) Then
                MaxError = Abs(objCameraControl.GuideErrorY)
            End If
        
            Call AddToStatus("Autoguider error: X=" & Format(objCameraControl.GuideErrorX, "0.00") & " Y=" & Format(objCameraControl.GuideErrorY, "0.00"))
        End If
        
        AutoguiderGuideErrorEvent = False
        AutoguiderGuideFailedEvent = False
        
        Counter = Counter + 1
    Loop Until MaxError < MaxAllowedError Or Aborted Or Counter = MaxNumCycles
    
    If Counter = MaxNumCycles And MaxError >= MaxAllowedError Then
        Call AddToStatus("Autoguider error failed to decrease to < " & MaxAllowedError & "!")
        SetupAutoguider = False
        
        Call objCameraControl.StopAutoguider
    ElseIf Aborted Then
        SetupAutoguider = False
    Else
        Autoguiding = True
        Call AddToStatus("Autoguider max error < " & MaxAllowedError & "!")
        SetupAutoguider = True
    End If
End Function

Public Function SetupAutoguiderAutomaticExposureTime(clsAction As ImagerAction) As Boolean
    Dim MinTime As Double
    Dim MaxTime As Double
    Dim MinADU As Long
    Dim MaxADU As Long
    Dim MaxImageVal As Long
    Dim NewTime As Double
    Dim Done As Boolean
'    Dim MaybeDone As Boolean
    Dim MaxImageX As Long
    Dim MaxImageY As Long
    Dim GuideBoxX As Long
    Dim GuideBoxY As Long
    Dim MaxStarMovement As Long
    Dim GuideStar As Variant
    Dim GuideStarSearch As Variant
    Dim GuideStarCollection As New Collection
    
    ReDim GuideStar(2)
    
    MinTime = Settings.MinimumGuideStarExposure
    MaxTime = Settings.MaximumGuideStarExposure
    MinADU = Settings.MinimumGuideStarADU
    MaxADU = Settings.MaximumGuideStarADU
    GuideBoxX = Settings.GuideBoxX / 2
    GuideBoxY = Settings.GuideBoxY / 2
    MaxStarMovement = Settings.MaxStarMovement
    
    If clsAction.AutoguiderDitherFreq > 0 Then
        'increase guide box by the dither step size
        'this will prevent an undesired guide star from being picked up when the dither is active
        GuideBoxX = GuideBoxX + CInt(clsAction.AutoguiderDitherStep)
        GuideBoxY = GuideBoxY + CInt(clsAction.AutoguiderDitherStep)
    End If
    
    NewTime = MinTime
    
    Done = False
    Do
        objCameraControl.AGExposureTime = NewTime
        
        objCameraControl.AGImageType = cdLight
    
        'Removing this so I don't muck up autoguiders that don't have a shutter
        'objCameraControl.AGReduction = cdAutoDark
        
        If Not Aborted Then
            Call AddToStatus("Taking a " & objCameraControl.AGExposureTime & " second autoguider exposure.")
            Call objCameraControl.AGTakeImage
        End If
        
        'wait for image
        Do Until objCameraControl.AGTakeImageComplete Or Aborted
            Call Wait(1)
        Loop
        
        Call Wait(1)
        
        Call AddToStatus("Beginning guide star search...")
        MaxImageVal = objCameraControl.FindGuideStar(MinADU, MaxADU, GuideBoxX, GuideBoxY, clsAction.AutoguiderDitherAmount, MaxImageX, MaxImageY)
        If Aborted Then Exit Do
                
        If MaxImageVal > MinADU Then
            If GuideStarCollection.Count = 0 Then
                'first possible guide star
                'add it to the collection
                Call AddToStatus("Found possible guide star at " & MaxImageX & ", " & MaxImageY & " with a max brightness of " & MaxImageVal & " ADU.")
                
                GuideStar(0) = MaxImageX
                GuideStar(1) = MaxImageY
                GuideStar(2) = MaxImageVal
                GuideStarCollection.Add GuideStar
            Else
                For Each GuideStarSearch In GuideStarCollection
                    If (MaxImageX >= GuideStarSearch(0) - MaxStarMovement And MaxImageX <= GuideStarSearch(0) + MaxStarMovement) And _
                        (MaxImageY >= GuideStarSearch(1) - MaxStarMovement And MaxImageY <= GuideStarSearch(1) + MaxStarMovement) Then
                        'Found a match!
                        'For sure this is a good star!
                        Call AddToStatus("Found appropriate guide star at " & MaxImageX & ", " & MaxImageY & " with a max brightness of " & MaxImageVal & " ADU.")
                        Call AddToStatus("Matches previous possible guide star at " & GuideStarSearch(0) & ", " & GuideStarSearch(1) & " with a max brightness of " & GuideStarSearch(2) & " ADU.")
                        
                        Done = True
                        Exit For
                    End If
                Next GuideStarSearch
                                
                If Not Done Then
                    'found another potential star
                    'add it to the collection
                    Call AddToStatus("Found possible guide star at " & MaxImageX & ", " & MaxImageY & " with a max brightness of " & MaxImageVal & " ADU.")
                    
                    GuideStar(0) = MaxImageX
                    GuideStar(1) = MaxImageY
                    GuideStar(2) = MaxImageVal
                    GuideStarCollection.Add GuideStar
                End If
                
                If Done = True Then
                    clsAction.AutoguiderExpTime = NewTime
                
                    clsAction.AutoguiderXPos = MaxImageX
                    clsAction.AutoguiderYPos = MaxImageY
                End If
            End If
        Else
            If NewTime = MaxTime Then
                Call AddToStatus("Cannot find star greater then " & MinADU & " ADU.")
                SetupAutoguiderAutomaticExposureTime = False
                Exit Function
            Else
                If MaxImageVal > 0 And MaxImageVal < 65536 Then
                    ' x1.1 is to shoot for slightly brighter then the minimum
                    ' Don't increase the exposure beyond the increment value
                    If (NewTime * MinADU * 1.1 / MaxImageVal) > (NewTime + Settings.GuideStarExposureIncrement) Then
                        NewTime = NewTime + Settings.GuideStarExposureIncrement
                    Else
                        NewTime = (NewTime * MinADU * 1.1 / MaxImageVal)
                    End If
                Else
                    NewTime = NewTime + Settings.GuideStarExposureIncrement
                End If
                
                If NewTime > MaxTime Then NewTime = MaxTime
                
                NewTime = CLng(NewTime * 100) / 100
            End If
        End If
    Loop Until Done
    
    If Not Aborted Then
        SetupAutoguiderAutomaticExposureTime = True
    Else
        SetupAutoguiderAutomaticExposureTime = False
    End If
End Function

Public Function TakeImageAndLink(clsAction As ImageLinkSyncAction, Optional MoveToAction As Boolean = False, Optional Precess As Boolean = True) As Boolean
    Dim RA As Double
    Dim Dec As Double
    Dim OriginalExpTime As Double
    Dim OriginalBin As Integer
    Dim NorthAngle As Double
    Dim RequiredNorthAngle As Double
    Dim SpecifiedNorthAngle As Double
    Dim PlateSolveObject As Object
    Dim ActualPixelScale As Double
    Dim LastReduction As ReductionType
    Dim RetryCount As Integer
    Dim Alt As Double
    Dim Az As Double
    Dim PlateSolveResult As Boolean
    Dim SolutionRADeg As Double
    Dim CurrentRADeg As Double
    Dim ErrorVectorMag As Double
    Dim ErrorVectorDir As Double
    Dim SaveToPath As String
    
    If Not Aborted Then
        OriginalBin = objCameraControl.BinX
        objCameraControl.BinXY = clsAction.Bin + 1
        
        Call AddToStatus("Setting imager bin mode to " & clsAction.Bin + 1 & "x" & clsAction.Bin + 1 & ".")
        
        ActualPixelScale = Settings.PixelScale * (clsAction.Bin + 1)
        
        If clsAction.Filter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
            Call AddToStatus("Setting filter to " & frmOptions.lstFilters.List(clsAction.Filter) & ".")
            
            'take a dummy exposure to get things setup
            Call Camera.ForceFilterChange(clsAction.Filter)
        End If
    End If
    
    If Not Aborted Then
        If clsAction.FrameSize = FullFrame Then
            Call objCameraControl.SubFrame(False)
        ElseIf clsAction.FrameSize = HalfFrame Then
            With objCameraControl
'                Call AddToStatus("Full Width = " & .WidthInPixels & ", Full Height = " & .HeightInPixels)
'                Call AddToStatus("Left = " & .WidthInPixels / 4 & ", Right = " & .WidthInPixels * 3 / 4 & ", Top = " & .HeightInPixels / 4 & ", Bottom = " & .HeightInPixels * 3 / 4)
                Call .SubFrame(True, .WidthInPixels / 4, .WidthInPixels * 3 / 4, .HeightInPixels / 4, .HeightInPixels * 3 / 4)
            End With
        ElseIf clsAction.FrameSize = QuarterFrame Then
            With objCameraControl
'                Call AddToStatus("Full Width = " & .WidthInPixels & ", Full Height = " & .HeightInPixels)
'                Call AddToStatus("Left = " & .WidthInPixels * 3 / 8 & ", Right = " & .WidthInPixels * 5 / 8 & ", Top = " & .HeightInPixels * 3 / 8 & ", Bottom = " & .HeightInPixels * 5 / 8)
                Call .SubFrame(True, .WidthInPixels * 3 / 8, .WidthInPixels * 5 / 8, .HeightInPixels * 3 / 8, .HeightInPixels * 5 / 8)
            End With
        End If
    End If
    
    OriginalExpTime = objCameraControl.ExposureTime
    If Not Aborted Then
        objCameraControl.ExposureTime = clsAction.ExpTime
    End If
    
    objCameraControl.ImageType = cdLight
    
    LastReduction = objCameraControl.ImageReduction
    
    If frmOptions.chkDarkSubtractPlateSolveImage.Value = vbChecked Then
        objCameraControl.ImageReduction = ReductionType.AutoDark
    Else
        objCameraControl.ImageReduction = ReductionType.NoReduction
    End If
    
    If clsAction.AutosaveExposure = vbChecked Then
        Call AddToStatus("Setting file name prefix to: " & Misc.FixFileName(clsAction.FileNamePrefix))
        If clsAction.UseGlobalSaveToLocation Then
            SaveToPath = frmOptions.txtSaveTo
        Else
            SaveToPath = Misc.FixFileName(clsAction.FileSavePath, True)
        End If
        
        'Make sure SaveToPath exists
        On Error Resume Next
        Call ChDir(SaveToPath)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Call Misc.CreatePath(SaveToPath)
        End If
        On Error GoTo 0
                    
        Call objCameraControl.AutoSave(True, SaveToPath, Misc.FixFileName(clsAction.FileNamePrefix))
    Else
        Call objCameraControl.AutoSave(False)
    End If
        
    If clsAction.RetryPlateSolveOnFailure Then
        RetryCount = 2
    Else
        RetryCount = 1
    End If
    
    Do
        If Not Aborted And clsAction.DelayTime > 0 Then
            Call AddToStatus("Starting imager delay...")
            Call Wait(clsAction.DelayTime)
        End If
        
        If Not Aborted Then
            Call AddToStatus("Taking " & clsAction.ExpTime & " second image for Plate Solve...")
            Call objCameraControl.TakeImage
        End If
        
        Do Until objCameraControl.TakeImageComplete Or Aborted
            Call Wait(1)
        Loop
        
        'save image
        If Not Aborted And Not MainMod.PlateSolveNoSaveImage Then
            Call objCameraControl.SaveImagerImage(App.Path & "\SavedImage.FIT")
        End If
        
        Camera.PixelSize = objCameraControl.PixelSize
        
        'do image link
        If Not Aborted Then
            If frmOptions.lstPlateSolve.ListIndex = 0 Then
                Set PlateSolveObject = New clsCCDSoftPlateSolve
            ElseIf frmOptions.lstPlateSolve.ListIndex = 1 Then
                Set PlateSolveObject = New clsMaximPlateSolve
            ElseIf frmOptions.lstPlateSolve.ListIndex = 2 Then
                Set PlateSolveObject = New clsPinPointPlateSolve
            ElseIf frmOptions.lstPlateSolve.ListIndex = 3 Then
                Set PlateSolveObject = New clsTheSkyXPlateSolve
            End If
        End If
        
        If Not Aborted Then
            Call AddToStatus("Executing plate solve function.")
            On Error Resume Next
            PlateSolveResult = PlateSolveObject.PlateSolve(App.Path & "\SavedImage.FIT", ActualPixelScale, NorthAngle, RA, Dec)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Call AddToStatus("Error in plate solve function.  Check CCD Commander Plate Solve Options & Settings.")
                TakeImageAndLink = False
            Else
                On Error GoTo 0
                If PlateSolveResult = False Then
                    Call AddToStatus("Plate Solve Failed.")
                    TakeImageAndLink = False
                Else
                    TakeImageAndLink = True
                End If
            End If
            
            If TakeImageAndLink = False And frmOptions.chkEMailAlert(EMailAlertIndexes.PlateSolveFailed).Value = vbChecked Then
                'Send e-mail!
                Call EMail.SendEMail(frmMain, "CCD Commander Plate Solve Failed", "The plate solve action has failed." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
            End If
        End If
        
        Set PlateSolveObject = Nothing
        
        If TakeImageAndLink And Not MoveToAction Then
            NorthAngle = DoubleModulus(NorthAngle, 360)
            SpecifiedNorthAngle = Settings.NorthAngle
            SpecifiedNorthAngle = DoubleModulus(SpecifiedNorthAngle, 360)
            
            'Compute North Angle if necessary
            If frmOptions.lstRotator.ListIndex <> RotatorControl.None Then
                'Call compute function in Rotator module
                RequiredNorthAngle = Rotator.GetCurrentAngleCorrectedForPlateSolve
            Else
                RequiredNorthAngle = Settings.NorthAngle
            End If
            
            'Check north angle
            If frmOptions.chkIgnoreNorthAngle.Value = vbUnchecked Then
                If Abs(NorthAngle - CDbl(RequiredNorthAngle)) > 2 And _
                    Abs(Abs(NorthAngle - CDbl(RequiredNorthAngle)) - 360) > 2 And _
                    Abs(Abs(NorthAngle - CDbl(RequiredNorthAngle)) - 180) > 2 Then
                    Call AddToStatus("North Angle is incorrect!  Specified angle is " & RequiredNorthAngle & " degrees.")
                    Call AddToStatus("Plate Solve Failed.")
                    TakeImageAndLink = False
                    
                    If frmOptions.chkEMailAlert(EMailAlertIndexes.PlateSolveFailed).Value = vbChecked Then
                        'Send e-mail!
                        Call EMail.SendEMail(frmMain, "CCD Commander Plate Solve Failed", "The plate solve action north angle verification failed." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
                    End If
                    
                    'Exit Function
                End If
            Else
                Call AddToStatus("North Angle verification is disabled.")
                TakeImageAndLink = True
            End If
        End If
        
        RetryCount = RetryCount - 1
        
        If Not TakeImageAndLink And RetryCount > 0 And clsAction.SlewMountForRetry And Not Aborted Then
            Call Mount.GetTelescopeAltAz(Alt, Az)
            
            If Not Aborted Then
                Alt = Alt + (CDbl(clsAction.ArcminutesToSlew) / 60)
                If Alt > 90 Then
                    Alt = 90 - (Alt - 90)
                    Az = Misc.DoubleModulus(Az - 180, 360)
                End If
                Call Mount.MoveToAltAz(Alt, Az)
            End If
        End If
    Loop Until RetryCount = 0 Or TakeImageAndLink = True Or Aborted
    
    If Not Aborted And TakeImageAndLink And Precess Then
        'this just Precesses the coordinates.  I really should be doing nutation and maybe obliquity correction
        Call Mount.PrecessCoordinates(RA, Dec)
    End If
    
    If Not Aborted And TakeImageAndLink And Not MoveToAction Then
        'Compute error vector
        SolutionRADeg = RA * Cos(Dec * PI / 180) * 15
        CurrentRADeg = CurrentRA * Cos(CurrentDec * PI / 180) * 15
        
        ErrorVectorMag = Sqr((SolutionRADeg - CurrentRADeg) ^ 2 + (Dec - CurrentDec) ^ 2)
        
        ErrorVectorDir = Misc.DoubleModulus((Misc.Atn360((Dec - CurrentDec), (SolutionRADeg - CurrentRADeg)) * 180 / PI), 360)
        
        Call AddToStatus("Pointing error vector = " & Format(ErrorVectorMag * 3600, "0.0") & " arcsec, " & Format(ErrorVectorDir, "0.0") & " degrees.")
    End If
    
    If Not Aborted And TakeImageAndLink Then
        If clsAction.SyncMode = MountSync Then
            Call Mount.Sync(RA, Dec, clsAction.SlewToOriginalLocation)
        ElseIf clsAction.SyncMode = Offset Then
            Call Mount.OffsetCoordinates(RA, Dec, clsAction.SlewToOriginalLocation)
        End If
        
        If (MoveToAction) Then
            frmMoveToRADec.txtRAH.Text = Fix(RA + (0.5 / 3600))
            frmMoveToRADec.txtRAM.Text = Fix((RA - CDbl(frmMoveToRADec.txtRAH.Text)) * 60# + (0.5 / 60))
            frmMoveToRADec.txtRAS.Text = Fix(((((RA - CDbl(frmMoveToRADec.txtRAH.Text)) * 60#) - CDbl(frmMoveToRADec.txtRAM.Text)) * 60#) + 0.5)
            
            If Dec < 0 And Dec > -1 Then
                frmMoveToRADec.txtDecD.Text = "-" & Fix(Dec + (0.5 / 3600))
            Else
                frmMoveToRADec.txtDecD.Text = Fix(Dec + (0.5 / 3600))
            End If
            frmMoveToRADec.txtDecM.Text = Fix(Abs(Dec - CDbl(frmMoveToRADec.txtDecD.Text)) * 60# + (0.5 / 60))
            frmMoveToRADec.txtDecS.Text = Fix((((Abs(Dec - CDbl(frmMoveToRADec.txtDecD.Text)) * 60#) - CDbl(frmMoveToRADec.txtDecM.Text)) * 60#) + 0.5)
        End If
    End If
    
    If Not Aborted And TakeImageAndLink And clsAction.RetryPlateSolveOnFailure And RetryCount = 0 And clsAction.SkipIfRetrySucceeds Then
        MainMod.SkipToNextMoveAction = True
    End If
    
    objCameraControl.BinXY = OriginalBin
    Call objCameraControl.SubFrame(False)
    objCameraControl.ExposureTime = OriginalExpTime
    objCameraControl.ImageReduction = LastReduction
    
    If TakeImageAndLink Then
        TakeImageAndLink = Not Aborted
    Else
        Call AddToStatus("Plate solve error!")
    End If
End Function

Public Sub CameraAbort()
    Autoguiding = False

    If Not (objCameraControl Is Nothing) Then
        Call objCameraControl.Abort
    End If
End Sub

Public Sub CameraSetup()
    If frmOptions.lstCameraControl.ListIndex = CameraControl.CCDSoft Then
        If (objCameraControl Is Nothing) Or (TypeName(objCameraControl) <> "clsCCDSoftControl") Then
            Set objCameraControl = New clsCCDSoftControl
        End If
    ElseIf frmOptions.lstCameraControl.ListIndex = CameraControl.MaxIm Then
        If (objCameraControl Is Nothing) Or (TypeName(objCameraControl) <> "clsMaxImControl") Then
            Set objCameraControl = New clsMaxImControl
        End If
    ElseIf frmOptions.lstCameraControl.ListIndex = CameraControl.CCDSoftAO Then
        If (objCameraControl Is Nothing) Or (TypeName(objCameraControl) <> "clsCCDSoftControlAO") Then
            Set objCameraControl = New clsCCDSoftControlAO
        End If
    ElseIf frmOptions.lstCameraControl.ListIndex = CameraControl.TheSkyX Then
        If (objCameraControl Is Nothing) Or (TypeName(objCameraControl) <> "clsTheSkyXCameraControl") Then
            Set objCameraControl = New clsTheSkyXCameraControl
        End If
    End If
    
    Call objCameraControl.ConnectToCamera

    Call frmTempGraph.SetupTempGraph
    frmTempGraph.cmdStop.Enabled = True
    
    If frmMain.mnuViewItem(0).Checked = True Then
        frmTempGraph.Timer1.Enabled = True
    End If
        
    UseAGBeforeTakeImageEvent = False
    Autoguiding = False
    
    LastGuideStarX = 0
    LastGuideStarY = 0
    StartRA = 25    'Never could be 25h
    StartDec = 91   'Never could be 91d
    StartMountSide = 99 'Never coule be 99
End Sub

Public Sub PutFilterDataIntoForms()
    ImageTypes = Array("Light", "Bias", "Dark", "Flat-Field")
    
    'done in the forms now...
End Sub

Public Sub CameraUnload()
    Set objCameraControl = Nothing
End Sub

Public Sub ForceFilterChange(Filter As Integer)
    If frmOptions.lstFocuserControl.ListIndex <> FocusControl.None And frmOptions.chkEnableFilterOffsets.Value = vbChecked Then
        Call AddToStatus("Current Filter Number = " & objCameraControl.FilterNumber)
        Call AddToStatus("New Filter Number = " & Filter)
        Call Focus.FocuserFilterOffset(objCameraControl.FilterNumber, Filter)
    End If

    If frmOptions.chkDisableForceFilterChange.Value <> vbChecked Then
        Call objCameraControl.TakeDummyImage(Filter)
    Else
        objCameraControl.FilterNumber = Filter
    End If
End Sub

Public Sub SetupFormsForCameraControlProgram()
    'done in the forms themselves now
End Sub

Public Sub CheckAndStopAutoguider()
    If Autoguiding Then
        Call AddToStatus("Stopping autoguider.")
        Autoguiding = False
        Call objCameraControl.StopAutoguider
        Call Wait(1)
    End If
End Sub

Public Sub RunAutoFlatAction(clsAction As AutoFlatAction)
    Dim CurrentExposureTime As Double
    Dim LastExposureTime As Double
    Dim DarkExposureTime As Variant
    Dim CurrentAverage As Double
    Dim LastAverage As Double
    Dim CurrentADUPerSecond As Double
    Dim CurrentTime As Date
    Dim LastTime As Date
    Dim LastADUPerSecond As Double
    Dim ADUPerSecondNow As Double
    Dim TimeNow As Date
    Dim Counter As Integer
    Dim RADec As Variant
    Dim Alt As Double
    Dim Azim As Double
    Dim SunSetTime As Date
    Dim FilterCounter As Long
    Dim RotationCounter As Long
    Dim SaveToPath As String
    Dim ExposureTimes As New Collection
    Dim SortedExposureTimes As New Collection
    Dim FileNameString As String
    Dim FailCounter As Integer
    
    Call AddToStatus("In Auto Flat Action.")
    
    If clsAction.FlatLocation = DawnSkyFlat Then
        'I might be here early
        'check alt of sun - wait for astronomical twilight
        'this is the time where the sun is 18 degrees below the horizon
        If DateDiff("s", Now, clsAction.ActualSkipToTime) > 0 Then
            Call AddToStatus("Waiting for dawn twilight...")
            Call Wait(DateDiff("s", Now, clsAction.ActualSkipToTime))
        End If
    ElseIf clsAction.FlatLocation = DuskSkyFlat Then
        'I might be here early
        'check alt of sun - wait for it to set
        If DateDiff("s", Now, clsAction.ActualSkipToTime) > 0 Then
            Call AddToStatus("Waiting for sun to set...")
            Call Wait(DateDiff("s", Now, clsAction.ActualSkipToTime))
        End If
    End If
    
    If Not Aborted And (clsAction.FlatLocation <> DoNotMove) Then _
        Call Mount.ConnectToTelescope
    
    'first put the mount where it needs to be
    If clsAction.FlatLocation = FlatParkMount Then
        Call Mount.ParkMount
    ElseIf clsAction.FlatLocation = DawnSkyFlat Then
        Alt = 85
        Azim = 270
    ElseIf clsAction.FlatLocation = DuskSkyFlat Then
        Alt = 85
        Azim = 90
    ElseIf clsAction.FlatLocation = FixedLocation Then
        If (clsAction.AltD < 0) Then
            Alt = clsAction.AltD - (clsAction.AltM / 60) - (clsAction.AltS / 3600)
        Else
            Alt = clsAction.AltD + (clsAction.AltM / 60) + (clsAction.AltS / 3600)
        End If
        Azim = clsAction.AzimD + (clsAction.AzimM / 60) + (clsAction.AzimS / 3600)
    End If
    
    If (clsAction.FlatLocation <> FlatParkMount) And (clsAction.FlatLocation <> DoNotMove) And Not Aborted Then
        Call AddToStatus("Slewing to flat target location.")
        Call Mount.MoveToAltAz(Alt, Azim)
        
        If Mount.objTele.CanSetTracking Then
            Mount.objTele.Tracking = False
        End If
    End If

    Autoguiding = False
    If Not Aborted Then
        Call objCameraControl.StopAutoguider
        Call Wait(1)
    End If
    
    'now do an exposure using the minimum time
    If Not Aborted Then
        objCameraControl.ImageType = 4
        objCameraControl.BinXY = clsAction.Bin + 1
        Call AddToStatus("Setting imager bin mode to " & clsAction.Bin + 1 & "x" & clsAction.Bin + 1 & ".")
    
        objCameraControl.ImageReduction = NoReduction
    End If

    CurrentExposureTime = clsAction.MinExpTime
    
    FilterCounter = 0
    Do
        If Aborted Then Exit Do
        FailCounter = 0
        
        RotationCounter = 0
        If clsAction.NumRotations > 0 Then
            If clsAction.FlipRotator <> 0 Then
                Call AddToStatus("Rotating to PA of " & (clsAction.Rotations(RotationCounter) + 180))
                Call Rotator.Rotate(clsAction.Rotations(RotationCounter) + 180)
            Else
                Call AddToStatus("Rotating to PA of " & clsAction.Rotations(RotationCounter))
                Call Rotator.Rotate(clsAction.Rotations(RotationCounter))
            End If
        End If
        
        If clsAction.NumFilters > 0 And Not Aborted Then
            Call AddToStatus("Setting filter to " & frmOptions.lstFilters.List(clsAction.Filters(FilterCounter)) & ".")
            
            'take a dummy exposure to get things setup
            If Not Aborted Then _
                Call Camera.ForceFilterChange(clsAction.Filters(FilterCounter))
        End If
        
        If Not Aborted Then
            If clsAction.SetupFrameSize = FullFrame Then
                Call AddToStatus("Setting imager to full frame.")
                Call objCameraControl.SubFrame(False)
            ElseIf clsAction.SetupFrameSize = HalfFrame Then
                Call AddToStatus("Setting imager to half frame.")
                With objCameraControl
                    Call .SubFrame(True, .WidthInPixels / 4, .WidthInPixels * 3 / 4, .HeightInPixels / 4, .HeightInPixels * 3 / 4)
                End With
            ElseIf clsAction.SetupFrameSize = QuarterFrame Then
                Call AddToStatus("Setting imager to quarter frame.")
                With objCameraControl
                    Call .SubFrame(True, .WidthInPixels * 3 / 8, .WidthInPixels * 5 / 8, .HeightInPixels * 3 / 8, .HeightInPixels * 5 / 8)
                End With
            End If
        End If
        
        If Not Aborted Then _
            Call objCameraControl.AutoSave(False)
                
        'Reset rate information - will be different with different filters
        LastADUPerSecond = 0
        
        Do
            If Aborted Then Exit Do
            
            If (clsAction.FlatLocation <> FlatParkMount) And (clsAction.FlatLocation <> DoNotMove) And Not Aborted And Not Mount.objTele.CanSetTracking Then
                'goto a set alt-az coordinate, repeatedly go to the same coordinates
                Call Mount.RecenterAltAz(Alt, Azim, False)
            End If
                        
            If Not Aborted Then
                objCameraControl.ExposureTime = CurrentExposureTime
                Call AddToStatus("Setting imager exposure time to " & Format(CurrentExposureTime, "0.000") & " seconds.")
            End If
            
            If Not Aborted Then
                Call AddToStatus("Starting imager exposure...")
                CurrentTime = Now
                
                Call objCameraControl.TakeImage
            End If
            
            Do Until objCameraControl.TakeImageComplete Or Aborted
                Call Wait(0.1)
            Loop
            
            If Aborted Then Exit Do
            
            Call AddToStatus("Exposure complete.  Computing average ADU.")
            CurrentAverage = objCameraControl.AverageADUOfExposure(FullFrame)   'Use full frame for the aveage - I want the average of whatever sub-frame is already defined
            Call AddToStatus("Average ADU = " & Format(CurrentAverage, "0"))
            Call AddToStatus("Minimum ADU = " & clsAction.AverageADU)
            Call AddToStatus("Maximum ADU = " & clsAction.MaximumADU)
            
            CurrentADUPerSecond = CurrentAverage / CurrentExposureTime
            Call AddToStatus("ADU/s = " & Format(CurrentADUPerSecond, "0"))
            
            If (CurrentAverage >= clsAction.AverageADU) And (CurrentAverage <= clsAction.MaximumADU) Then
                
                'I don't care about the change in brightness now, just roll with the current exposure time
                If clsAction.ContinuouslyAdjust = 0 Then
                    Call AddToStatus("Current exposure time good!  Taking Flats.")
                    
                    LastADUPerSecond = CurrentADUPerSecond
                    LastTime = CurrentTime
                    
                    Exit Do
                Else
                    Call AddToStatus("Current ADU rate good!")
                    
                    LastExposureTime = CurrentExposureTime
                    
                    If (LastADUPerSecond = 0) Then
                        'Haven't taken enough exposures, just guestimate based on this last exposure
                        CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / CurrentADUPerSecond
                        
                        'Round
                        If (clsAction.MinExpTime >= 0.1) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                        ElseIf (clsAction.MinExpTime >= 0.01) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                        Else 'If (clsAction.MinExpTime >= 0.001) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                        End If
                    Else
                        TimeNow = Now
                        ADUPerSecondNow = ((CurrentADUPerSecond - LastADUPerSecond) / CDbl(DateDiff("s", LastTime, CurrentTime))) * CDbl(DateDiff("s", LastTime, TimeNow)) + LastADUPerSecond
                        
                        If (ADUPerSecondNow < 0) Then
                            'this is an odd case - the ADU/s linear fit has failed and the expected ADU/s is now negative.  Just use the previous exposure time
                            'just guestimate based on this last exposure
                            CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / CurrentADUPerSecond
                        Else
                            CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / ADUPerSecondNow
                        End If
                        
                        'Round
                        If (clsAction.MinExpTime >= 0.1) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                        ElseIf (clsAction.MinExpTime >= 0.01) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                        Else 'If (clsAction.MinExpTime >= 0.001) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                        End If
                    
                        If (CurrentExposureTime < 0) Then
                            'this shouldn't be possible, something is amiss with the change in ADU/s.
                            'Just guestimage based on this last exposure alone
                            CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / CurrentADUPerSecond
                            
                            'Round
                            If (clsAction.MinExpTime >= 0.1) Then
                                CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                            ElseIf (clsAction.MinExpTime >= 0.01) Then
                                CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                            Else 'If (clsAction.MinExpTime >= 0.001) Then
                                CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                            End If
                        End If
                    End If
                    
                    Call AddToStatus("New exposure time computed to be " & Format(CurrentExposureTime, "0.000") & " seconds")
                    
                    LastADUPerSecond = CurrentADUPerSecond
                    LastTime = CurrentTime
                    
                    If ((CurrentExposureTime < clsAction.MinExpTime) Or (CurrentExposureTime > clsAction.MaxExpTime)) And _
                        ((CurrentAverage >= clsAction.AverageADU) And (CurrentAverage <= clsAction.MaximumADU)) And Not Aborted Then
                        
                        'New exposure out of bounds, but last exposure was okay - just roll with last exposure time
                        Call AddToStatus("New exposure time out of bounds, but last exposure okay.")
                        Call AddToStatus("Using previous exposure time of " & Format(LastExposureTime, "0.000") & " seconds")
                        CurrentExposureTime = LastExposureTime
                    End If
                
                    Call AddToStatus("Taking Flats.")
                    Exit Do
                End If
                
            Else
                If (LastADUPerSecond = 0) Then
                    'Haven't taken enough exposures, just guestimate based on this last exposure
                    CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / CurrentADUPerSecond
                
                    'Round
                    If (clsAction.MinExpTime >= 0.1) Then
                        CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                    ElseIf (clsAction.MinExpTime >= 0.01) Then
                        CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                    Else 'If (clsAction.MinExpTime >= 0.001) Then
                        CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                    End If
                Else
                    TimeNow = Now
                    ADUPerSecondNow = ((CurrentADUPerSecond - LastADUPerSecond) / CDbl(DateDiff("s", LastTime, CurrentTime))) * CDbl(DateDiff("s", LastTime, TimeNow)) + LastADUPerSecond
                    
                    CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / ADUPerSecondNow
                    
                    'Round
                    If (clsAction.MinExpTime >= 0.1) Then
                        CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                    ElseIf (clsAction.MinExpTime >= 0.01) Then
                        CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                    Else 'If (clsAction.MinExpTime >= 0.001) Then
                        CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                    End If
                    
                    If (CurrentExposureTime < clsAction.MinExpTime) Or (CurrentExposureTime > clsAction.MaxExpTime) Then
                        'Current ADU/s rate is not optimal, just use the single image calculation
                        CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / CurrentADUPerSecond
                    
                        'Round
                        If (clsAction.MinExpTime >= 0.1) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                        ElseIf (clsAction.MinExpTime >= 0.01) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                        Else 'If (clsAction.MinExpTime >= 0.001) Then
                            CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                        End If
                    End If
                End If
                                
                Call AddToStatus("New exposure time computed to be " & Format(CurrentExposureTime, "0.000") & " seconds")
                
                LastADUPerSecond = CurrentADUPerSecond
                LastTime = CurrentTime
            End If
            
            If ((CurrentExposureTime < clsAction.MinExpTime) Or _
                ((CurrentAverage > clsAction.MaximumADU) And (CurrentExposureTime = clsAction.MinExpTime))) _
                And Not Aborted Then
                
                'needed exposure is too short
                'if this is a twilight flat, I can wait for the sky to dim
                'if not, I need to give up
                If clsAction.FlatLocation = DawnSkyFlat Then
                    Call AddToStatus("Sky is too bright for current filter to continue.")
                    Call AddToStatus("Skipping to next filter.")
                    
                    'CurrentExposureTime = clsAction.MinExpTime
                    
                    Exit Do
                Else
                    If (FailCounter = 0) Then
                        'first time getting here, store LastAverage
                        LastAverage = CurrentAverage
                    ElseIf ((LastAverage - CurrentAverage) > 100) Then
                        'Getting closer to where I need to be - reset FailCounter
                        FailCounter = 0
                        LastAverage = CurrentAverage
                    End If
            
                    CurrentExposureTime = clsAction.MinExpTime
                    FailCounter = FailCounter + 1
                End If
            ElseIf ((CurrentExposureTime > clsAction.MaxExpTime) Or _
                ((CurrentAverage < clsAction.AverageADU) And (CurrentExposureTime = clsAction.MaxExpTime))) _
                And Not Aborted Then
                
                'needed exposure is too long
                'if this is a dawn flat, I can wait for the sky to brighten
                'if not, I need to give up
                If clsAction.FlatLocation = DuskSkyFlat Then
                    Call AddToStatus("Sky is too dark for current filter to continue.")
                    Call AddToStatus("Skipping to next filter.")
                    
                    'CurrentExposureTime = clsAction.MaxExpTime
                    
                    Exit Do
                Else
                    If (FailCounter = 0) Then
                        'first time getting here, store LastAverage
                        LastAverage = CurrentAverage
                    ElseIf ((CurrentAverage - LastAverage) > 100) Then
                        'Getting closer to where I need to be - reset FailCounter
                        FailCounter = 0
                        LastAverage = CurrentAverage
                    End If
            
                    CurrentExposureTime = clsAction.MaxExpTime
                    FailCounter = FailCounter + 1
                End If
            End If
        Loop Until Aborted Or ((FailCounter > 10) And (clsAction.FlatLocation <> DuskSkyFlat) And (clsAction.FlatLocation <> DawnSkyFlat))
        
        If ((FailCounter > 10) And (clsAction.FlatLocation <> DuskSkyFlat) And (clsAction.FlatLocation <> DawnSkyFlat)) Then
            Exit Do
        End If
            
        If Not Aborted And CurrentExposureTime <= clsAction.MaxExpTime And CurrentExposureTime >= clsAction.MinExpTime Then
            If Not Aborted Then
                If clsAction.FrameSize = FullFrame Then
                    Call AddToStatus("Setting imager to full frame.")
                    Call objCameraControl.SubFrame(False)
                ElseIf clsAction.FrameSize = HalfFrame Then
                    Call AddToStatus("Setting imager to half frame.")
                    With objCameraControl
                        Call .SubFrame(True, .WidthInPixels / 4, .WidthInPixels * 3 / 4, .HeightInPixels / 4, .HeightInPixels * 3 / 4)
                    End With
                ElseIf clsAction.FrameSize = QuarterFrame Then
                    Call AddToStatus("Setting imager to quarter frame.")
                    With objCameraControl
                        Call .SubFrame(True, .WidthInPixels * 3 / 8, .WidthInPixels * 5 / 8, .HeightInPixels * 3 / 8, .HeightInPixels * 5 / 8)
                    End With
                End If
            End If
            
            For Counter = 1 To clsAction.NumExp
                If Aborted Then Exit For
                
                If (clsAction.FlatLocation <> FlatParkMount) And (clsAction.FlatLocation <> DoNotMove) And Not Aborted And Not Mount.objTele.CanSetTracking Then
                    'goto a set alt-az coordinate, repeatedly go to the same coordinates
                    Call Mount.RecenterAltAz(Alt, Azim, False)
                End If
            
                If clsAction.AutosaveExposure = vbChecked And Not Aborted And CurrentExposureTime <= clsAction.MaxExpTime And CurrentExposureTime >= clsAction.MinExpTime Then
                    If clsAction.UseGlobalSaveToLocation Then
                        SaveToPath = frmOptions.txtSaveTo
                    Else
                        SaveToPath = Misc.FixFileName(clsAction.FileSavePath, True)
                    End If
                    
                    'Make sure SaveToPath exists
                    On Error Resume Next
                    Call ChDir(SaveToPath)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                        Call Misc.CreatePath(SaveToPath)
                    End If
                    On Error GoTo 0
                    
                    If InStr(UCase(clsAction.FileNamePrefix), "<FILTER>") = 0 And clsAction.NumFilters > 0 Then
                        FileNameString = "<Filter>"
                    Else
                        FileNameString = ""
                    End If
                    
                    FileNameString = FileNameString & clsAction.FileNamePrefix
                    
                    If InStr(UCase(clsAction.FileNamePrefix), "<ROTATION>") = 0 And clsAction.NumRotations > 0 Then
                        FileNameString = FileNameString & "<Rotation>"
                    End If
                    
                    Call objCameraControl.AutoSave(True, SaveToPath, Misc.FixFileName(FileNameString))
                End If
                
                If Not Aborted Then
                    objCameraControl.ExposureTime = CurrentExposureTime
                    Call AddToStatus("Setting imager exposure time to " & Format(CurrentExposureTime, "0.000") & " seconds.")
                
                    If clsAction.TakeMatchingDarks Then
                        'Add this exposure time to the list of exposure times
                        ExposureTimes.Add CurrentExposureTime
                    End If
                End If
                
                If Not Aborted Then
                    Call AddToStatus("Starting imager exposure #" & Counter)
                    CurrentTime = Now
                    Call objCameraControl.TakeImage
                End If
                
                Do Until objCameraControl.TakeImageComplete Or Aborted
                    Call Wait(0.1)
                Loop
                
                If Not Aborted Then
                    Call AddToStatus("Exposure complete.  Computing average ADU.")
                    
                    'Now use the size of the original sub-frame.  This will ensure I'm comparing apples-to-apples
                    CurrentAverage = objCameraControl.AverageADUOfExposure(clsAction.SetupFrameSize)
                    
                    Call AddToStatus("Average ADU = " & Format(CurrentAverage, "0"))
                
                    CurrentADUPerSecond = CurrentAverage / CurrentExposureTime
                    Call AddToStatus("ADU/s = " & Format(CurrentADUPerSecond, "0"))
                End If
                
                If Counter = clsAction.NumExp And RotationCounter < (clsAction.NumRotations - 1) Then
                    Counter = 0 'Set to 0 so that it will go to 1 after the "Next Counter" line
                    RotationCounter = RotationCounter + 1
                    
                    Call AddToStatus("Rotating to PA of " & clsAction.Rotations(RotationCounter))
                    Call Rotator.Rotate(clsAction.Rotations(RotationCounter))
                End If
                
                If Counter < clsAction.NumExp Then
                    If clsAction.ContinuouslyAdjust = 1 Then
                        Call AddToStatus("Recomputing exposure time...")
                        
                        TimeNow = Now
                        ADUPerSecondNow = ((CurrentADUPerSecond - LastADUPerSecond) / CDbl(DateDiff("s", LastTime, CurrentTime))) * CDbl(DateDiff("s", LastTime, TimeNow)) + LastADUPerSecond
                        
                        CurrentExposureTime = (clsAction.AverageADU + clsAction.MaximumADU) / 2 / ADUPerSecondNow
    
                        LastADUPerSecond = CurrentADUPerSecond
                        LastTime = CurrentTime
    
                        If Not Aborted Then
                            'Round
                            If (clsAction.MinExpTime >= 0.1) Then
                                CurrentExposureTime = Int(CurrentExposureTime * 10 + 0.5) / 10
                            ElseIf (clsAction.MinExpTime >= 0.01) Then
                                CurrentExposureTime = Int(CurrentExposureTime * 100 + 0.5) / 100
                            Else 'If (clsAction.MinExpTime >= 0.001) Then
                                CurrentExposureTime = Int(CurrentExposureTime * 1000 + 0.5) / 1000
                            End If
                            Call AddToStatus("New exposure time computed to be " & Format(CurrentExposureTime, "0.000") & " seconds")
                        End If
                    End If
                     
                    If CurrentExposureTime < clsAction.MinExpTime And Not Aborted Then
                        'exposure is too fast
                        'if this is a dusk flat, I can wait for the sky to dim
                        'if not, I need to give up
                        If clsAction.FlatLocation = DawnSkyFlat Then
                            Call AddToStatus("Sky is too bright for current filter to continue.")
                            Call AddToStatus("Skipping to next filter.")
                            
                            CurrentExposureTime = clsAction.MinExpTime
                            
                            Exit For
                        Else
                            If (FailCounter = 0) Then
                                'first time getting here, store LastAverage
                                LastAverage = CurrentAverage
                            ElseIf ((LastAverage - CurrentAverage) > 100) Then
                                'Getting closer to where I need to be - reset FailCounter
                                FailCounter = 0
                                LastAverage = CurrentAverage
                            End If
                            
                            FailCounter = FailCounter + 1
                            CurrentExposureTime = clsAction.MinExpTime
                        End If
                    ElseIf CurrentExposureTime > clsAction.MaxExpTime And Not Aborted Then
                        'exposure is too long
                        'if this is a dawn flat, I can wait for the sky to brighten
                        'if not, I need to give up
                        If clsAction.FlatLocation = DuskSkyFlat Then
                            Call AddToStatus("Sky is too dark for current filter to continue.")
                            Call AddToStatus("Skipping to next filter.")
                            
                            CurrentExposureTime = clsAction.MaxExpTime
                            
                            Exit For
                        Else
                            If (FailCounter = 0) Then
                                'first time getting here, store LastAverage
                                LastAverage = CurrentAverage
                            ElseIf ((CurrentAverage - LastAverage) > 100) Then
                                'Getting closer to where I need to be - reset FailCounter
                                FailCounter = 0
                                LastAverage = CurrentAverage
                            End If
                            
                            FailCounter = FailCounter + 1
                            CurrentExposureTime = clsAction.MaxExpTime
                        End If
                    End If
                End If
            Next Counter
        Else
            If CurrentExposureTime < clsAction.MinExpTime Then
                If (FailCounter = 0) Then
                    'first time getting here, store LastAverage
                    LastAverage = CurrentAverage
                ElseIf ((LastAverage - CurrentAverage) > 100) Then
                    'Getting closer to where I need to be - reset FailCounter
                    FailCounter = 0
                    LastAverage = CurrentAverage
                End If
                                
                CurrentExposureTime = clsAction.MinExpTime
                FailCounter = FailCounter + 1
            End If
            
            If CurrentExposureTime > clsAction.MaxExpTime Then
                If (FailCounter = 0) Then
                    'first time getting here, store LastAverage
                    LastAverage = CurrentAverage
                ElseIf ((CurrentAverage - LastAverage) > 100) Then
                    'Getting closer to where I need to be - reset FailCounter
                    FailCounter = 0
                    LastAverage = CurrentAverage
                End If
                
                CurrentExposureTime = clsAction.MaxExpTime
                FailCounter = FailCounter + 1
            End If
        End If
        
        FilterCounter = FilterCounter + 1
    Loop Until FilterCounter >= clsAction.NumFilters Or Aborted Or ((FailCounter > 10) And (clsAction.FlatLocation <> DuskSkyFlat) And (clsAction.FlatLocation <> DawnSkyFlat))
    
    If ((FailCounter > 10) And (clsAction.FlatLocation <> DuskSkyFlat) And (clsAction.FlatLocation <> DawnSkyFlat)) Then
        Call AddToStatus("Cannot achieve desired exposure time.  Giving up...")
    Else
        If clsAction.TakeMatchingDarks And Not Aborted Then
            Call AddToStatus("Flats complete - taking matching darks.")
                    
            'Sort the Exposure times
            For Each DarkExposureTime In ExposureTimes
                If SortedExposureTimes.Count = 0 Then
                    'just add this one
                    SortedExposureTimes.Add DarkExposureTime
                ElseIf DarkExposureTime < SortedExposureTimes(1) Then
                    'less than the first entry
                    SortedExposureTimes.Add DarkExposureTime, , 1
                Else
                    'find the place in the list
                    Counter = 1
                    Do While Counter <= SortedExposureTimes.Count
                        If DarkExposureTime < SortedExposureTimes(Counter) Then
                            Exit Do
                        End If
                        Counter = Counter + 1
                    Loop
                    
                    If Counter > SortedExposureTimes.Count Then
                        SortedExposureTimes.Add DarkExposureTime
                    Else
                        SortedExposureTimes.Add DarkExposureTime, , Counter
                    End If
                End If
            Next DarkExposureTime
            
            If Not Aborted Then
                If clsAction.FrameSize = FullFrame Then
                    Call AddToStatus("Setting imager to full frame.")
                    Call objCameraControl.SubFrame(False)
                ElseIf clsAction.FrameSize = HalfFrame Then
                    Call AddToStatus("Setting imager to half frame.")
                    With objCameraControl
                        Call .SubFrame(True, .WidthInPixels / 4, .WidthInPixels * 3 / 4, .HeightInPixels / 4, .HeightInPixels * 3 / 4)
                    End With
                ElseIf clsAction.FrameSize = QuarterFrame Then
                    Call AddToStatus("Setting imager to quarter frame.")
                    With objCameraControl
                        Call .SubFrame(True, .WidthInPixels * 3 / 8, .WidthInPixels * 5 / 8, .HeightInPixels * 3 / 8, .HeightInPixels * 5 / 8)
                    End With
                End If
            End If
            
            If Not Aborted Then
                objCameraControl.ImageType = 3
                Call AddToStatus("Setting image type to DARK.")
            End If
            
            'all sorted now start taking exposures
            'this will ensure that I take the first exposure
            CurrentExposureTime = clsAction.MinExpTime - clsAction.DarkFrameTolerance
            For Each DarkExposureTime In SortedExposureTimes
                If DarkExposureTime >= CurrentExposureTime + clsAction.DarkFrameTolerance Then
                    'ok to take this exposure!
                    CurrentExposureTime = DarkExposureTime
                    
                    If clsAction.AutosaveExposure = vbChecked And Not Aborted Then
                        If clsAction.UseGlobalSaveToLocation Then
                            SaveToPath = frmOptions.txtSaveTo
                        Else
                            SaveToPath = clsAction.FileSavePath
                        End If
                        
                        'Make sure SaveToPath exists
                        On Error Resume Next
                        Call ChDir(SaveToPath)
                        If Err.Number <> 0 Then
                            On Error GoTo 0
                            Call Misc.CreatePath(SaveToPath)
                        End If
                        On Error GoTo 0
                        
                        If InStr(UCase(clsAction.FileNamePrefix), UCase("<ImageType>")) = 0 Then
                            FileNameString = "<ImageType>"
                        Else
                            FileNameString = ""
                        End If
                        
                        If InStr(UCase(clsAction.FileNamePrefix), UCase("<ExposureTime>")) = 0 Then
                            FileNameString = "<ExposureTime>"
                        Else
                            FileNameString = ""
                        End If
                        
                        FileNameString = clsAction.FileNamePrefix & FileNameString
                    End If
            
                    For Counter = 1 To clsAction.NumberOfDarksPerFlat
                        If Aborted Then Exit For
                        
                        If (clsAction.FlatLocation <> FlatParkMount) And (clsAction.FlatLocation <> DoNotMove) And Not Aborted And Not Mount.objTele.CanSetTracking Then
                            'goto a set alt-az coordinate, repeatedly go to the same coordinates
                            Call Mount.RecenterAltAz(Alt, Azim, False)
                        End If
                    
                        If Not Aborted Then
                            objCameraControl.ExposureTime = DarkExposureTime
                            Call AddToStatus("Setting imager exposure time to " & Format(DarkExposureTime, "0.0") & " seconds.")
                            Call objCameraControl.AutoSave(True, SaveToPath, Misc.FixFileName(FileNameString))
                        End If
                        
                        If Not Aborted Then
                            Call AddToStatus("Starting dark exposure #" & Counter)
                            CurrentTime = Now
                            Call objCameraControl.TakeImage
                        End If
                        
                        Do Until objCameraControl.TakeImageComplete Or Aborted
                            Call Wait(0.1)
                        Loop
                        
                        If Not Aborted Then
                            Call AddToStatus("Exposure complete.")
                        End If
                    Next Counter
                    
                End If
            Next DarkExposureTime
        End If
    End If
    
    If Mount.objTele.CanSetTracking Then
        If Not Mount.objTele.Tracking And (clsAction.FlatLocation <> FlatParkMount) And (clsAction.FlatLocation <> DoNotMove) Then
            Mount.objTele.Tracking = True
        End If
    End If
    
    Call AddToStatus("Automatic Flat Action complete.")
End Sub

Public Sub CheckAutoguiderError()
    'Presently, only for TheSkyX.  May also use for CCDSoft w/AO
    If TypeName(objCameraControl) = "clsTheSkyXCameraControl" Then
        Call objCameraControl.CheckAndRecordGuideError
    End If
End Sub
