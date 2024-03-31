Attribute VB_Name = "Rotator"
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

Enum RotatorControl
    None = 0
    Pyxis = 1
    TAKometer = 2
    PIR = 3
    ASCOMRotator = 4
    Manual = 5
End Enum

Public CurrentAngle As Double

Public RotatorConnected As Boolean

Private CurrentRotatorAngle As Integer

Private objRotator As Object

Public Sub RotatorSetup()
    Dim GuiderCalAngle As Double
    
    If frmOptions.lstRotator.ListIndex = RotatorControl.Pyxis Then
        Set objRotator = New clsRotatorPyxis
    ElseIf frmOptions.lstRotator.ListIndex = RotatorControl.TAKometer Then
        Set objRotator = New clsRotatorTAKometer
    ElseIf frmOptions.lstRotator.ListIndex = RotatorControl.PIR Then
        Set objRotator = New clsRotatorPIR
    ElseIf frmOptions.lstRotator.ListIndex = RotatorControl.ASCOMRotator Then
        Set objRotator = New clsRotatorASCOM
    ElseIf frmOptions.lstRotator.ListIndex = RotatorControl.Manual Then
        Set objRotator = New clsRotatorSim
    End If
    
    objRotator.RotatorCOMPort = CInt(Settings.RotatorCOMNumber)
    
    RotatorConnected = objRotator.ConnectToRotator
    
    If Not RotatorConnected Then Exit Sub
    
    If frmOptions.chkReverseRotatorDirection = vbChecked Then
        CurrentRotatorAngle = Misc.DoubleModulus(360 - objRotator.CurrentAngle, 360)
    Else
        CurrentRotatorAngle = objRotator.CurrentAngle
    End If
    
    CurrentAngle = Misc.DoubleModulus(CurrentRotatorAngle + Settings.HomeRotationAngle, 360)
        
    If frmOptions.optMountType(0).Value Then
        If Mount.MountSide = EastSide And frmOptions.optGuiderCal(0).Value Then
            Call AddToStatus("In western sky - angles computed in eastern sky, adding 180 degrees.")
            CurrentAngle = Misc.DoubleModulus(CurrentAngle + 180, 360)
        ElseIf Mount.MountSide = WestSide And frmOptions.optGuiderCal(1).Value Then
            Call AddToStatus("In eastern sky - angles computed in western sky, adding 180 degrees.")
            CurrentAngle = Misc.DoubleModulus(CurrentAngle + 180, 360)
        End If
    End If
    
    Call AddToStatus("Actual position angle is " & Format(CurrentAngle, "0.00") & " degrees.")
        
    If frmOptions.chkRotateTheSkyFOVI.Value = vbChecked Then
        Call Planetarium.SetFOVIPositionAngle(CurrentAngle)
    End If
    
    Call CheckAndChangeGuiderAngle
    
    'This is necessary to correct for an issue where the scope finished on one side, but is now on the other side of the meridian.
    'Actually appears to only be necessary when CCD Commander finishes in the west and then starts in the east.
    'So, FlipGuiderCalAngle is set appropriately above.
    '
    'Actually this may only be necessary with the Paramount - it certainly isn't elsewhere from what I'm seeing now...
    '!!!!!!!!!!!!!!!!!!!!!!!! May not be necessary at all - this may be causing problems !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'    If (FlipGuiderCalAngle) And Not Camera.objCameraControl.ReverseXNecessary Then
'            Settings.GuiderCalibrationAngle = Misc.DoubleModulus(Settings.GuiderCalibrationAngle - 180, 360)
'            frmOptions.txtGuiderCalAngle.Text = Format(Settings.GuiderCalibrationAngle, "0.00")
'            Call SaveMySetting("ProgramSettings", "GuiderCalAngle", frmOptions.txtGuiderCalAngle.Text)
'    End If
End Sub

Private Sub CheckAndChangeGuiderAngle()
    Dim GuiderCalAngle As Double
    
    If frmOptions.chkGuiderRotates.Value = vbChecked Then
        GuiderCalAngle = Settings.GuiderCalibrationAngle
        
        If frmOptions.optMountType(0).Value Then
            If (((Mount.MountSide = WestSide) And frmOptions.optGuiderCal(1).Value) Or _
                ((Mount.MountSide = EastSide) And frmOptions.optGuiderCal(0).Value)) And _
                Not Camera.objCameraControl.ReverseXNecessary Then

                'DirectGuide doesn't need to change the calibration angle after a meridian flip
                Call AddToStatus("Rotating Guider Cal Angle by 180 degrees.")
                GuiderCalAngle = Misc.DoubleModulus(GuiderCalAngle - 180, 360)
                Settings.GuiderCalibrationAngle = GuiderCalAngle
                frmOptions.txtGuiderCalAngle.Text = Format(Settings.GuiderCalibrationAngle, "0.00")
                Call SaveMySetting("ProgramSettings", "GuiderCalAngle", frmOptions.txtGuiderCalAngle.Text)
            End If
        End If

        If Abs(CurrentAngle - GuiderCalAngle) > 2 And _
            Abs(CurrentAngle - GuiderCalAngle - 360) > 2 Then
            '2 is arbitraty
            'Recompute guider angle
        
            Call AddToStatus("Guider calibration angle different than current angle.")
            Call AddToStatus("Recomputing calibration coefficients.")
            Call AddToStatus("Old Guider Cal Angle = " & Format(Settings.GuiderCalibrationAngle, "0.00"))
            Call AddToStatus("Current Angle = " & Format(CurrentAngle, "0.00"))
            Call Camera.objCameraControl.RecomputeGuiderCalibration(CurrentAngle, GuiderCalAngle)
            
            'Save current angle as guider cal angle
'            If (((Mount.MountSide = WestSide) And frmOptions.optGuiderCal(1).Value) Or _
'                ((Mount.MountSide = EastSide) And frmOptions.optGuiderCal(0).Value)) And _
'                Not Camera.objCameraControl.ReverseXNecessary Then
'
'                'This makes the angle match the opposite side of the meridian correctly
'                frmOptions.txtGuiderCalAngle.Text = Format(Misc.DoubleModulus(CurrentAngle - 180, 360), "0.00")
'                Settings.GuiderCalibrationAngle = Misc.DoubleModulus(CurrentAngle - 180, 360)
'            Else
                frmOptions.txtGuiderCalAngle.Text = Format(CurrentAngle, "0.00")
                Settings.GuiderCalibrationAngle = CurrentAngle
'            End If
            
            Call SaveMySetting("ProgramSettings", "GuiderCalAngle", frmOptions.txtGuiderCalAngle.Text)
        End If
        
        If (((Mount.MountSide = WestSide) And frmOptions.optGuiderCal(1).Value) Or _
            ((Mount.MountSide = EastSide) And frmOptions.optGuiderCal(0).Value)) And _
            frmOptions.optMountType(0).Value Then
            
            If Camera.objCameraControl.ReverseXNecessary Then
                'Guider was calibrated on the opposite side of the meridian!
                'Need to flip the Y axis corrections
                Call AddToStatus("Guider calibrated on opposite side of meridian.")
                Call AddToStatus("Flipping Y-axis corrections.")
                Call Camera.objCameraControl.ReverseYGuiderDirections
            End If
            
            'Now switch the "calibrated on" selection
            If frmOptions.optGuiderCal(0).Value Then
                Call AddToStatus("Changing autoguider calibration side to Western Sky.")
                frmOptions.optGuiderCal(1).Value = True
                Call SaveMySetting("MountParameters", "GEMGuideCal", "East")
            Else
                Call AddToStatus("Changing autoguider calibration side to Eastern Sky.")
                frmOptions.optGuiderCal(0).Value = True
                Call SaveMySetting("MountParameters", "GEMGuideCal", "West")
            End If
            
            'Now rotate the "Home" position by 180 degrees
            Call AddToStatus("Rotating Home Rotation position by 180 degrees.")
            Settings.HomeRotationAngle = Misc.DoubleModulus(Settings.HomeRotationAngle + 180, 360)
            frmOptions.txtHomeRotationAngle.Text = Format(Settings.HomeRotationAngle, "0.00")
            Call SaveMySetting("ProgramSettings", "RotatorHomeAngle", frmOptions.txtHomeRotationAngle.Text)
        End If
    Else
        Call AddToStatus("Guider doesn't rotate, no changes to guider calibration necessary.")
    End If
End Sub

Public Sub RotateAction(objRotate As RotatorAction)
    Call Rotate(objRotate.RotationAngle)
End Sub

Public Sub Rotate(ByVal PositionAngle As Double)
    Dim NewAngle As Double
    Dim Counter As Integer
    Dim AngleDiff As Double
    
    If frmOptions.lstRotator.ListIndex > 0 And RotatorConnected Then
        Call AddToStatus("In Rotate Function. Current PA = " & Format(CurrentAngle, "0.00") & ", New PA = " & Format(PositionAngle, "0.00"))
        
        If frmOptions.optMountType(0).Value Then
            If Mount.MountSide = EastSide And frmOptions.optGuiderCal(0).Value Then
                Call AddToStatus("In western sky - angles computed in eastern sky, adding 180 degrees.")
                PositionAngle = Misc.DoubleModulus(PositionAngle + 180, 360)
                Call AddToStatus("Adjusted PA = " & Format(PositionAngle, "0.00"))
            ElseIf Mount.MountSide = WestSide And frmOptions.optGuiderCal(1).Value Then
                Call AddToStatus("In eastern sky - angles computed in western sky, adding 180 degrees.")
                PositionAngle = Misc.DoubleModulus(PositionAngle + 180, 360)
                Call AddToStatus("Adjusted PA = " & Format(PositionAngle, "0.00"))
            End If
        End If
        
        NewAngle = Misc.DoubleModulus(PositionAngle - Settings.HomeRotationAngle, 360)
        
        Call AddToStatus("Current Rotator Angle = " & Format(CurrentRotatorAngle, "0.00"))
        Call AddToStatus("New Rotator Angle = " & Format(NewAngle, "0.00"))
        
        AngleDiff = Misc.DoubleModulus(NewAngle - CurrentRotatorAngle, 360)
        
        If AngleDiff < 1 Or AngleDiff > 359 Then
            Call AddToStatus("No rotator movement necessary.")
        Else
            Call AddToStatus("Moving rotator...")
            
            If frmOptions.chkReverseRotatorDirection = vbChecked Then
                objRotator.CurrentAngle = Misc.DoubleModulus(360 - NewAngle, 360)
            Else
                objRotator.CurrentAngle = NewAngle
            End If
            
            CurrentRotatorAngle = NewAngle
        End If
        
        If (Mount.MountSide = EastSide And frmOptions.optGuiderCal(0).Value) Or _
            (Mount.MountSide = WestSide And frmOptions.optGuiderCal(1).Value) Then
            CurrentAngle = Misc.DoubleModulus(PositionAngle - 180, 360)
        Else
            CurrentAngle = PositionAngle
        End If
        
        If frmOptions.chkRotateTheSkyFOVI.Value = vbChecked Then
            Call Planetarium.SetFOVIPositionAngle(CurrentAngle)
        End If
        
        Call CheckAndChangeGuiderAngle
    Else
        Call AddToStatus("No rotator enabled.  Check the program settings.")
    End If
End Sub

Public Function GetCurrentAngleCorrectedForPlateSolve() As Double
    Dim TempAngle As Double
    
    'Couple of things to do here.
    'CCDSoft reports PA as 270-90.  270 is reported for an image with West up.  90 is reportd for an image with East up.
    '   If it is outside the range of 270-90, the image is rotated 180 degrees and then solved.
    'PinPoint reports PA as -90 to +90.  -90 is reported for an image with West up. +90 is reported for an image with East up.
    '   Like CCDSoft, if it is outside the range of -90 to +90, the image is rotated 180 degrees and then solved.
    'SOOOO... PA is represented with 0 at North, 90 at East, 180 at South, and 270 at West.
    
    Call AddToStatus("Adjusting current PA for Plate Solve.  CurrentAngle = " & Format(CurrentAngle, "0.00"))
    
    'For CCDSoft, I simply need to add 180 and modulus 360 if the current angle is between 90 and 270.
    If frmOptions.lstPlateSolve.ListIndex = 0 Then
        If CurrentAngle > 90 And CurrentAngle < 270 Then
            TempAngle = Misc.DoubleModulus(CurrentAngle + 180, 360)
        Else
            TempAngle = CurrentAngle
        End If
    'For PinPoint I need to flip and add (as necessary)
    Else 'If frmOptions.lstPlateSolve.ListIndex = 1 Or 2 Then
        If CurrentAngle > 90 And CurrentAngle < 270 Then
            TempAngle = Misc.DoubleModulus(CurrentAngle + 180, 360)
        Else
            TempAngle = CurrentAngle
        End If
        
        If TempAngle <= 90 Then
            TempAngle = -TempAngle
        ElseIf TempAngle >= 270 Then
            TempAngle = Misc.DoubleModulus(-TempAngle, 360)
        End If
    End If
    
    Call AddToStatus("Adjusted PA = " & Format(TempAngle, "0.00"))
    
    GetCurrentAngleCorrectedForPlateSolve = TempAngle
End Function

Public Function GetCurrentRotatorAngle() As Double
    On Error Resume Next
    GetCurrentRotatorAngle = objRotator.CurrentAngle
    If Err.Number <> 0 Then
        GetCurrentRotatorAngle = 0
    End If
    On Error GoTo 0
End Function

Public Sub RotatorUnload()
    Set objRotator = Nothing
End Sub
