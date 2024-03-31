Attribute VB_Name = "TempControl"
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

Public Sub IntelligentTempControl(clsAction As IntelligentTempAction)
    Dim CurrentStep As Long
    Dim MyTime As Date
    Dim ElapsedTime As Double
    Dim Counter As Long
    Dim Temp1 As Double
    Dim Temp2 As Double
    Dim Cooler1 As Double
    Dim Cooler2 As Double
    
    frmTempGraph.Timer1.Enabled = False
    
    If clsAction.FanOn = 1 Then
        Call AddToStatus("Turning on Fan.")
        Camera.objCameraControl.FanOn = True
    Else
        Call AddToStatus("Turning off Fan.")
        Camera.objCameraControl.FanOn = False
    End If
    
    If clsAction.CoolerOn And clsAction.IntelligentAction = 1 Then
        'first determine where I am in the list - this is in case the cooler is already on, or it is colder than the start temp
        For Counter = 0 To clsAction.NumTemps
            If Camera.objCameraControl.Temperature > clsAction.DesiredTemperatures(Counter) Then
                CurrentStep = Counter
                Exit For
            End If
        Next Counter
        
        If (Counter >= clsAction.NumTemps) Then
            CurrentStep = clsAction.NumTemps
        End If
        
        'now force the cooler on and go to the next temp in the list
        Call AddToStatus("Turning on cooler.")
        Camera.objCameraControl.CoolerState = True
        
        Do While CurrentStep <= clsAction.NumTemps And Not Aborted
            Call AddToStatus("Setting cooler = " & clsAction.DesiredTemperatures(CurrentStep))
            
            Camera.objCameraControl.Temperature = clsAction.DesiredTemperatures(CurrentStep)
            MyTime = Now
            
            If Aborted Then Exit Sub
            
            Do
                Call Wait(10)
                Call AddToStatus("Current temperature = " & Format(Camera.objCameraControl.Temperature, "0.0"))
                ElapsedTime = CDbl(DateDiff("s", MyTime, Now)) / 60
            Loop Until ElapsedTime > clsAction.MaxTime Or (Abs(Camera.objCameraControl.Temperature - clsAction.DesiredTemperatures(CurrentStep)) < clsAction.Deviation) Or Aborted
            
            If Aborted Then Exit Sub
            
            If ElapsedTime > clsAction.MaxTime Then
                Call AddToStatus("Maximum time exceeded.  Current temperature = " & Format(Camera.objCameraControl.Temperature, "0.0"))
                If CurrentStep > 0 Then
                    'back off to the previous step
                    Call AddToStatus("Backing off, setting temperature = " & clsAction.DesiredTemperatures(CurrentStep - 1))
                    Camera.objCameraControl.Temperature = clsAction.DesiredTemperatures(CurrentStep - 1)
                    Exit Do
                Else
                    'cannot go any warmer, leave cooler where it is and fail
                    Call AddToStatus("Error.  Cannot achieve minimum acceptable temperature.  Quiting.")
                    Exit Sub
                End If
            Else
                'wait for current step to stabilize
                Call AddToStatus("Waiting for current temperature to stabilize.")
                
                Temp2 = Camera.objCameraControl.Temperature
                Cooler2 = Camera.objCameraControl.CoolerPower
                If Aborted Then Exit Sub
                
                Do
                    Temp1 = Temp2
                    Cooler1 = Cooler2
                    Call Wait(20)
                    Temp2 = Camera.objCameraControl.Temperature
                    Cooler2 = Camera.objCameraControl.CoolerPower
                    
                    Call AddToStatus("Desired temp = " & clsAction.DesiredTemperatures(CurrentStep) & ", Previous Temp = " & Format(Temp1, "0.0") & ", Current Temp = " & Format(Temp2, "0.0"))
                    Call AddToStatus("Previous Cooler Power = " & Cooler1 & "%, Current Cooler Power = " & Cooler2 & "%")
                    
                    ElapsedTime = CDbl(DateDiff("s", MyTime, Now)) / 60
                    
                Loop Until ElapsedTime > clsAction.MaxTime Or _
                    ((Abs(Temp1 - Temp2) < clsAction.Deviation) And _
                    (Abs(Temp1 - clsAction.DesiredTemperatures(CurrentStep)) < clsAction.Deviation) And _
                    (Abs(Temp2 - clsAction.DesiredTemperatures(CurrentStep)) < clsAction.Deviation) And _
                    (Abs(Cooler1 - Cooler2) < clsAction.CoolerDeviation)) Or _
                    (Camera.objCameraControl.CoolerPower < clsAction.MaxCoolerPower * 0.9) Or Aborted
                    
                If ElapsedTime > clsAction.MaxTime Then
                    Call AddToStatus("Time limit exceeded.")
                End If
                
                If Aborted Then Exit Sub
                
                If Camera.objCameraControl.CoolerPower > clsAction.MaxCoolerPower Then
                    Call AddToStatus("Cooler power is too high.  Current power = " & Camera.objCameraControl.CoolerPower & ", Max Power = " & clsAction.MaxCoolerPower)
                    
                    'Cooler to high, back off one step
                    If CurrentStep > 0 Then
                        Call AddToStatus("Backing off, setting temperature = " & clsAction.DesiredTemperatures(CurrentStep - 1))
                        Camera.objCameraControl.Temperature = clsAction.DesiredTemperatures(CurrentStep - 1)
                        Exit Do
                    Else
                        'cannot go any warmer, leave cooler where it is and fail
                        Call AddToStatus("Error.  Cannot achieve minimum acceptable temperature.  Quiting.")
                        Exit Sub
                    End If
                ElseIf Camera.objCameraControl.CoolerPower < clsAction.MaxCoolerPower * 0.9 Then
                    Call AddToStatus("Cooler power is low enough (" & Camera.objCameraControl.CoolerPower & "%), no need to wait any longer.")
                End If
            End If
            
            CurrentStep = CurrentStep + 1
        Loop
        
        If Aborted Then Exit Sub
        
        CurrentStep = CurrentStep - 1
        
        'wait again for current step to stabilize
        Call AddToStatus("Waiting for final temperature to stabilize.")
        Temp2 = Camera.objCameraControl.Temperature
        Cooler2 = Camera.objCameraControl.CoolerPower
        Do
            Temp1 = Temp2
            Cooler1 = Cooler2
            Call Wait(20)
            Temp2 = Camera.objCameraControl.Temperature
            Cooler2 = Camera.objCameraControl.CoolerPower
            
            Call AddToStatus("Desired temp = " & clsAction.DesiredTemperatures(CurrentStep) & ", Previous Temp = " & Format(Temp1, "0.0") & ", Current Temp = " & Format(Temp2, "0.0"))
            Call AddToStatus("Previous Cooler Power = " & Cooler1 & "%, Current Cooler Power = " & Cooler2 & "%")
            
        Loop Until ((Abs(Temp1 - Temp2) < clsAction.Deviation) And _
            (Abs(Temp1 - clsAction.DesiredTemperatures(CurrentStep)) < clsAction.Deviation) And _
            (Abs(Temp2 - clsAction.DesiredTemperatures(CurrentStep)) < clsAction.Deviation) And _
            (Abs(Cooler1 - Cooler2) < clsAction.CoolerDeviation)) Or Aborted
            
        'done!
        Call AddToStatus("Intelligent Temperature Control complete.")
            
        If frmMain.mnuViewItem(0).Checked = True Then
            frmTempGraph.Timer1.Enabled = True
        End If
    ElseIf clsAction.CoolerOn And clsAction.IntelligentAction = 0 Then
        'now force the cooler on and go to the next temp in the list
        Call AddToStatus("Turning on cooler.")
        Camera.objCameraControl.CoolerState = True
       
        Call AddToStatus("Setting cooler = " & clsAction.DesiredTemperatures(0))
        Camera.objCameraControl.Temperature = clsAction.DesiredTemperatures(0)
        
        Call AddToStatus("Simple Temperature Control complete.")
    Else
        'cooler off
        'Warm up the cooler in 10 degree steps until the cooler power is less than 15%
        If (clsAction.RampWarmUp) Then
            Call AddToStatus("Warming up camera...")
            Do
                Temp1 = Camera.objCameraControl.Temperature + 10
                
                If Temp1 > 25 Then Temp1 = 25
                
                Call AddToStatus("Setting camera temperature to " & Format(Temp1, "0.0"))
                Camera.objCameraControl.Temperature = Temp1
                Counter = 0
                Do
                    Temp2 = Camera.objCameraControl.Temperature
                    If Counter = 0 Then
                        Call AddToStatus("Waiting for temperature to rise...")
                    End If
                    Call Wait(10)
                    Counter = Counter + 1
                    If Counter = 6 Then Counter = 0
                Loop While Camera.objCameraControl.Temperature < Temp1 And Camera.objCameraControl.Temperature > Temp2 And Not Aborted
            Loop Until Camera.objCameraControl.CoolerPower < 15 Or Aborted
        End If
        
        Call AddToStatus("Shutting down cooler.")
        Camera.objCameraControl.CoolerState = False
    End If
End Sub
