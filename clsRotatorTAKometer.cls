VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRotatorTAKometer"
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

Private myCOMPort As Integer
Private objTAKometer As Object

Public Property Get CurrentAngle() As Double
    CurrentAngle = objTAKometer.CurrentAngle
    
    Call AddToStatus("TAKometer reports current angle as " & CLng(CurrentAngle) & " degrees.")
    
    If (CurrentAngle < 0) Then
        CurrentAngle = 360 + CurrentAngle
        Call AddToStatus("Recomputed TAKometer angle to be " & CLng(CurrentAngle) & " degrees.")
    End If

End Property

Public Property Let CurrentAngle(Value As Double)
    Dim Counter As Integer
    
    Value = CLng(Misc.DoubleModulus(CInt(Value), 360))
    
    If Value > 180 Then
        Value = Value - 360
        Call AddToStatus("Recomputed TAKometer angle to be " & CLng(Value) & " degrees.")
    End If
        
    Call objTAKometer.MoveToNewAngle(CLng(Value), True)
    
    Counter = 0
    Do While objTAKometer.Moving And Counter < 240 And Not Aborted
        Call Wait(1)
        Counter = Counter + 1
    Loop

    If Aborted Then
        Call AddToStatus("Rotation aborted.")
        Call objTAKometer.Abort
    ElseIf Counter > 240 Then
        Call AddToStatus("TAKometer never completed rotation.")
    Else
        Call AddToStatus("Rotation complete.")
    End If
End Property

Public Property Let RotatorCOMPort(Value As Integer)
    myCOMPort = Value
End Property

Public Function ConnectToRotator() As Boolean
    Set objTAKometer = CreateObject("TAKometerControlLib.TAKometerControl")
    
    If Not objTAKometer.Connected Then
        objTAKometer.CommPort = myCOMPort
        ConnectToRotator = objTAKometer.Connect
        
        If Not ConnectToRotator Then
            'try new TAKometer object
            Set objTAKometer = CreateObject("TAKometerControlLib2.TAKometerControl")
            
            If Not objTAKometer.Connected Then
                objTAKometer.CommPort = myCOMPort
                ConnectToRotator = objTAKometer.Connect
            Else
                ConnectToRotator = True
            End If
        End If
    Else
        ConnectToRotator = True
    End If
    
    If ConnectToRotator Then
        Call AddToStatus("Connected to TAKometer rotator.")
    Else
        Call AddToStatus("Unable to connect to TAKometer.")
    End If
End Function

Private Sub Class_Initialize()
    myCOMPort = 1
End Sub

Private Sub Class_Terminate()
    Set objTAKometer = Nothing
End Sub
