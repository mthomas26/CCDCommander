VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClarityIIControl"
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

Private objCloud As Object

Public Sub ConnectToWeatherMonitor()
    Set objCloud = CreateObject("ClarityII.CloudSensorII")
End Sub

Public Property Get CloudStatus() As CloudConditions
    Dim mySensorStatus As Integer
    
    mySensorStatus = objCloud.CloudCondition

    If mySensorStatus = 0 Then
        CloudStatus = CloudConditions.Unknown
    ElseIf mySensorStatus = 1 Then
        CloudStatus = Clear
    ElseIf mySensorStatus = 2 Then
        CloudStatus = Cloudy
    ElseIf mySensorStatus = 3 Then
        CloudStatus = VeryCloudy
    End If
End Property

Public Property Get WindStatus() As WindConditions
    Dim mySensorStatus As Integer
    
    mySensorStatus = objCloud.WindCondition

    If mySensorStatus = 0 Then
        WindStatus = WindConditions.Unknown
    ElseIf mySensorStatus = 1 Then
        WindStatus = Calm
    ElseIf mySensorStatus = 2 Then
        WindStatus = Windy
    ElseIf mySensorStatus = 3 Then
        WindStatus = VeryWindy
    End If
End Property

Public Property Get RainStatus() As RainConditions
    Dim mySensorStatus As Integer
    
    mySensorStatus = objCloud.RainCondition
    
    If mySensorStatus = 0 Then
        RainStatus = RainConditions.Unknown
    ElseIf mySensorStatus = 1 Then
        RainStatus = Dry
    ElseIf mySensorStatus = 2 Then
        RainStatus = Wet
    ElseIf mySensorStatus = 3 Then
        RainStatus = Rain
    End If
End Property

Public Property Get LightStatus() As DayLightConditions
    Dim mySensorStatus As Integer
    
    mySensorStatus = objCloud.DayCondition
    
    If mySensorStatus = 0 Then
        LightStatus = DayLightConditions.Unknown
    ElseIf mySensorStatus = 1 Then
        LightStatus = Dark
    ElseIf mySensorStatus = 2 Then
        LightStatus = Light
    ElseIf mySensorStatus = 3 Then
        LightStatus = VeryLight
    End If
End Property

