VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClairtyIControl"
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
    Set objCloud = CreateObject("Clarity.CloudSensor")
End Sub

Public Property Get CloudStatus() As CloudConditions
    Dim mySensorStatus As Integer
    
    mySensorStatus = objCloud.SkyCondition

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
    WindStatus = WindConditions.NotImplemented
End Property

Public Property Get RainStatus() As RainConditions
    Dim mySensorStatus As Integer
    
    mySensorStatus = objCloud.SkyCondition
    
    If mySensorStatus = 4 Then
        RainStatus = Wet
    Else
        RainStatus = Dry
    End If
End Property

Public Property Get LightStatus() As DayLightConditions
    LightStatus = DayLightConditions.NotImplemented
End Property
