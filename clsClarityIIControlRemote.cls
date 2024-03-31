VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClarityIIControlRemote"
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

Private LastSensorTime As Date
Private LastReadTime As Date
Private myCloudStatus As CloudConditions
Private myWindStatus As WindConditions
Private myRainStatus As RainConditions
Private myLightStatus As DayLightConditions
Private NoUpdateCounter As Integer

Public Sub ConnectToWeatherMonitor()
    'Confirm I can read the file - get the most recent data
    If Not ReadCloudSensorFile() Then
        'problem - toss an error
        Call Err.Raise(1001, "Clarity II Control Remote", "Unable to get Clarity data file.")
    End If
End Sub

Private Function ReadCloudSensorFile() As Boolean
    Dim FileNo As Integer
    Dim tempTime As Date
    Dim ClarityData As String
    Dim SensorData As Integer
    
    LastReadTime = Now
    
    'Read in the file
    FileNo = FreeFile()
    On Error Resume Next
    tempTime = Now
    Do
        Err.Clear
        Open frmOptions.txtWeatherMonitorRemoteFile.Text For Input Lock Read Write As #FileNo
    Loop While Err.Number = 70 And DateDiff("s", tempTime, Now) < 10
    If Err.Number <> 0 Then
        Err.Clear
        Close FileNo
        
        GoTo ReadCloudSensorError
    End If
    
    Line Input #FileNo, ClarityData
    Close FileNo
    
    On Error GoTo ReadCloudSensorError
    tempTime = DateSerial(CInt(Left(ClarityData, 4)), CInt(Mid(ClarityData, 6, 2)), CInt(Mid(ClarityData, 9, 2))) + TimeSerial(CInt(Mid(ClarityData, 12, 2)), CInt(Mid(ClarityData, 15, 2)), CDbl(Mid(ClarityData, 18, 5)))
    If DateDiff("s", tempTime, LastSensorTime) = 0 Then
        'uh oh, no update - something is wrong
        NoUpdateCounter = NoUpdateCounter + 1
        Call AddToStatus("Clarity data file unchanged since last read! Count = " & NoUpdateCounter)
        If NoUpdateCounter >= 5 Then
            GoTo ReadCloudSensorError
        End If
    Else
        NoUpdateCounter = 0
    End If
    LastSensorTime = tempTime
    
    'get Cloud Condition
    SensorData = CInt(Mid(ClarityData, 94, 1))
    If SensorData = 0 Then
        myCloudStatus = CloudConditions.Unknown
    ElseIf SensorData = 1 Then
        myCloudStatus = Clear
    ElseIf SensorData = 2 Then
        myCloudStatus = Cloudy
    ElseIf SensorData = 3 Then
        myCloudStatus = VeryCloudy
    End If
    
    'get wind condition
    SensorData = CInt(Mid(ClarityData, 96, 1))
    If SensorData = 0 Then
        myWindStatus = WindConditions.Unknown
    ElseIf SensorData = 1 Then
        myWindStatus = Calm
    ElseIf SensorData = 2 Then
        myWindStatus = Windy
    ElseIf SensorData = 3 Then
        myWindStatus = VeryWindy
    End If
    
    'get rain condition
    SensorData = CInt(Mid(ClarityData, 98, 1))
    If SensorData = 0 Then
        myRainStatus = RainConditions.Unknown
    ElseIf SensorData = 1 Then
        myRainStatus = Dry
    ElseIf SensorData = 2 Then
        myRainStatus = Wet
    ElseIf SensorData = 3 Then
        myRainStatus = Rain
    End If
    
    'get light condition
    SensorData = CInt(Mid(ClarityData, 100, 1))
    If SensorData = 0 Then
        myLightStatus = DayLightConditions.Unknown
    ElseIf SensorData = 1 Then
        myLightStatus = Dark
    ElseIf SensorData = 2 Then
        myLightStatus = Light
    ElseIf SensorData = 3 Then
        myLightStatus = VeryLight
    End If
    
    ReadCloudSensorFile = True
    
    On Error GoTo 0
    Exit Function
    
ReadCloudSensorError:
    On Error GoTo 0
    
    Call AddToStatus("Error getting weather monitor status.  Setting all status to Unknown.")
    
    myCloudStatus = CloudConditions.Unknown
    myWindStatus = WindConditions.Unknown
    myRainStatus = RainConditions.Unknown
    myLightStatus = DayLightConditions.Unknown
    
    ReadCloudSensorFile = False
End Function

Public Property Get CloudStatus() As CloudConditions
    If DateDiff("s", LastReadTime, Now) > Settings.CloudMonitorQueryPeriod * 0.9 Then
        Call ReadCloudSensorFile
    End If
    
    CloudStatus = myCloudStatus
End Property

Public Property Get WindStatus() As WindConditions
    If DateDiff("s", LastReadTime, Now) > Settings.CloudMonitorQueryPeriod * 0.9 Then
        Call ReadCloudSensorFile
    End If
    
    WindStatus = myWindStatus
End Property

Public Property Get RainStatus() As RainConditions
    If DateDiff("s", LastReadTime, Now) > Settings.CloudMonitorQueryPeriod * 0.9 Then
        Call ReadCloudSensorFile
    End If
    
    RainStatus = myRainStatus
End Property

Public Property Get LightStatus() As DayLightConditions
    If DateDiff("s", LastReadTime, Now) > Settings.CloudMonitorQueryPeriod * 0.9 Then
        Call ReadCloudSensorFile
    End If
    
    LightStatus = myLightStatus
End Property

