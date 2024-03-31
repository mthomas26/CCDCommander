VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCOMTelescopeControl"
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

Private objTele As Object

Public Sub Abort()
    On Error Resume Next
    Call objTele.AbortSlew
    On Error GoTo 0
End Sub

Public Property Get Alt() As Double
    Alt = objTele.Altitude
End Property

Public Property Get Az() As Double
    Az = objTele.Azimuth
End Property

Public Sub ConnectToMount()
    Dim InterfaceVersion As Integer
    objTele.Connected = True
    
    'check if this is a V1 driver
    On Error Resume Next
    InterfaceVersion = objTele.InterfaceVersion
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "This driver is not supported." & vbCrLf & "Only ASCOM Telescope Drivers v2 and greater are supported.", vbCritical
        Aborted = True
        Exit Sub
    End If
    On Error GoTo 0
    
    ' check that the driver implementes the proper functions
    If Not objTele.CanSlewAsync Then
        MsgBox "This telescope driver is not supported." & vbCrLf & "It must implement the SlewAsync function.", vbCritical
        Aborted = True
        Exit Sub
    End If
    
    'Unpark telescope if necessary
    If objTele.AtPark Then
        If objTele.CanUnpark Then
            objTele.Unpark
        Else
            MsgBox "Mount is parked, but driver does not support UnPark." & vbCrLf & "Please unpark the mount and restart.", vbCritical
            Aborted = True
            Exit Sub
        End If
    End If
    
    'Confirm the tracking is enabled
    If objTele.CanSetTracking Then
        objTele.Tracking = True
    End If
End Sub

Public Function CanSetTracking() As Boolean
    CanSetTracking = objTele.CanSetTracking
End Function

Public Property Get Dec() As Double
    Dec = objTele.Declination
End Property

Public Property Get Elevation() As Double
    On Error Resume Next
    Elevation = objTele.SiteElevation
    If Err.Number <> 0 Then Elevation = 0
    On Error GoTo 0
End Property

Public Property Get IsSlewComplete() As Boolean
    IsSlewComplete = Not objTele.Slewing
End Property

Public Property Get Latitude() As Double
    Latitude = objTele.SiteLatitude
End Property

Public Property Get LastError() As Long
    LastError = 0
End Property

Public Property Get LocalSiderealTime() As Double
    LocalSiderealTime = objTele.SiderealTime
End Property

Public Property Get Longitude() As Double
    Longitude = -objTele.SiteLongitude
End Property

Public Sub Park()
    If objTele.CanPark Then
        On Error Resume Next    'just in case FindHome is not supported
        Call objTele.FindHome
        On Error GoTo 0
        Call objTele.Park
    Else
        Call AddToStatus("Driver does not support the Park function.  Cannot park the mount!")
    End If
End Sub

Public Sub Home()
    On Error Resume Next    'just in case FindHome is not supported
    Call objTele.FindHome
    On Error GoTo 0
End Sub

Public Property Get RA() As Double
    RA = objTele.RightAscension
End Property
 
Public Property Get SideOfPier() As typMountSide
    On Error GoTo SideOfPierError
    If objTele.SideOfPier = 0 Then
        SideOfPier = typMountSide.EastSide
    ElseIf objTele.SideOfPier = 1 Then
        SideOfPier = typMountSide.WestSide
    End If
    On Error GoTo 0
    
    Exit Property

SideOfPierError:
    On Error GoTo 0
    
    'Property must not be impelented
    SideOfPier = typMountSide.Unknown
End Property
 
Public Sub SlewToRADec(RA As Double, Dec As Double, ObjectName As String)
    Call objTele.SlewToCoordinatesAsync(RA, Dec)
End Sub

Public Sub SlewToAzAlt(Az As Double, Alt As Double, ObjectName As String)
    objTele.Tracking = False
    Call objTele.SlewToAltAzAsync(Az, Alt)
End Sub

Public Sub Sync(RA As Double, Dec As Double)
    If objTele.CanSync Then
        Call objTele.SyncToCoordinates(RA, Dec)
    Else
        Call AddToStatus("Driver does not support the Sync function.  Cannot sync the mount!")
    End If
End Sub

Public Property Get TelescopeType() As String
    TelescopeType = objTele.Name
End Property

Public Property Let Tracking(Value As Boolean)
    objTele.Tracking = Value
End Property

Public Property Get Tracking() As Boolean
    Tracking = objTele.Tracking
End Property

Private Sub Class_Initialize()
    Set objTele = CreateObject(GetMySetting("ProgramSettings", "ASCOMScopeProgID", "ScopeSim.Telescope"))
End Sub

Private Sub Class_Terminate()
    Set objTele = Nothing
End Sub
