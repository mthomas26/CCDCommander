VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheSky6TelescopeControl"
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

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CURRENT_USER = &H80000001

Private objTele As TheSky6Library.TheSkyRASCOMTele
Private SkyUtil As TheSky6Library.Utils
Private TheSkyStarChart As TheSky6Library.StarChart

Private IsTracking As Boolean

Public Sub Abort()
    On Error Resume Next
    Call objTele.Abort
    On Error GoTo 0
End Sub

Public Property Get Alt() As Double
    TheSkyStarChart.UpdateSolarSystem
    Call objTele.GetAzAlt
    Alt = objTele.dAlt
End Property

Public Property Get Az() As Double
    TheSkyStarChart.UpdateSolarSystem
    Call objTele.GetAzAlt
    Az = objTele.dAz
End Property

Public Sub ConnectToMount()
    objTele.Connect
    
    Tracking = True
End Sub

Public Function CanSetTracking() As Boolean
    CanSetTracking = True
End Function

Public Property Get Dec() As Double
    Call objTele.GetRaDec
    Dec = objTele.dDec
End Property

Public Property Get Elevation() As Double
    Elevation = TheSkyStarChart.DocumentProperty(skDocProp_ElevationInMeters)
End Property

Public Property Get IsSlewComplete() As Boolean
    If objTele.IsSlewComplete <> 1 Then
        IsSlewComplete = False
    Else
        IsSlewComplete = True
    End If
End Property

Public Property Get Latitude() As Double
    Latitude = TheSkyStarChart.DocumentProperty(skDocProp_Latitude)
End Property

Public Property Get LastError() As Long
    LastError = objTele.GetLastSlewError
End Property

Public Property Get LocalSiderealTime() As Double
    TheSkyStarChart.UpdateSolarSystem
    LocalSiderealTime = SkyUtil.ComputeLocalSiderealTime
End Property

Public Property Get Longitude() As Double
    Longitude = TheSkyStarChart.DocumentProperty(skDocProp_Longitude)
End Property

Public Sub Park()
    Call objTele.Park
End Sub

Public Sub Home()
    Call objTele.FindHome
End Sub

Public Property Get RA() As Double
    Call objTele.GetRaDec
    RA = objTele.dRA
End Property

Public Property Get SideOfPier() As typMountSide
    SideOfPier = typMountSide.Unknown
End Property

Public Sub SlewToRADec(RA As Double, Dec As Double, ObjectName As String)
    Call objTele.SlewToRADec(RA, Dec, ObjectName)
End Sub

Public Sub SlewToAzAlt(Az As Double, Alt As Double, ObjectName As String)
    Call objTele.SlewToAzAlt(Az, Alt, ObjectName)
End Sub

Public Sub Sync(RA As Double, Dec As Double)
    Dim ErrNo As Long
    Dim ErrDesc As String
    Dim ErrSrce As String
    On Error Resume Next
    Call objTele.Sync(RA, Dec, "CCDCommanderPos")
    
    ErrNo = Err.Number
    ErrDesc = Err.Description
    ErrSrce = Err.Source
    
    On Error GoTo 0
    If ErrNo <> 0 Then
        If ErrNo = -2147220286 Then
            Call AddToStatus("Sync failed.  TheSky reports that the limits have been exceeded.")
        Else
            Err.Raise ErrNo, ErrSrce, ErrDesc
        End If
    End If
End Sub

Public Property Get TelescopeType() As String
    Dim KeyPtr As Long
    Dim Buff As String * 255
    Dim BufLen As Long
        
    Call RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Software Bisque\TheSky6\TELESCOPE", 0, 1, KeyPtr)
    
    BufLen = Len(Buff)
    Call RegQueryValueEx(KeyPtr, "Name", 0, 0, Buff, BufLen)
    
    Call RegCloseKey(KeyPtr)
    
    TelescopeType = Left(Buff, BufLen)
End Property

Public Property Let Tracking(Value As Boolean)
    IsTracking = Value
    If Value Then
        Call objTele.SetTracking(1, 1, 0, 0)
    Else
        Call objTele.SetTracking(0, 1, 0, 0)
    End If
End Property

Public Property Get Tracking() As Boolean
    Tracking = IsTracking
End Property

Private Sub Class_Initialize()
    Set objTele = New TheSky6Library.TheSkyRASCOMTele
    Set SkyUtil = New TheSky6Library.Utils
    Set TheSkyStarChart = New TheSky6Library.StarChart
    
    objTele.Asynchronous = True
End Sub

Private Sub Class_Terminate()
    Set objTele = Nothing
    Set SkyUtil = Nothing
    Set TheSkyStarChart = Nothing
End Sub
