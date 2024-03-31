Attribute VB_Name = "Planetarium"
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

Public Enum PlanetariumControl
    None = 0
    TheSky6 = 1
    TheSkyX = 2
End Enum

Private myPlanetarium As Object

Public Sub ConnectToPlanetarium()
    If frmOptions.lstPlanetarium.ListIndex = PlanetariumControl.TheSky6 Then
        If TypeName(myPlanetarium) <> "TheSky6Planetarium" Then
            Set myPlanetarium = New TheSky6Planetarium
        End If
    ElseIf frmOptions.lstPlanetarium.ListIndex = PlanetariumControl.TheSkyX Then
        If TypeName(myPlanetarium) <> "TheSkyXPlanetarium" Then
            Set myPlanetarium = New TheSkyXPlanetarium
        End If
    Else
        Set myPlanetarium = Nothing
    End If
End Sub

Public Function GetFOVIPositionAngle() As Double
    Call ConnectToPlanetarium
    
    If Not (myPlanetarium Is Nothing) Then
        GetFOVIPositionAngle = myPlanetarium.GetFOVIPositionAngle
    End If
End Function

Public Sub SetFOVIPositionAngle(Value As Double)
    Call ConnectToPlanetarium
    
    If Not (myPlanetarium Is Nothing) Then
        Call myPlanetarium.SetFOVIPositionAngle(Value)
    End If
End Sub

Public Sub GetObjectRADec(ObjectName As String, RA As Double, Dec As Double, Optional J2000 As Boolean = False)
    Call ConnectToPlanetarium
    
    If Not (myPlanetarium Is Nothing) Then
        Call myPlanetarium.GetObjectRADec(ObjectName, RA, Dec, J2000)
    End If
End Sub

Public Function GetSelectedObject(RA As Double, Dec As Double, Optional J2000 As Boolean = False) As String
    Call ConnectToPlanetarium
    
    If Not (myPlanetarium Is Nothing) Then
        GetSelectedObject = myPlanetarium.GetSelectedObject(RA, Dec, J2000)
    End If
End Function
