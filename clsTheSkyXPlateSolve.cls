VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTheSkyXPlateSolve"
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

Public Function PlateSolve(FileName As String, PixelScale As Double, NorthAngle As Double, RA As Double, Dec As Double) As Boolean
    Dim ImageLink As Object 'theskyxLib.ImageLink
    Dim ImageLinkResults As Object 'theskyxLib.ImageLinkResults
    Dim RADecArray() As Variant
    Dim myErrorNum As Long
    
    Set ImageLink = CreateObject("theskyx.ImageLink")
    Set ImageLinkResults = CreateObject("theskyx.ImageLinkResults")
    
    With ImageLink
        .pathToFITS = FileName
    
        'Image Link
        .Scale = PixelScale
        Call AddToStatus("Performing Plate Solve with TheSkyX...")
    
        On Error Resume Next
        Call .execute
        myErrorNum = Err.Number
        On Error GoTo 0
    End With
    
    With ImageLinkResults
        If myErrorNum = 0 And .succeeded = 1 Then
            NorthAngle = .imagePositionAngle
            
            RA = .imageCenterRAJ2000
            Dec = .imageCenterDecJ2000
            
            Call AddToStatus("Plate solve results:")
            
            On Error Resume Next
            Call AddToStatus("Pixel Scale = " & Format(.ImageScale, "0.00") & " asp.")
            Call AddToStatus("Focal Length = " & Format(206.265 * CDbl(.FITSKeyword("XPIXSZ")) / .ImageScale, "0.00") & " mm.")
            Call AddToStatus("North Angle = " & Format(NorthAngle, "0.00") & " degrees.")
            Call AddToStatus("J2000 Coordinates = " & Misc.ConvertEquatorialToString(RA, Dec, False))
            On Error GoTo 0
            
            PlateSolve = True
        Else
            PlateSolve = False
        End If
    End With
    
    Set ImageLinkResults = Nothing
    Set ImageLink = Nothing
End Function



