VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCCDSoftPlateSolve"
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
    Dim objImage As CCDSoftLib.Image
    Dim RADecArray() As Variant
    Dim myErrorNum As Long
    
    On Error Resume Next
    'This will remove any soultion files
    Call Kill(App.Path & "\*.SRC")
    On Error GoTo 0
    
    Set objImage = CreateObject("CCDSoft.Image")
    
    With objImage
        .Path = FileName
        Call .Open
    
        'Image Link
        .ScaleInArcsecondsPerPixel = PixelScale
        Call AddToStatus("Performing Plate Solve with CCDSoft...")
    
        On Error Resume Next
        Call .InsertWCS(True)
        myErrorNum = Err.Number
        On Error GoTo 0
    
        If myErrorNum = 0 Then
            NorthAngle = .NorthAngle
            RADecArray = .XYToRADec(.Width / 2, .Height / 2)
            RA = RADecArray(0)
            Dec = RADecArray(1)
            
            Call AddToStatus("Plate solve results:")
            
            On Error Resume Next
            Call AddToStatus("Pixel Scale = " & Format(.ScaleInArcsecondsPerPixel, "0.00") & " asp.")
            Call AddToStatus("Focal Length = " & Format(206.265 * CDbl(.FITSKeyword("XPIXSZ")) / .ScaleInArcsecondsPerPixel, "0.00") & " mm.")
            Call AddToStatus("North Angle = " & Format(NorthAngle, "0.00") & " degrees.")
            Call AddToStatus("J2000 Coordinates = " & Misc.ConvertEquatorialToString(RA, Dec, False))
            On Error GoTo 0
            
            PlateSolve = True
        Else
            PlateSolve = False
        End If
        
        .Close
    End With
End Function


