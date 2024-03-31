VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPinPointPlateSolve"
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

Private WithEvents myPinPoint As PinPoint.Plate
Attribute myPinPoint.VB_VarHelpID = -1

Public Function PlateSolve(FileName As String, PixelScale As Double, NorthAngle As Double, RA As Double, Dec As Double) As Boolean
    Dim Retry As Boolean
    
    Retry = False
    Do
        Set myPinPoint = New PinPoint.Plate
        
        With myPinPoint
            Call .AttachFITS(FileName)
            .ArcsecPerPixelHoriz = PixelScale
            .ArcsecPerPixelVert = PixelScale
            .RightAscension = .TargetRightAscension
            .Declination = .TargetDeclination
            
            .Catalog = frmOptions.lstCatalog.ItemData(frmOptions.lstCatalog.ListIndex)
            .CatalogPath = frmOptions.txtCatalogPath.Text
            
            .CatalogMaximumMagnitude = Settings.CatalogMagMin
            .CatalogMinimumMagnitude = Settings.CatalogMagMax
            .CatalogExpansion = (Sqr(Settings.SearchArea / 100) - 1) / 2
            
            .MaxSolveStars = Settings.MaxNumStars
            .SigmaAboveMean = Settings.StandardDeviation
            
            .MaxSolveTime = Settings.PinPointLETimeout * 60
            
            Call AddToStatus("Performing Plate Solve with PinPoint...")
            On Error Resume Next
            Call .Solve
                    
            If Err.Number = 0 Then
                PlateSolve = True
                NorthAngle = .RollAngle
                RA = .RightAscension
                Dec = .Declination
                
                Call AddToStatus("Plate solve results:")
                Call AddToStatus("Pixel Scale = " & Format((.ArcsecPerPixelHoriz + .ArcsecPerPixelVert) / 2, "0.00") & " asp.")
                Call AddToStatus("Focal Length = " & Format(206.265 * Camera.PixelSize / Abs(.ArcsecPerPixelHoriz), "0.00") & " mm.")
                Call AddToStatus("North Angle = " & Format(NorthAngle, "0.00") & " degrees.")
                Call AddToStatus("J2000 Coordinates = " & Misc.ConvertEquatorialToString(RA, Dec, False))
                
                Retry = False
            Else
                Call AddToStatus("Plate Solve Error Number: " & Err.Number)
                Call AddToStatus(Err.Description)
                
                PlateSolve = False
                
                If Retry Then
                    Retry = False
                ElseIf Err.Number = -2147220413 Then
                    Call AddToStatus("Trying again...")
                    Call Wait(1)
                    Retry = True
                End If
            End If
        
            On Error GoTo 0
        End With
        
        Set myPinPoint = Nothing
    Loop Until Retry = False
End Function

Private Sub myPinPoint_Checkpoint(Continue As Boolean)
    DoEvents
    
    If Aborted Then
        Continue = False
    Else
        Continue = True
    End If
End Sub

