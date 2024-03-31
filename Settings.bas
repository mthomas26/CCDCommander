Attribute VB_Name = "Settings"
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

'Right now this just contains values for the text boxes on the Options & Settings window.
'I should probably include everything here...

Public RotatorCOMNumber As Long
Public GuiderCalibrationAngle As Double
Public HomeRotationAngle As Double

Public MaxPointingError As Double

Public NorthAngle As Double
Public PinPointLETimeout As Double
Public PixelScale As Double

Public CloudMonitorQueryPeriod As Double
Public CloudMonitorClearTime As Double

Public DelayAfterSlew As Double
Public EasternLimit As Double
Public WesternLimit As Double

Public GuideBoxX As Long
Public GuideBoxY As Long
Public GuideStarFWHM As Integer
Public MaxStarMovement As Long
Public MinimumGuideStarADU As Long
Public MaximumGuideStarADU As Long
Public MinimumGuideStarExposure As Double
Public GuideStarExposureIncrement As Double
Public MaximumGuideStarExposure As Double
Public GuiderRestartError As Integer
Public GuiderRestartCycles As Integer

Public CatalogMagMax As Double
Public CatalogMagMin As Double
Public SearchArea As Double
Public MinStarBrightness As Long
Public MaxNumStars As Integer
Public StandardDeviation As Double

Public SMTPPort As Long

Public RetryFocusCount As Integer
Public FocusTimeOut As Integer
Public FocusMaxDisconnectReconnect As Boolean

Public DDWTimeout As Integer
Public DDWRetryCount As Integer

Public MaximumStarFadedErrors As Integer

Public WeatherMonitorRestartActionList As Boolean
