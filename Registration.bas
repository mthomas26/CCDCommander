Attribute VB_Name = "Registration"
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

Public Const ReleaseYear = 2024
Private Const ReleaseMonth = 1
Private Const ReleaseDay = 1

Public Registered As Boolean
Public TrialDaysLeft As Integer
Public ValidForUpgrades As Boolean

Public RegName As String
Public RegAddress As String
Public RegCSZ As String
Public RegEMail As String
Public RegExperation As String
Public RegCode As String

Public Function CheckTrialAndRegistration() As Boolean
    'registered!
    CheckTrialAndRegistration = True
    TrialDaysLeft = 365
    ValidForUpgrades = True
    
    RegName = "Open Source Version"
    RegAddress = ""
    RegCSZ = ""
    RegEMail = ""
    RegExperation = ""
    RegCode = ""
End Function

