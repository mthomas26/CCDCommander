VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASCOMDomeControl"
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

Private objDome As Object

Public Function IsDomeOpen(Optional Force As Boolean = False, Optional State As Boolean = False) As Boolean
    If objDome.ShutterStatus = 0 Or objDome.ShutterStatus = 2 Then
        IsDomeOpen = True
    Else
        IsDomeOpen = False
    End If
End Function

Public Property Get IsDomeCoupled() As Boolean
    IsDomeCoupled = objDome.Slaved
End Property

Public Sub DomeOpen()
    objDome.OpenShutter
    Call Wait(1)
    
    Do While objDome.ShutterStatus <> 0 And Not Aborted
        Call Wait(0.1)
        
        If objDome.ShutterStatus = 4 Then
            Call Err.Raise(65534, "ASCOM Dome Shutter Error", "Shutter Status reports Error")
        End If
    Loop
End Sub

Public Sub DomeClose()
    objDome.CloseShutter
    Call Wait(1)

    Do While objDome.ShutterStatus <> 1 And Not Aborted
        Call Wait(0.1)
    
        If objDome.ShutterStatus = 4 Then
            Call Err.Raise(65534, "ASCOM Dome Shutter Error", "Shutter Status reports Error")
        End If
    Loop
End Sub

Public Sub ConnectToDome()
    Set objDome = CreateObject(GetMySetting("ProgramSettings", "ASCOMDomeProgID", "DomeSim.Dome"))
    objDome.Connected = True
End Sub

Public Sub DomeCouple()
    Dim Counter As Integer
    
    If objDome.CanSlave Then
        objDome.Slaved = True

        Counter = 0
        Do While Not objDome.Slewing And Not Aborted And Counter < 25
            Call Wait(0.1)
            Counter = Counter + 1
        Loop
    
        Do While objDome.Slewing And Not Aborted
            Call Wait(0.1)
        Loop
    End If
End Sub

Public Sub DomeUnCouple()
    If objDome.CanSlave Then
        objDome.Slaved = False
    
        Do While objDome.Slewing And Not Aborted
            Call Wait(0.1)
        Loop

    End If
End Sub

Public Sub DomeHome()
    objDome.FindHome

    Do While objDome.Slewing And Not Aborted
        Call Wait(0.1)
    Loop
End Sub

Public Sub DomePark()
    objDome.Park

    Do While objDome.Slewing And Not Aborted
        Call Wait(0.1)
    Loop
End Sub

Public Sub DisconnectFromDome()
    objDome.Connected = False
    
    Set objDome = Nothing
End Sub

Public Sub SlewToAzimuth(Azimuth As Integer)
    Dim Counter As Integer
    
    Call DomeUnCouple   'Cannot slew if coupled
    
    Call objDome.SlewToAzimuth(Azimuth)

    Counter = 0
    Do While Not objDome.Slewing And Not Aborted And Counter < 25
        Call Wait(0.1)
        Counter = Counter + 1
    Loop

    Do While objDome.Slewing And Not Aborted
        Call Wait(0.1)
    Loop
End Sub
