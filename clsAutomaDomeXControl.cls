VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutomaDomeXControl"
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

Private objRASCOMDome As Object 'theskyxLib.sky6Dome

Private myDomeOpen As Boolean
Private myDomeCoupled As Boolean

Public Function IsDomeOpen(Optional Force As Boolean = False, Optional State As Boolean = False) As Boolean
    If Force Then
        myDomeOpen = State
    End If
    
    IsDomeOpen = myDomeOpen
End Function

Public Property Get IsDomeCoupled() As Boolean
    IsDomeCoupled = myDomeCoupled
End Property

Public Sub DomeOpen()
    objRASCOMDome.OpenSlit
    Do
        Call Wait(1)
    Loop While (objRASCOMDome.IsOpenComplete = 0) And Not Aborted
    
    myDomeOpen = True
End Sub

Public Sub DomeClose()
    objRASCOMDome.CloseSlit
    Do
        Call Wait(1)
    Loop While (objRASCOMDome.IsCloseComplete = 0) And Not Aborted
    
    myDomeOpen = False
End Sub

Public Sub ConnectToDome()
    'Set objRASCOMDome = New RASCOMLib.RASCOMDome
    On Error Resume Next
    
    Set objRASCOMDome = CreateObject("TheSkyX.sky6Dome")
    
    'Try the old AutomaDome interface
'    If Err.Number <> 0 Then
'        Set objRASCOMDome = CreateObject("AutomaDome.Dome")
'    End If
'
'    If Err.Number <> 0 Then
'        On Error GoTo 0
'        'Try the old RASCOM dome interface
'        Set objRASCOMDome = CreateObject("RASCOM.RASCOMDome")
'    Else
'        On Error GoTo 0
'    End If
    
    'objRASCOMDome.IsAsynchronous = False 'This function doesn't work
    objRASCOMDome.Connect
End Sub

Public Sub DomeCouple()
    'Dim objTheSky As New TheSky6Library.TheSkyRASCOMTheSky
    Dim objTheSkyX As Object 'New theskyxLib.sky6RASCOMTheSky
    Set objTheSkyX = CreateObject("theskyx.sky6RASCOMTheSky")
    
    objTheSkyX.ConnectDome
    objTheSkyX.CoupleDome
    
    myDomeCoupled = True
End Sub

Public Sub DomeUnCouple()
    'Dim objTheSky As New TheSky6Library.TheSkyRASCOMTheSky
    Dim objTheSkyX As Object 'New theskyxLib.sky6RASCOMTheSky
    Set objTheSkyX = CreateObject("theskyx.sky6RASCOMTheSky")
    
    objTheSkyX.DisconnectDome
    
    myDomeCoupled = False
End Sub

Public Sub DomeHome()
    objRASCOMDome.FindHome
    Do
        Call Wait(1)
    Loop While (objRASCOMDome.IsFindHomeComplete = 0) And Not Aborted
End Sub

Public Sub DomePark()
    objRASCOMDome.Park
    Do
        Call Wait(1)
    Loop While (objRASCOMDome.IsParkComplete = 0) And Not Aborted
End Sub

Public Sub DisconnectFromDome()
    objRASCOMDome.Disconnect
    
    Set objRASCOMDome = Nothing
End Sub

Public Sub SlewToAzimuth(Azimuth As Integer)
    Dim Counter As Integer
    
    Call DomeUnCouple   'Cannot slew if coupled
    
    'The above disconnects the dome, not just from the mount.
    'Make sure to reconnect just the dome
    objRASCOMDome.Connect
    
    Call objRASCOMDome.GotoAzEl(CDbl(Azimuth), objRASCOMDome.dEl)
    
    Do
        Call Wait(1)
    Loop While (objRASCOMDome.IsGotoComplete = 0) And Not Aborted
End Sub


