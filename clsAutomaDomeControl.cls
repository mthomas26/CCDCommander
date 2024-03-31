VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutomaDomeControl"
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

Private objRASCOMDome As Object 'RASCOMLib.RASCOMDome

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
    myDomeOpen = True
End Sub

Public Sub DomeClose()
    objRASCOMDome.CloseSlit
    myDomeOpen = False
End Sub

Public Sub ConnectToDome()
    'Set objRASCOMDome = New RASCOMLib.RASCOMDome
    On Error Resume Next
    
    Set objRASCOMDome = CreateObject("AutomaDome.Dome")
    If Err.Number <> 0 Then
        On Error GoTo 0
        'Try the old RASCOM dome interface
        Set objRASCOMDome = CreateObject("RASCOM.RASCOMDome")
    Else
        On Error GoTo 0
    End If
    
    'objRASCOMDome.IsAsynchronous = False 'This function doesn't work
    objRASCOMDome.Connect
End Sub

Public Sub DomeCouple()
    Dim objTheSky As New TheSky6Library.TheSkyRASCOMTheSky
    
    objTheSky.ConnectDome
    objTheSky.CoupleDome
    
    myDomeCoupled = True
End Sub

Public Sub DomeUnCouple()
    Dim objTheSky As New TheSky6Library.TheSkyRASCOMTheSky
    
    objTheSky.DisconnectDome
    
    myDomeCoupled = False
End Sub

Public Sub DomeHome()
    objRASCOMDome.FindHome
End Sub

Public Sub DomePark()
    objRASCOMDome.Park
End Sub

Public Sub DisconnectFromDome()
    objRASCOMDome.DisconnectDome
    
    Set objRASCOMDome = Nothing
End Sub

Public Sub SlewToAzimuth(Azimuth As Integer)
    Dim Counter As Integer
    
    Call DomeUnCouple   'Cannot slew if coupled
    
    Call objRASCOMDome.GotoAzEl(CDbl(Azimuth), objRASCOMDome.dEl)
End Sub

