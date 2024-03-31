VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRotatorASCOM"
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

Dim objRotator As Object

Public Property Get CurrentAngle() As Double
    CurrentAngle = objRotator.Position
    
    Call AddToStatus("The ASCOM Rotator reports current angle as " & CLng(CurrentAngle) & " degrees.")
End Property

Public Property Let CurrentAngle(Value As Double)
    Dim Counter As Integer
    
    Call objRotator.MoveAbsolute(Value)
    
    Call Wait(5)
    
    Counter = 0
    Do While objRotator.IsMoving And Counter < 240 And Not Aborted
        Call Wait(1)
        Counter = Counter + 1
    Loop

    If Aborted Then
        Call AddToStatus("Rotation aborted.")
        Call objRotator.Halt
    ElseIf Counter > 240 Then
        Call AddToStatus("The ASCOM Rotator never completed rotation.")
    Else
        Call AddToStatus("Rotation complete.")
    End If
    
End Property

Public Property Let RotatorCOMPort(Value As Integer)
    'Not used
End Property

Public Function ConnectToRotator() As Boolean
    Set objRotator = CreateObject(GetMySetting("ProgramSettings", "ASCOMRotatorProgID", ""))

    objRotator.Connected = True
    
    If objRotator.Connected Then
        Call AddToStatus("Connected to the ASCOM Rotator.")
        ConnectToRotator = True
    Else
        Call AddToStatus("Unable to connect to the ASCOM Rotator.")
        ConnectToRotator = False
    End If
End Function

Private Sub Class_Terminate()
    Set objRotator = Nothing
End Sub

