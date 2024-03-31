VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRotatorSim"
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

Private myCurrentAngle As Double

Public Property Get CurrentAngle() As Double
    CurrentAngle = myCurrentAngle
    
    Call AddToStatus("Current angle is " & CLng(CurrentAngle) & " degrees.")
End Property

Public Property Let CurrentAngle(Value As Double)
    Dim Counter As Integer
    
    myCurrentAngle = Value
    
    Call MsgBox("Rotate to " & Format(Value, "0.00") & " degrees.  Click Ok when done.", vbOKOnly)

    Call AddToStatus("Rotation complete.")
End Property

Public Property Let RotatorCOMPort(Value As Integer)
    'Not used
End Property

Public Function ConnectToRotator() As Boolean
    Call AddToStatus("Connected to Rotator.")
    ConnectToRotator = True
    
    myCurrentAngle = CDbl(InputBox("Enter current rotation.", "Manual Rotator", "0"))
End Function

