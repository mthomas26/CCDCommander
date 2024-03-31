VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRotatorPyxis"
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

Private myCOMPort As Integer
Private myConnected As Boolean

Public Property Get CurrentAngle() As Double
    Dim Buffer As String
    Dim Counter As Integer
    
    frmMain.MSComm.Output = "CGETPA"
    
    Counter = 0
    Do While frmMain.MSComm.InBufferCount < 5 And Counter <= 50
        Call Wait(1)
        Counter = Counter + 1
    Loop

    If Counter > 50 And frmMain.MSComm.InBufferCount = 0 Then
        Call AddToStatus("Unable to get Pyxis current position angle - no response.")
        Exit Sub
    ElseIf Counter > 50 Then
        Buffer = ""
        Buffer = frmMain.MSComm.Input
        
        Call AddToStatus("Unable to get Pyxis current position angle - incorrect response.")
        Exit Sub
    Else
        Buffer = ""
        Buffer = frmMain.MSComm.Input
    End If

    Call AddToStatus("Pyxis reports current angle as " & CInt(Left(Buffer, 3)) & " degrees.")
        
    CurrentAngle = CInt(Left(Buffer, 3))
End Property

Public Property Let CurrentAngle(Value As Double)
    Dim Buffer As String
    Dim Counter As Integer
    
    frmMain.MSComm.Output = "CPA" & Format(Misc.DoubleModulus(Int(Value + 0.5), 360), "000")
    
    Do
        Call Wait(1)
        
        If frmMain.MSComm.InBufferCount >= 1 Then
            Buffer = frmMain.MSComm.Input
        End If
        
        Counter = Counter + 1
    Loop Until InStr(Buffer, "F") > 0 Or Counter > 600 Or Aborted
    
    If Aborted Then
        Call AddToStatus("Rotation aborted.")
    ElseIf Counter > 600 Then
        Call AddToStatus("Pyxis never completed rotation.")
    Else
        Call AddToStatus("Rotation complete.")
    End If
End Property

Public Property Let RotatorCOMPort(Value As Integer)
    myCOMPort = Value
End Property

Public Function ConnectToRotator() As Boolean
    Dim Buffer As String
    Dim Counter As Integer
    
    frmMain.MSComm.CommPort = myCOMPort
    
    If Not myConnected Then
        On Error Resume Next
        frmMain.MSComm.PortOpen = True
        If Err.Number <> 0 Then
            Call AddToStatus("Unable to open COM" & Settings.RotatorCOMNumber & ".")
            myConnected = False
        Else
            myConnected = True
        End If
        On Error GoTo 0
    End If
    
    If myConnected Then
        frmMain.MSComm.Output = "CCLINK"
        
        Counter = 0
        Do While frmMain.MSComm.InBufferCount < 3 And Counter <= 50
            Call Wait(1)
            Counter = Counter + 1
        Loop
        
        If Counter > 50 And frmMain.MSComm.InBufferCount = 0 Then
            Call AddToStatus("Unable to comunicate with Pyxis - no responce.")
            ConnectToRotator = False
        ElseIf Counter > 50 Then
            Buffer = ""
            Buffer = frmMain.MSComm.Input
            
            Call AddToStatus("Unable to comunicate with Pyxis - incorrect responce.")
            ConnectToRotator = False
        Else
            Buffer = ""
            Buffer = frmMain.MSComm.Input
            
            Call AddToStatus("Connected to Pyxis rotator.")
            ConnectToRotator = True
        End If
    Else
        ConnectToRotator = False
    End If
End Function

Private Sub Class_Initialize()
    myConnected = False
    myCOMPort = 1
End Sub

Private Sub Class_Terminate()
    If myConnected Then
        frmMain.MSComm.PortOpen = False
    End If
End Sub
