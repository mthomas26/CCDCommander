VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParkMountAction"
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

Public RunTimeStatus As String  'This is for run time use only, no need to save it!

Dim DummyVar As Byte

Public DoSimulatedPark As Boolean
Public DoHomePark As Boolean
Public DoTrackingOff As Boolean

Public AltD As Double
Public AltM As Double
Public AltS As Double
Public AzimD As Double
Public AzimM As Double
Public AzimS As Double

Public Selected As Boolean

Public ParkRotator As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 6
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(DummyVar)
    MySize = MySize + Len(DoSimulatedPark)
    MySize = MySize + Len(DoHomePark)
    MySize = MySize + Len(DoTrackingOff)
    MySize = MySize + Len(AltD)
    MySize = MySize + Len(AltM)
    MySize = MySize + Len(AltS)
    MySize = MySize + Len(AzimD)
    MySize = MySize + Len(AzimM)
    MySize = MySize + Len(AzimS)
    
    MySize = MySize + Len(Selected)
    
    MySize = MySize + Len(ParkRotator)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), DummyVar, Len(DummyVar))
    Index = Index + Len(DummyVar)
    Call CopyMemory(ptrByteArray(Index), DoSimulatedPark, Len(DoSimulatedPark))
    Index = Index + Len(DoSimulatedPark)
    Call CopyMemory(ptrByteArray(Index), AltD, Len(AltD))
    Index = Index + Len(AltD)
    Call CopyMemory(ptrByteArray(Index), AltM, Len(AltM))
    Index = Index + Len(AltM)
    Call CopyMemory(ptrByteArray(Index), AltS, Len(AltS))
    Index = Index + Len(AltS)
    Call CopyMemory(ptrByteArray(Index), AzimD, Len(AzimD))
    Index = Index + Len(AzimD)
    Call CopyMemory(ptrByteArray(Index), AzimM, Len(AzimM))
    Index = Index + Len(AzimM)
    Call CopyMemory(ptrByteArray(Index), AzimS, Len(AzimS))
    Index = Index + Len(AzimS)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)

    Call CopyMemory(ptrByteArray(Index), ParkRotator, Len(ParkRotator))
    Index = Index + Len(ParkRotator)

    Call CopyMemory(ptrByteArray(Index), DoHomePark, Len(DoHomePark))
    Index = Index + Len(DoHomePark)

    Call CopyMemory(ptrByteArray(Index), DoTrackingOff, Len(DoTrackingOff))
    Index = Index + Len(DoTrackingOff)

End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(DummyVar, ptrByteArray(Index), Len(DummyVar))
    Index = Index + Len(DummyVar)

    If SavedVersionNumber >= 2 Then
        Call CopyMemory(DoSimulatedPark, ptrByteArray(Index), Len(DoSimulatedPark))
        Index = Index + Len(DoSimulatedPark)
        Call CopyMemory(AltD, ptrByteArray(Index), Len(AltD))
        Index = Index + Len(AltD)
        Call CopyMemory(AltM, ptrByteArray(Index), Len(AltM))
        Index = Index + Len(AltM)
        Call CopyMemory(AltS, ptrByteArray(Index), Len(AltS))
        Index = Index + Len(AltS)
        Call CopyMemory(AzimD, ptrByteArray(Index), Len(AzimD))
        Index = Index + Len(AzimD)
        Call CopyMemory(AzimM, ptrByteArray(Index), Len(AzimM))
        Index = Index + Len(AzimM)
        Call CopyMemory(AzimS, ptrByteArray(Index), Len(AzimS))
        Index = Index + Len(AzimS)
    Else
        DoSimulatedPark = False
        AltD = 0
        AltM = 0
        AltS = 0
        AzimD = 0
        AzimM = 0
        AzimS = 0
    End If

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If

    If SavedVersionNumber >= 4 Then
        Call CopyMemory(ParkRotator, ptrByteArray(Index), Len(ParkRotator))
        Index = Index + Len(ParkRotator)
    Else
        ParkRotator = True
    End If

    If SavedVersionNumber >= 5 Then
        Call CopyMemory(DoHomePark, ptrByteArray(Index), Len(DoHomePark))
        Index = Index + Len(DoHomePark)
    Else
        DoHomePark = False
    End If

    If SavedVersionNumber >= 6 Then
        Call CopyMemory(DoTrackingOff, ptrByteArray(Index), Len(DoTrackingOff))
        Index = Index + Len(DoTrackingOff)
    Else
        DoTrackingOff = False
    End If
End Sub

Public Function BuildActionListString() As String
    If Not DoSimulatedPark And Not DoHomePark And Not DoTrackingOff Then
        BuildActionListString = "Park Mount."
    ElseIf DoSimulatedPark Then
        BuildActionListString = "Simulated Park."
    ElseIf DoHomePark Then
        BuildActionListString = "Home & Tracking Off."
    ElseIf DoTrackingOff Then
        BuildActionListString = "Tracking Off."
    End If
End Function

