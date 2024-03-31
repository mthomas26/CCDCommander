VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntelligentTempAction"
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

Public CoolerOn As Boolean
Public MaxCoolerPower As Double
Public MaxTime As Double
Public Deviation As Double
Public CoolerDeviation As Double
Public FanOn As Integer
Public IntelligentAction As Integer

Public RampWarmUp As Integer

Const MaxTemps = 20
Private myDesiredTemperatures(MaxTemps) As Double
Private myNumTemps As Long

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 6
End Sub

Public Property Let NumTemps(Value As Long)
    If Value > MaxTemps Then Value = MaxTemps
    
    myNumTemps = Value
End Property

Public Property Get NumTemps() As Long
    NumTemps = myNumTemps
End Property

Public Property Let DesiredTemperatures(Index As Long, Value As Double)
    If Index >= 0 And Index <= NumTemps Then
        myDesiredTemperatures(Index) = Value
    End If
End Property

Public Property Get DesiredTemperatures(Index As Long) As Double
    If Index >= 0 And Index <= NumTemps Then
        DesiredTemperatures = myDesiredTemperatures(Index)
    Else
        DesiredTemperatures = 99
    End If
End Property

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim Counter As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(CoolerOn)
    MySize = MySize + Len(MaxCoolerPower)
    MySize = MySize + Len(MaxTime)
    MySize = MySize + Len(Deviation)
    MySize = MySize + Len(CoolerDeviation)
    MySize = MySize + Len(FanOn)
    MySize = MySize + Len(IntelligentAction)
    MySize = MySize + Len(myNumTemps)
    For Counter = 0 To MaxTemps
        MySize = MySize + Len(myDesiredTemperatures(Counter))
    Next Counter
    MySize = MySize + Len(RampWarmUp)
    
    MySize = MySize + Len(Selected)

    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim Counter As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), CoolerOn, Len(CoolerOn))
    Index = Index + Len(CoolerOn)
    Call CopyMemory(ptrByteArray(Index), MaxCoolerPower, Len(MaxCoolerPower))
    Index = Index + Len(MaxCoolerPower)
    Call CopyMemory(ptrByteArray(Index), MaxTime, Len(MaxTime))
    Index = Index + Len(MaxTime)
    Call CopyMemory(ptrByteArray(Index), Deviation, Len(Deviation))
    Index = Index + Len(Deviation)
    Call CopyMemory(ptrByteArray(Index), CoolerDeviation, Len(CoolerDeviation))
    Index = Index + Len(CoolerDeviation)
    Call CopyMemory(ptrByteArray(Index), FanOn, Len(FanOn))
    Index = Index + Len(FanOn)
    Call CopyMemory(ptrByteArray(Index), IntelligentAction, Len(IntelligentAction))
    Index = Index + Len(IntelligentAction)
    Call CopyMemory(ptrByteArray(Index), myNumTemps, Len(myNumTemps))
    Index = Index + Len(myNumTemps)

    For Counter = 0 To MaxTemps
        Call CopyMemory(ptrByteArray(Index), myDesiredTemperatures(Counter), Len(myDesiredTemperatures(Counter)))
        Index = Index + Len(myDesiredTemperatures(Counter))
    Next Counter

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)

    Call CopyMemory(ptrByteArray(Index), RampWarmUp, Len(RampWarmUp))
    Index = Index + Len(RampWarmUp)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim Counter As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(CoolerOn, ptrByteArray(Index), Len(CoolerOn))
    Index = Index + Len(CoolerOn)
    Call CopyMemory(MaxCoolerPower, ptrByteArray(Index), Len(MaxCoolerPower))
    Index = Index + Len(MaxCoolerPower)
    Call CopyMemory(MaxTime, ptrByteArray(Index), Len(MaxTime))
    Index = Index + Len(MaxTime)
    Call CopyMemory(Deviation, ptrByteArray(Index), Len(Deviation))
    Index = Index + Len(Deviation)

    If SavedVersionNumber >= 2 Then
        Call CopyMemory(CoolerDeviation, ptrByteArray(Index), Len(CoolerDeviation))
        Index = Index + Len(CoolerDeviation)
    Else
        CoolerDeviation = 4
    End If

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(FanOn, ptrByteArray(Index), Len(FanOn))
        Index = Index + Len(FanOn)
    Else
        FanOn = 1
    End If
    
    If SavedVersionNumber >= 4 Then
        Call CopyMemory(IntelligentAction, ptrByteArray(Index), Len(IntelligentAction))
        Index = Index + Len(IntelligentAction)
    Else
        IntelligentAction = 1
    End If

    Call CopyMemory(myNumTemps, ptrByteArray(Index), Len(myNumTemps))
    Index = Index + Len(myNumTemps)
    For Counter = 0 To MaxTemps
        Call CopyMemory(myDesiredTemperatures(Counter), ptrByteArray(Index), Len(myDesiredTemperatures(Counter)))
        Index = Index + Len(myDesiredTemperatures(Counter))
    Next Counter

    If SavedVersionNumber >= 5 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If

    If SavedVersionNumber >= 6 Then
        Call CopyMemory(RampWarmUp, ptrByteArray(Index), Len(RampWarmUp))
        Index = Index + Len(RampWarmUp)
    Else
        RampWarmUp = 1
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    If Me.IntelligentAction = 1 Or Me.CoolerOn = False Then
        ListString = "Intelligent Cooler Action: Cooler "
    
        If Me.CoolerOn Then
            ListString = ListString & "On, Max Power " & Me.MaxCoolerPower & "%"
        Else
            ListString = ListString & "Off"
        End If
    Else
        ListString = "Simple Cooler Action: Cooler to " & Format(Me.DesiredTemperatures(0), "0.0") & " " & Chr(176) & "C"
    End If
        
    If Me.FanOn = 1 Then
        ListString = ListString & ", Fan On."
    Else
        ListString = ListString & ", Fan Off."
    End If

    BuildActionListString = ListString
End Function

