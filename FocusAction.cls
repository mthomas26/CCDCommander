VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FocusAction"
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

Public ImagerFilter As Integer
Public UseOffset As Integer
Public Offset As Long
Public ExposureTime As Double

Public Selected As Boolean

Public SpecifyFocusStarPosition As Boolean
Public StarXPosition As Long
Public StarYPosition As Long

Public FastReadout As Boolean

Public FocusAverages As Long

Public TempComp As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 8
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(ImagerFilter)
    MySize = MySize + Len(UseOffset)
    MySize = MySize + Len(Offset)
    
    MySize = MySize + Len(Selected)

    MySize = MySize + Len(ExposureTime)

    MySize = MySize + Len(SpecifyFocusStarPosition)
    MySize = MySize + Len(StarXPosition)
    MySize = MySize + Len(StarYPosition)
    
    MySize = MySize + Len(FastReadout)
    
    MySize = MySize + Len(FocusAverages)
    MySize = MySize + Len(TempComp)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), ImagerFilter, Len(ImagerFilter))
    Index = Index + Len(ImagerFilter)
    Call CopyMemory(ptrByteArray(Index), UseOffset, Len(UseOffset))
    Index = Index + Len(UseOffset)
    Call CopyMemory(ptrByteArray(Index), Offset, Len(Offset))
    Index = Index + Len(Offset)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)

    Call CopyMemory(ptrByteArray(Index), ExposureTime, Len(ExposureTime))
    Index = Index + Len(ExposureTime)

    Call CopyMemory(ptrByteArray(Index), SpecifyFocusStarPosition, Len(SpecifyFocusStarPosition))
    Index = Index + Len(SpecifyFocusStarPosition)
    Call CopyMemory(ptrByteArray(Index), StarXPosition, Len(StarXPosition))
    Index = Index + Len(StarXPosition)
    Call CopyMemory(ptrByteArray(Index), StarYPosition, Len(StarYPosition))
    Index = Index + Len(StarYPosition)

    Call CopyMemory(ptrByteArray(Index), FastReadout, Len(FastReadout))
    Index = Index + Len(FastReadout)

    Call CopyMemory(ptrByteArray(Index), FocusAverages, Len(FocusAverages))
    Index = Index + Len(FocusAverages)
    
    Call CopyMemory(ptrByteArray(Index), TempComp, Len(TempComp))
    Index = Index + Len(TempComp)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    
    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(ImagerFilter, ptrByteArray(Index), Len(ImagerFilter))
    Index = Index + Len(ImagerFilter)

    If SavedVersionNumber >= 2 Then
        Call CopyMemory(UseOffset, ptrByteArray(Index), Len(UseOffset))
        Index = Index + Len(UseOffset)
        Call CopyMemory(Offset, ptrByteArray(Index), Len(Offset))
        Index = Index + Len(Offset)
    Else
        UseOffset = 0
        Offset = 0
    End If

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If
    
    If SavedVersionNumber >= 4 Then
        Call CopyMemory(ExposureTime, ptrByteArray(Index), Len(ExposureTime))
        Index = Index + Len(ExposureTime)
    Else
        ExposureTime = 0.11
    End If

    If SavedVersionNumber >= 5 Then
        Call CopyMemory(SpecifyFocusStarPosition, ptrByteArray(Index), Len(SpecifyFocusStarPosition))
        Index = Index + Len(SpecifyFocusStarPosition)
        Call CopyMemory(StarXPosition, ptrByteArray(Index), Len(StarXPosition))
        Index = Index + Len(StarXPosition)
        Call CopyMemory(StarYPosition, ptrByteArray(Index), Len(StarYPosition))
        Index = Index + Len(StarYPosition)
    Else
        SpecifyFocusStarPosition = False
        StarXPosition = 0
        StarYPosition = 0
    End If

    If SavedVersionNumber >= 6 Then
        Call CopyMemory(FastReadout, ptrByteArray(Index), Len(FastReadout))
        Index = Index + Len(FastReadout)
    Else
        FastReadout = False
    End If

    If SavedVersionNumber >= 7 Then
        Call CopyMemory(FocusAverages, ptrByteArray(Index), Len(FocusAverages))
        Index = Index + Len(FocusAverages)
    Else
        FocusAverages = 1
    End If
    
    If SavedVersionNumber >= 8 Then
        Call CopyMemory(TempComp, ptrByteArray(Index), Len(TempComp))
        Index = Index + Len(TempComp)
    Else
        TempComp = False
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    If UseOffset Then
        If ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
            ListString = "Offset Focuser " & Offset & " steps for " & frmOptions.lstFilters.List(ImagerFilter) & " filter."
        Else
            ListString = "Offset Focuser " & Offset & " steps."
        End If
    Else
        If ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
            ListString = "Focus with " & frmOptions.lstFilters.List(ImagerFilter) & " filter."
        Else
            ListString = "Focus."
        End If
    End If

    BuildActionListString = ListString
End Function

