VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DomeAction"
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

Public Enum DomeActionTypes
    actCloseDome = 0
    actOpenDome = 1
    actCoupleDome = 2
    actUnCoupleDome = 3
    actHomeDome = 4
    actParkDome = 5
    actSlewDomeAzimuth = 6
End Enum

Public ThisDomeAction As DomeActionTypes
Public Azimuth As Integer

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 3
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(ThisDomeAction)
    
    MySize = MySize + Len(Selected)
    
    MySize = MySize + Len(Azimuth)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), ThisDomeAction, Len(ThisDomeAction))
    Index = Index + Len(ThisDomeAction)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)

    Call CopyMemory(ptrByteArray(Index), Azimuth, Len(Azimuth))
    Index = Index + Len(Azimuth)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(ThisDomeAction, ptrByteArray(Index), Len(ThisDomeAction))
    Index = Index + Len(ThisDomeAction)

    If SavedVersionNumber >= 2 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(Azimuth, ptrByteArray(Index), Len(Azimuth))
        Index = Index + Len(Azimuth)
    Else
        Azimuth = 0
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    If ThisDomeAction = actCloseDome Then
        ListString = "Close dome."
    ElseIf ThisDomeAction = actOpenDome Then
        ListString = "Open dome."
    ElseIf ThisDomeAction = actCoupleDome Then
        ListString = "Couple the dome to the mount."
    ElseIf ThisDomeAction = actUnCoupleDome Then
        ListString = "Uncouple the dome from the mount."
    ElseIf ThisDomeAction = actHomeDome Then
        ListString = "Home dome."
    ElseIf ThisDomeAction = actParkDome Then
        ListString = "Park dome."
    ElseIf ThisDomeAction = actSlewDomeAzimuth Then
        ListString = "Slew dome to azimuth = " & Azimuth & " degrees."
    End If

    BuildActionListString = ListString
End Function

