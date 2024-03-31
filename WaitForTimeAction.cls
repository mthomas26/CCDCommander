VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaitForTimeAction"
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

Public Hour As Integer
Public Minute As Integer
Public Second As Integer

Public Selected As Boolean

Private VersionNumber As Long

Public AbsoluteTime As Boolean

Private Sub Class_Initialize()
    VersionNumber = 4
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(Hour)
    MySize = MySize + Len(Minute)
    
    MySize = MySize + Len(Selected)
    
    MySize = MySize + Len(AbsoluteTime)
    
    MySize = MySize + Len(Second)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), Hour, Len(Hour))
    Index = Index + Len(Hour)
    Call CopyMemory(ptrByteArray(Index), Minute, Len(Minute))
    Index = Index + Len(Minute)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)

    Call CopyMemory(ptrByteArray(Index), AbsoluteTime, Len(AbsoluteTime))
    Index = Index + Len(AbsoluteTime)

    Call CopyMemory(ptrByteArray(Index), Second, Len(Second))
    Index = Index + Len(Second)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(Hour, ptrByteArray(Index), Len(Hour))
    Index = Index + Len(Hour)
    Call CopyMemory(Minute, ptrByteArray(Index), Len(Minute))
    Index = Index + Len(Minute)
    
    If SavedVersionNumber >= 2 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(AbsoluteTime, ptrByteArray(Index), Len(AbsoluteTime))
        Index = Index + Len(AbsoluteTime)
    Else
        AbsoluteTime = True
    End If

    If SavedVersionNumber >= 4 Then
        Call CopyMemory(Second, ptrByteArray(Index), Len(Second))
        Index = Index + Len(Second)
    Else
        Second = 0
    End If
End Sub

Public Sub LoadAction(FileNum As Integer)
    Dim SavedVersionNumber As Long
    Get #FileNum, , SavedVersionNumber
    Get #FileNum, , Hour
    Get #FileNum, , Minute
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String
    Dim Time As Date
    
    Time = Hour & ":" & Minute & ":" & Second
    
    If Me.AbsoluteTime Then
        ListString = "Wait until " & Format(Time, "h:nn:ss")
    Else
        If Hour > 0 Then
            ListString = "Wait for " & Hour & " hours, " & Minute & " minutes, " & Second & " seconds."
        ElseIf Minute > 0 Then
            ListString = "Wait for " & Minute & " minutes, " & Second & " seconds."
        Else
            ListString = "Wait for " & Second & " seconds."
        End If
    End If

    BuildActionListString = ListString
End Function

