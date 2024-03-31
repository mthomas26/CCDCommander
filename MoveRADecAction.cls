VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoveRADecAction"
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

Public Enum EpochTypes
    JNow
    J2000
End Enum

Public RA As Double
Public Dec As Double
Public Name As String
Public RecomputeObjectCoordinates As Integer
Public Epoch As EpochTypes

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 5
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(RA)
    MySize = MySize + Len(Dec)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(Name)
    MySize = MySize + Len(RecomputeObjectCoordinates)
    MySize = MySize + Len(Epoch)
    
    MySize = MySize + Len(Selected)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), RA, Len(RA))
    Index = Index + Len(RA)
    Call CopyMemory(ptrByteArray(Index), Dec, Len(Dec))
    Index = Index + Len(Dec)
    
    StringLen = Len(Name)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray(), Index, Name)
    
    Call CopyMemory(ptrByteArray(Index), RecomputeObjectCoordinates, Len(RecomputeObjectCoordinates))
    Index = Index + Len(RecomputeObjectCoordinates)
    Call CopyMemory(ptrByteArray(Index), Epoch, Len(Epoch))
    Index = Index + Len(Epoch)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(RA, ptrByteArray(Index), Len(RA))
    Index = Index + Len(RA)
    Call CopyMemory(Dec, ptrByteArray(Index), Len(Dec))
    Index = Index + Len(Dec)

    If SavedVersionNumber >= 2 Then
        Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
        Index = Index + Len(StringLen)
        Name = String(StringLen, " ")
        Call CopyByteArrayToString(Name, ptrByteArray(), Index)
    Else
        Name = ""
    End If
    
    If SavedVersionNumber >= 3 Then
        Call CopyMemory(RecomputeObjectCoordinates, ptrByteArray(Index), Len(RecomputeObjectCoordinates))
        Index = Index + Len(RecomputeObjectCoordinates)
    Else
        RecomputeObjectCoordinates = 0
    End If
    
    If SavedVersionNumber >= 4 Then
        Call CopyMemory(Epoch, ptrByteArray(Index), Len(Epoch))
        Index = Index + Len(Epoch)
    Else
        Epoch = JNow
    End If

    If SavedVersionNumber >= 5 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String
    
    If Name = "" Then
        ListString = "Move to " & Misc.ConvertEquatorialToString(RA, Dec, False)
        If Epoch = J2000 Then
            ListString = ListString & " (J2000)"
        End If
    Else
        ListString = "Move to " & Name & "."
    End If

    BuildActionListString = ListString
End Function


