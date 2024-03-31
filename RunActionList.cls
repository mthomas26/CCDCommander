VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RunActionList"
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

Public ActionListName As String
Public RepeatMode As Integer
Public TimesToRepeat As Integer
Public RepeatTime As Double

Public ActionCollection As Collection

Public LinkToFile As Boolean

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 4
End Sub

'Special Note about the ActionCollection and saving/loading
'The action collection is saved/loaded in the save/load functions, not here
'This allows the save/load functions to recurse through all of the collections in a particular setup.

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(ActionListName)
    MySize = MySize + Len(RepeatMode)
    MySize = MySize + Len(TimesToRepeat)
    MySize = MySize + Len(RepeatTime)
    MySize = MySize + Len(LinkToFile)
    
    MySize = MySize + Len(Selected)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    
    StringLen = Len(ActionListName)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray(), Index, ActionListName)
    
    Call CopyMemory(ptrByteArray(Index), RepeatMode, Len(RepeatMode))
    Index = Index + Len(RepeatMode)
    Call CopyMemory(ptrByteArray(Index), TimesToRepeat, Len(TimesToRepeat))
    Index = Index + Len(TimesToRepeat)
    Call CopyMemory(ptrByteArray(Index), RepeatTime, Len(RepeatTime))
    Index = Index + Len(RepeatTime)
    
    Call CopyMemory(ptrByteArray(Index), LinkToFile, Len(LinkToFile))
    Index = Index + Len(LinkToFile)
    
    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    
    Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
    Index = Index + Len(StringLen)
    ActionListName = String(StringLen, " ")
    Call CopyByteArrayToString(ActionListName, ptrByteArray(), Index)
    
    If SavedVersionNumber >= 2 Then
        Call CopyMemory(RepeatMode, ptrByteArray(Index), Len(RepeatMode))
        Index = Index + Len(RepeatMode)
        Call CopyMemory(TimesToRepeat, ptrByteArray(Index), Len(TimesToRepeat))
        Index = Index + Len(TimesToRepeat)
        Call CopyMemory(RepeatTime, ptrByteArray(Index), Len(RepeatTime))
        Index = Index + Len(RepeatTime)
    Else
        RepeatMode = 0
        TimesToRepeat = 2
        RepeatTime = 1
    End If
    
    If SavedVersionNumber >= 3 Then
        Call CopyMemory(LinkToFile, ptrByteArray(Index), Len(LinkToFile))
        Index = Index + Len(LinkToFile)
    Else
        LinkToFile = True
    End If

    If SavedVersionNumber >= 4 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    If LinkToFile Then
        ListString = "Run list in " & Mid(ActionListName, InStrRev(ActionListName, "\") + 1)
    Else
        If ActionListName <> "" Then
            If InStrRev(ActionListName, "\") > 0 Then
                'still looks like a file name/path, parse just the name
                ListString = "Run list named " & Mid(ActionListName, InStrRev(ActionListName, "\") + 1)
            Else
                ListString = "Run list named " & ActionListName
            End If
        Else
            ListString = "Run sub-action list"
        End If
    End If

    Select Case RepeatMode
        Case 0
            ListString = ListString & " once."
        Case 1
            ListString = ListString & " " & TimesToRepeat & " times."
        Case 2
            ListString = ListString & " for " & Format(RepeatTime, "0.00") & " minutes."
        Case 3
            ListString = ListString & " forever."
    End Select
    
    BuildActionListString = ListString
End Function

