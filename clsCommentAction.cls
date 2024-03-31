VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommentAction"
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

Public Comment As String
Public SendAsEmail As Boolean
Public DoNotSkip As Boolean

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 3
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(Comment)
    
    MySize = MySize + Len(Selected)
    
    MySize = MySize + Len(SendAsEmail)
    
    MySize = MySize + Len(DoNotSkip)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    
    StringLen = Len(Comment)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray, Index, Comment)     'Index is passed ByRef, so I don't need to increment it

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)
    
    Call CopyMemory(ptrByteArray(Index), SendAsEmail, Len(SendAsEmail))
    Index = Index + Len(SendAsEmail)

    Call CopyMemory(ptrByteArray(Index), DoNotSkip, Len(DoNotSkip))
    Index = Index + Len(DoNotSkip)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    
    Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
    Index = Index + Len(StringLen)
    Comment = String(StringLen, " ")
    Call CopyByteArrayToString(Comment, ptrByteArray(), Index)
    
    Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
    Index = Index + Len(Selected)
    
    If SavedVersionNumber >= 2 Then
        Call CopyMemory(SendAsEmail, ptrByteArray(Index), Len(SendAsEmail))
        Index = Index + Len(SendAsEmail)
    Else
        SendAsEmail = False
    End If

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(DoNotSkip, ptrByteArray(Index), Len(DoNotSkip))
        Index = Index + Len(DoNotSkip)
    Else
        DoNotSkip = True
    End If
End Sub

Public Sub LoadAction(FileNum As Integer)
    'deprecated
End Sub

Public Function BuildActionListString() As String
    Dim MaxChars As Integer
    
    MaxChars = InStr(Comment, Chr(13))
    If MaxChars > 80 Or MaxChars = 0 Then
        MaxChars = 80
    End If
    BuildActionListString = "Comment: " & Left(Comment, MaxChars)
End Function


