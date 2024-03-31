VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RunScriptAction"
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

Public ScriptName As String
Public ScriptArguments As String
Public WaitForScript As Integer
Public ProgramIsScript As Boolean

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 4
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(ScriptName)
    MySize = MySize + Len(WaitForScript)
    MySize = MySize + Len(ProgramIsScript)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(ScriptArguments)
    
    MySize = MySize + Len(Selected)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    
    StringLen = Len(ScriptName)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray(), Index, ScriptName)
    
    Call CopyMemory(ptrByteArray(Index), WaitForScript, Len(WaitForScript))
    Index = Index + Len(WaitForScript)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)

    Call CopyMemory(ptrByteArray(Index), ProgramIsScript, Len(ProgramIsScript))
    Index = Index + Len(ProgramIsScript)

    StringLen = Len(ScriptArguments)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray(), Index, ScriptArguments)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    
    Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
    Index = Index + Len(StringLen)
    ScriptName = String(StringLen, " ")
    Call CopyByteArrayToString(ScriptName, ptrByteArray(), Index)

    Call CopyMemory(WaitForScript, ptrByteArray(Index), Len(WaitForScript))
    Index = Index + Len(WaitForScript)
    
    If SavedVersionNumber >= 2 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If
    
    If SavedVersionNumber >= 3 Then
        Call CopyMemory(ProgramIsScript, ptrByteArray(Index), Len(ProgramIsScript))
        Index = Index + Len(ProgramIsScript)
    Else
        ProgramIsScript = True
    End If
    
    If SavedVersionNumber >= 4 Then
        Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
        Index = Index + Len(StringLen)
        ScriptArguments = String(StringLen, " ")
        Call CopyByteArrayToString(ScriptArguments, ptrByteArray(), Index)
    Else
        ScriptArguments = ""
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    ListString = "Run " & Mid(ScriptName, InStrRev(ScriptName, "\") + 1) & " " & ScriptArguments
    If WaitForScript = vbChecked Then
        ListString = ListString & " and wait."
    End If

    BuildActionListString = ListString
End Function

