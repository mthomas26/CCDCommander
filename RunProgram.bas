Attribute VB_Name = "RunProgram"
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
 
'Private Const SW_HIDE = 0
'Private Const SW_SHOWNORMAL = 1
'Private Const SW_NORMAL = 1
'Private Const SW_SHOWMINIMIZED = 2
'Private Const SW_SHOWMAXIMIZED = 3
'Private Const SW_MAXIMIZE = 3
'Private Const SW_SHOWNOACTIVATE = 4
'Private Const SW_SHOW = 5
'Private Const SW_MINIMIZE = 6
'Private Const SW_SHOWMINNOACTIVE = 7
'Private Const SW_SHOWNA = 8
'Private Const SW_RESTORE = 9
'Private Const SW_SHOWDEFAULT = 10
'Private Const SW_MAX = 10
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Sub ShellAndWait(ByVal strPath As String, ByVal iWindowStyle As Integer, ByRef lreturnCode As Long, _
    Optional sWinTitle As String = "", Optional sDirectoryPath As String = "")

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long

    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start) ' you must set the size
    start.dwFlags = &H1& ' STARTF_USESHOWWINDOW Use Show Window
    start.wShowWindow = iWindowStyle
    If Not IsMissing(sWinTitle) Then
        ' if there is a title set the window title
        start.lpTitle = sWinTitle
    End If

    ' Start the shelled application:
    ret = CreateProcessA(0&, strPath, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, _
        sDirectoryPath, start, _
        proc)

    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, 100&)
    Do While ret <> 0
        If ret < 0 Then
            Exit Sub
        End If
        Call Wait(1)

        ret = WaitForSingleObject(proc.hProcess, _
           100&)
    Loop

    'get the return code
    ret = GetExitCodeProcess(proc.hProcess, _
        lreturnCode)

    'close the process handles
    ret = CloseHandle(proc.hProcess)
End Sub

Public Sub ShellNoWait(ByVal strPath As String, ByVal iWindowStyle As Integer, Optional sWinTitle As String = "", _
    Optional sDirectoryPath As String = "")

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long

    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start) ' you must set the size
    start.dwFlags = &H1& ' STARTF_USESHOWWINDOW Use Show Window
    start.wShowWindow = iWindowStyle
    If Not IsMissing(sWinTitle) Then
        ' if there is a title set the window title
        start.lpTitle = sWinTitle
    End If

    ' Start the shelled application:
    ret = CreateProcessA(0&, strPath, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, _
        sDirectoryPath, start, _
        proc)

'    'close the process handles
'    Ret = CloseHandle(proc.hProcess)
End Sub


Public Sub RunScript(clsAction As RunScriptAction)
    Dim lRet As Long

    If clsAction.WaitForScript = vbChecked Then
        If clsAction.ProgramIsScript Then
            Call AddToStatus("Running script from " & clsAction.ScriptName)
            Call ShellAndWait("wscript.exe " & Chr(34) & clsAction.ScriptName & Chr(34) & " " & clsAction.ScriptArguments, 1, lRet, "", App.Path)
            Call AddToStatus("Script complete.")
        Else
            Call AddToStatus("Running program from " & clsAction.ScriptName)
            Call ShellAndWait(Chr(34) & clsAction.ScriptName & Chr(34) & " " & clsAction.ScriptArguments, 1, lRet, "", App.Path)
            Call AddToStatus("Program complete.")
        End If
    Else
        If clsAction.ProgramIsScript Then
            Call AddToStatus("Running script from " & clsAction.ScriptName)
            Call ShellNoWait("wscript.exe " & Chr(34) & clsAction.ScriptName & Chr(34) & " " & clsAction.ScriptArguments, 1, "", App.Path)
        Else
            Call AddToStatus("Running program from " & clsAction.ScriptName)
            Call ShellNoWait(Chr(34) & clsAction.ScriptName & Chr(34) & " " & clsAction.ScriptArguments, 1, "", App.Path)
        End If
    End If
End Sub

Public Sub RunScriptDirect(ScriptName As String, WaitForComplete As Boolean, Optional Arguments As String = "")
    Dim lRet As Long

    If (ScriptName <> "") Then
        If WaitForComplete = True Then
            Call AddToStatus("Running script from " & ScriptName)
            Call ShellAndWait("wscript.exe " & Chr(34) & ScriptName & Chr(34) & " " & Arguments, 1, lRet, "", App.Path)
            Call AddToStatus("Script complete.")
        Else
            Call AddToStatus("Running script from " & ScriptName)
            Call ShellNoWait("wscript.exe " & Chr(34) & ScriptName & Chr(34) & " " & Arguments, 1, "", App.Path)
        End If
    End If
End Sub
