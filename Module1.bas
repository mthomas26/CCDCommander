Attribute VB_Name = "MainMod"
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

Public colAction As New Collection
Public colAutoguiderErrors As New Collection
Private SkipAheadTimes As Collection
Private SkipAheadSoftSkips As Collection
Public Aborted As Boolean
Public Paused As Boolean
Public Pausing As Boolean
Public PauseBetweenActions As Boolean
Public SkipToNextMoveAction As Boolean
Public RetryMoveAction As Boolean
Public SkipToNextSkipToAction As Boolean
Public AbortButton As Boolean
Public SoftSkip As Boolean
Public RestartActionList As Boolean

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Sub get_Result Lib "sbcb3.dll" (ByRef Val As Long)
Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" (ByRef lpSystemTimeAsFileTime As FILETIME)

Private StatusFile As Integer
Private MissedTargetFile As Integer

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
    ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Exiting As Boolean

Public Enum FrameSizes
    FullFrame
    HalfFrame
    QuarterFrame
    CustomFrame
End Enum

Public RunSelectedActionsOnly As Boolean

Public RunningActionListLevel As Integer

Public FollowRunningAction As Boolean

' Test flags
Public PlateSolveNoSaveImage As Boolean
Public UseFakeGuideExposure As Boolean
Public UseFakeFlatImage As Boolean

Public MoveToRADecPlateSolveStatus As Boolean

' Disabling watchdog for initial open source release
'Private Watchdog As Object

Sub Main()
    Dim RetVal As Long
    Dim commandline As String
    Dim NextCommand As String
    Dim ImportedActionCollection As Collection
    Dim AutoRun As Boolean
    Dim AutoLoad As Boolean
    Dim AutoLoadFileName As String
    Dim TestFileNum As Integer
        
'#If Not EnaBisqueEventDLL Then
'    Call get_Result(RetVal)
'
'    If RetVal <> 0 Then
'        MsgBox "Undetermined error.  Cannot continue."
'        GoTo MyEnd
'    End If
'#End If

    If Left(CheckFileAssociation(".act"), 10) <> "actionlist" Then
        If MsgBox("CCD Commander is not the default program for Action List Files." & vbCrLf & "Do you want to make this the default program?", vbYesNo + vbQuestion, "CCD Commander") = vbYes Then
            Call Associate_File(".act", App.Path & "\CCDCommander.exe", "actionlist", "CCD Commander Action List", App.Path & "\CCDCommander.exe")
        End If
    End If
    
    commandline = Command

    'Clear all test flags
    PlateSolveNoSaveImage = False
    
    Exiting = False
    
    App.HelpFile = App.Path & "\CCDCommander.chm"

    Settings.FocusMaxDisconnectReconnect = False
    Settings.WeatherMonitorRestartActionList = False
    
    UseFakeGuideExposure = False
    MoveToRADecPlateSolveStatus = False
    UseFakeFlatImage = False

    AutoRun = False
    AutoLoad = False
    If commandline <> "" Then
        'First assume it is a file name...if the load fails, then try parsing something else
        AutoLoadFileName = commandline
        
        AutoLoadFileName = FormatFileName(commandline)
        
        'try opening file now
        TestFileNum = FreeFile()
        On Error Resume Next
        Open AutoLoadFileName For Input Access Read As #TestFileNum
        
        If Err.Number = 0 Then
            'file opened successfully, good to go
            AutoLoad = True
        Else
            AutoLoad = False
        End If
        
        Close #TestFileNum
        On Error GoTo 0
        
        If AutoLoad Then
            'do nothing here
        ElseIf UCase(Left(commandline, 7)) = "AUTORUN" Then
            AutoLoadFileName = FormatFileName(Mid(commandline, 9))
            
            AutoLoad = True
            AutoRun = True
        ElseIf UCase(Left(commandline, 5)) = "RESET" Then
            commandline = Mid(commandline, 7)
            Do
                If InStr(commandline, " ") Then
                    NextCommand = Left(commandline, InStr(commandline, " ") - 1)
                    commandline = Mid(commandline, InStr(commandline, " ") + 1)
                Else
                    NextCommand = commandline
                    commandline = ""
                End If
                
                Select Case UCase(NextCommand)
                    Case "WINDOWPOS"
                        Call DeleteMySetting("WindowPositions")
                    Case "MAXIMCOUNT"
                        Call SaveMySetting("ProgramSettings", "MaxImSequenceNum", "0")
                    Case "ALL"
                        Call DeleteMySetting("AutoFlatAction")
                        Call DeleteMySetting("AutoGuideExposure")
                        Call DeleteMySetting("CameraAction")
                        Call DeleteMySetting("CloudMonitorAction")
                        Call DeleteMySetting("CommentAction")
                        Call DeleteMySetting("DomeAction")
                        Call DeleteMySetting("FocusAction")
                        Call DeleteMySetting("ImageLinkAction")
                        Call DeleteMySetting("Import Utility")
                        Call DeleteMySetting("IntelligentTempAction")
                        Call DeleteMySetting("MountParameters")
                        Call DeleteMySetting("MoveRADecAction")
                        Call DeleteMySetting("ParkAction")
                        Call DeleteMySetting("ProgramSettings")
                        Call DeleteMySetting("RotatorAction")
                        Call DeleteMySetting("RunActionList")
                        Call DeleteMySetting("RunScriptAction")
                        Call DeleteMySetting("SkipAheadAtTime")
                        Call DeleteMySetting("SkipAheadAtAlt")
                        Call DeleteMySetting("Startup")
                        Call DeleteMySetting("TempGraph")
                        Call DeleteMySetting("WaitForAlt")
                        Call DeleteMySetting("WaitForTime")
                        Call DeleteMySetting("WindowPositions")
                End Select
            Loop Until Len(commandline) = 0
        ElseIf UCase(Left(commandline, 6)) = "IMPORT" Then
            commandline = Mid(commandline, 8)
            
            Call SetupStuff
            
            Load frmImport
        
            Set ImportedActionCollection = New Collection
            
            If Left(commandline, 1) = Chr(34) Then
                frmImport.lblFileName.Caption = FormatFileName(Left(commandline, InStr(Mid(commandline, 2), Chr(34)) + 1))
                commandline = Mid(commandline, InStr(Mid(commandline, 2), Chr(34)) + 3)
            Else
                frmImport.lblFileName.Caption = FormatFileName(Left(commandline, InStr(commandline, " ") - 1))
                commandline = Mid(commandline, InStr(commandline, " ") + 1)
            End If
            
            Call frmImport.ImportToList(ImportedActionCollection)
                
            Call SaveAction(FormatFileName(commandline), ImportedActionCollection)
            
            End
        ElseIf UCase(Left(commandline, 4)) = "TEST" Then
            commandline = Mid(commandline, 6)
            
            Select Case UCase(commandline)
                Case "PLATESOLVENOSAVEIMAGE"
                    PlateSolveNoSaveImage = True
                    
                Case UCase("UseFakeGuideExposure")
                    UseFakeGuideExposure = True
                    
                Case UCase("UseFakeFlatImage")
                    UseFakeFlatImage = True
                    
                Case UCase("FocusMaxDisconnectReconnect")
                    Settings.FocusMaxDisconnectReconnect = True
                    
                Case UCase("WeatherMonitorRestartActionList")
                    Settings.WeatherMonitorRestartActionList = True
                    
            End Select
        End If
    End If
    
    frmMain.Show
    frmMain.Enabled = False
    
    If CheckTrialAndRegistration() = False Then
        GoTo MyEnd
    End If
    
    If Not CBool(GetMySetting("Startup", "DoneGettingStarted2", "False")) Then
        Call HtmlHelp(frmMain.hwnd, App.HelpFile, 15, 3)
        If MsgBox("Have you completed all the Getting Started tasks?", vbYesNo) = vbNo Then
            GoTo MyEnd
        Else
            Call SaveMySetting("Startup", "DoneGettingStarted2", "True")
        End If
    End If

    DoEvents
    
    Call SetupStuff
    Autoguiding = True
    
    If AutoLoad Then
        On Error Resume Next
        Call MainMod.LoadAction(AutoLoadFileName)
        If Err.Number <> 0 Then
            If MsgBox("Could not load file: " & vbCrLf & AutoLoadFileName, vbCritical + vbOKCancel, "CCD Commander") = vbCancel Then
                GoTo MyEnd
            Else
                Call MainMod.ClearAll
            End If
            
            If AutoRun Then AutoRun = False
        Else
            If InStr(commandline, "\") > 0 Then
                frmMain.CurrentFilePath = AutoLoadFileName
                frmMain.CurrentFileName = Mid(AutoLoadFileName, InStrRev(commandline, "\"))
            Else
                frmMain.CurrentFilePath = App.Path & "\" & AutoLoadFileName
                frmMain.CurrentFileName = AutoLoadFileName
            End If
                
            frmMain.Caption = "CCD Commander - " & frmMain.CurrentFileName
        End If
    End If

    frmMain.Enabled = True
    
    If AutoRun Then
        'should probably use the constant - but it isn't available right now
        '14 = EditMenuItems.UseCheckBoxes

        RunSelectedActionsOnly = frmMain.mnuEditItem(14).Checked
        Call frmMain.StartAndAbortAction
    Else
        frmAbout.Show vbModal, frmMain
        Unload frmAbout
    End If
    Exit Sub
    
MyEnd:
    Unload frmMain
End Sub

Public Function FormatFileName(UnformattedFileName As String) As String
    Dim FileName As String
    
    FileName = UnformattedFileName

    'Remove any quotation marks around the file name
    If Left(FileName, 1) = Chr(34) Then
        FileName = Mid(FileName, 2)
    End If
    If Right(FileName, 1) = Chr(34) Then
        FileName = Left(FileName, Len(FileName) - 1)
    End If
    
    'Try to check if it is a relative or absolute path
    If InStr(FileName, ":") > 0 Then
        'absolute path
        'Can be:
        ' C:this.act
        ' D:\1234\abcd\this.act
        ' ftp://mysite.com/this.act
        ' etc...if there is a : it is absolute
    ElseIf Left(FileName, 2) = "\\" Then
        'absolute path
        'Can be a network path like:
        ' \\Desktop\C\Test.act
    Else
        'must be a relative path, prepend app.path to it
        FileName = App.Path & "\" & FileName
    End If
    
    FormatFileName = FileName
End Function

Private Function WatchdogSetup() As Boolean
    Dim lRet As Long
    
    On Error GoTo WatchdogSetupError
' Disabling watchdog for initial open source release
'    Call RunProgram.ShellAndWait(App.Path & "\CCDCommanderWatchdog.exe /unregserver", 1, lRet, "", App.Path)
'    Call RunProgram.ShellAndWait(App.Path & "\CCDCommanderWatchdog.exe /regserver", 1, lRet, "", App.Path)
    
'    Set Watchdog = CreateObject("CCDCWatchdog.Watchdog")
    On Error GoTo 0
    WatchdogSetup = True
    Exit Function

WatchdogSetupError:
    WatchdogSetup = False
    
End Function

Public Sub StartAction()
    Dim clsAction As Object
    Dim myListIndex As Long
    Dim ErrorSource As String
    
    On Error Resume Next
    Call Kill(App.Path & "\Logs\MissedTargets.lst")
    
    Call SaveMySetting("Test", "Aborted", "0")
    
    If (frmOptions.chkEnableWatchdog.Value = vbChecked) Then
' Disabling watchdog for initial open source release
'        On Error Resume Next
'        Set Watchdog = CreateObject("CCDCWatchdog.Watchdog")
'        If Err.Number <> 0 Then
'            On Error GoTo 0
'            If Not WatchdogSetup() Then
'                On Error GoTo StartActionError
'                Call Err.Raise(65001, "CCDCommander", "CCD Commander Watchdog Failed to start.")
'            End If
'        End If
        
        On Error GoTo StartActionError
        
' Disabling watchdog for initial open source release
'        Watchdog.StartWatchdog
    End If
    
    On Error GoTo StartActionError
        
    frmMain.txtStatus.Text = ""
    frmMain.txtCurrentAction.Text = ""
    
    StatusFile = FreeFile()
    Open App.Path & "\Logs\" & Format(Now, "yymmdd_hhmmss") & ".log" For Output As #StatusFile Len = 1
    
    MissedTargetFile = FreeFile()
    Open App.Path & "\Logs\MissedTargets.lst" For Output As #MissedTargetFile
    
    Call WriteHeaderToLogFile
    
    Aborted = False
    SoftSkip = False
    AbortButton = False
    Paused = False
    Pausing = False
    PauseBetweenActions = False
    SkipToNextMoveAction = False
    RetryMoveAction = False
    SkipToNextSkipToAction = False
    FollowRunningAction = True
    
    frmMain.EditingActionNumber = 0
    frmMain.EditingActionListLevel = 0
    
    Call AddToStatus("Action starting.")
    
    On Error Resume Next
    Call ChDir(GetMySetting("ProgramSettings", "SaveToPath", App.Path & "\Images\"))
    If Err.Number <> 0 Then
        'Path doesn't exist - create it
        Call AddToStatus("Global Image Save Location doesn't exist - creating it.")
        Call Misc.CreatePath(GetMySetting("ProgramSettings", "SaveToPath", App.Path & "\Images\"))
    End If
    
    On Error GoTo StartActionError
    
    If Dir(App.Path & "\RunFirst.vbs") <> "" Then
        Call RunProgram.RunScriptDirect(App.Path & "\RunFirst.vbs", True)
    End If
    
#If Not InDebug Then
    Call AddToStatus("Connecting to the Camera...")
    Call CameraSetup
    
    Call AddToStatus("Connecting to the Mount...")
    'set this to false so the mount position is reinitialized
    'this is just in case the mount was moved between actions.
    Mount.TelescopeConnected = False
    
    Call MountSetup
    If InStr(Mount.GetTelescopeType, "Simulator") <> 0 Then
        Call AddToStatus("WARNING! Telescope is set to Simulator!")
        Beep
    End If
    Call Mount.ConnectToTelescope
    
    If frmOptions.lstFocuserControl.ListIndex <> FocusControl.None Then
        Call AddToStatus("Connecting to Focuser...")
        Call FocusSetup
    Else
        Call AddToStatus("WARNING! No Focuser selected!")
        Call FocusSetup 'Call FocusSetup to clear to Focus Object
        Beep
    End If
    
    If frmOptions.lstRotator.ListIndex <> RotatorControl.None Then
        Call AddToStatus("Connecting to Rotator...")
        Call RotatorSetup
    End If
    
    If frmOptions.lstCloudSensor.ListIndex <> WeatherMonitorControl.None Then
        Call AddToStatus("Connecting to Cloud Sensor...")
        Call CloudSensorSetup
    End If
    
    If frmOptions.lstDomeControl.ListIndex <> DomeControlTypes.None Then
        Call AddToStatus("Connecting to the Dome...")
        Call DomeControl.SetupDome
    End If
    
    Call Windows.SetForegroundWindow(frmMain.hwnd)

#End If

    Do
        If (RestartActionList) Then
            Aborted = False
            SoftSkip = False
            AbortButton = False
            Paused = False
            Pausing = False
            PauseBetweenActions = False
            SkipToNextMoveAction = False
            RetryMoveAction = False
            SkipToNextSkipToAction = False
        End If
        
        RestartActionList = False
    
        Call SetupSubActions(colAction)
    
        If Aborted Then
            'major problem loading sub actions - get out of here
            GoTo StartActionError
        End If
    
        Set SkipAheadTimes = New Collection
        Set SkipAheadSoftSkips = New Collection
        
        Call SearchActionListForSkipActions(colAction)
        
        Call frmMain.StartTimer
        
        Call RunActionList(colAction, 1)
        
        'Clean up Completed, Skipped, Running messages
        Call CleanUpGUIActionList(colAction, 1)
    Loop While RestartActionList
        
    If Not Aborted Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.ActionListComplete).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander Action List Complete", "Action list complete." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
        
        Call AddToStatus("Action complete.")
    Else
#If Not InDebug Then
        If Not AbortButton Then
            'Must have aborted due to a failure!  Abort cameras and park mount!
            Call Camera.CameraAbort
            Call Mount.MountAbort
            
            'Aborted = False 'set to false so the ParkMount function will work
            'SoftSkip = False
            
            If frmOptions.chkEMailAlert(EMailAlertIndexes.GenericError).Value = vbChecked Then
                'Send e-mail!
                Call EMail.SendEMail(frmMain, "CCD Commander Action List Aborted", "Action list aborted due to an action failure." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
            End If
            
            GoTo StartActionError2
            
'            If frmOptions.chkParkMount = 1 Then
'                Call AddToStatus("Problem, attempting to park mount!")
'                Call ParkMount
'            Else
'                Call AddToStatus("Problem, stopping action.")
'            End If
        End If
#End If

        Call AddToStatus("Action stopped.")
    End If
    
    Close StatusFile
    StatusFile = 0
    Close MissedTargetFile
    MissedTargetFile = 0
    frmMain.txtCurrentAction.Text = ""
    
' Disabling watchdog for initial open source release
'    If Not (Watchdog Is Nothing) Then
'        Watchdog.StopWatchdog
'        Set Watchdog = Nothing
'    End If
    
    If (frmOptions.chkDisconnectAtEnd = vbChecked) And Not Mount.SimulatedPark Then
        On Error Resume Next
        Call Camera.CameraUnload
        Call Mount.MountUnload
        Call Focus.FocusUnload
        Call Rotator.RotatorUnload
        Call CloudSensor.CloudSensorUnload
        Call DomeControl.DomeUnload
        On Error GoTo 0
    End If
    
    Exit Sub
    
StartActionError:
    If (Err.Number <> 0) Then
        Call AddToStatus("Error Number: " & Err.Number)
        Call AddToStatus(Err.Description)
        Call AddToStatus(Err.Source)
        ErrorSource = Err.Source
        Err.Clear
    End If
    
    On Error GoTo 0
    
    Aborted = True
    
    myListIndex = 1
    For Each clsAction In colAction
        frmMain.lstAction(1).ListItems(myListIndex).Text = clsAction.BuildActionListString()
        myListIndex = myListIndex + 1
    Next clsAction
    
    Call frmMain.RemoveSubActionTabsAfter(1)
    
    'Clean up Completed, Skipped, Running messages
    Call CleanUpGUIActionList(colAction, 1)
    
    If frmOptions.chkEMailAlert(EMailAlertIndexes.GenericError).Value = vbChecked Then
        'Send e-mail!
        Call EMail.SendEMail(frmMain, "CCD Commander Action List Aborted", "Action list aborted due to a program error." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
    End If
    
StartActionError2:
    'Need to clear this so the park waits appropriately
    Aborted = False
    SoftSkip = False
    
    If frmOptions.chkParkMountFirst.Value = vbChecked Then
        If ErrorSource <> "TheSky6.Document" And frmOptions.chkParkMount = 1 Then
#If Not InDebug Then
            Call ParkMountOnError
#End If
        End If
    End If
    
    If frmOptions.chkCloseDomeOnError.Value = vbChecked And frmOptions.lstDomeControl.ListIndex <> DomeControlTypes.None Then
#If Not InDebug Then
        Call CloseDomeOnError
#End If
    End If
    
    If frmOptions.chkParkMountFirst.Value <> vbChecked Then
        If ErrorSource <> "TheSky6.Document" And frmOptions.chkParkMount = 1 Then
#If Not InDebug Then
            Call ParkMountOnError
#End If
        End If
    End If
    
    Call AddToStatus("Stopping.")
    
    Close StatusFile
    Close MissedTargetFile
    frmMain.txtCurrentAction.Text = ""
    
' Disabling watchdog for initial open source release
'    If Not (Watchdog Is Nothing) Then
'        Watchdog.StopWatchdog
'        Set Watchdog = Nothing
'    End If
    
End Sub

' This is in a subroutine so the error trapping can work
' Sometimes the Park doesn't work due to the time of the error and what may be going on
Private Sub ParkMountOnError()
    Call AddToStatus("Attempting to park mount!")
    
    On Error GoTo ParkOnError_Error
    Call ParkMount
ParkOnError_Error:
    If Err.Number <> 0 Then
        Call AddToStatus("Failed to park!")
    End If
    On Error GoTo 0
End Sub

' This is in a subroutine so the error trapping can work
' Sometimes this doesn't work due to the time of the error and what may be going on
Private Sub CloseDomeOnError()
    Call AddToStatus("Attempting to close dome!")
    
    On Error GoTo CloseDomeOnError_Error
    Call DomeControl.UnCoupleDome
    Call DomeControl.CloseDome
CloseDomeOnError_Error:
    If Err.Number <> 0 Then
        Call AddToStatus("Failed to close dome!")
    End If
    On Error GoTo 0
End Sub

Private Sub SetupSubActions(ActionCollection As Collection)
    Dim clsAction As Object
    Dim myListIndex As Long
    
    myListIndex = 0
    Do While Not Aborted And ActionCollection.Count > 0
        Set clsAction = ActionCollection.Item(myListIndex + 1)
        
        clsAction.RunTimeStatus = ""
        
        If TypeName(clsAction) = "RunActionList" Then
            If clsAction.LinkToFile = True Then
                Call MainMod.ClearCollection(clsAction.ActionCollection)
                
                Call LoadActionForRunActionList(clsAction.ActionListName, clsAction.ActionCollection)
            End If
            
            Call SetupSubActions(clsAction.ActionCollection)
        End If
    
        myListIndex = myListIndex + 1
        
        If Aborted Then Exit Do
                        
        If myListIndex >= ActionCollection.Count Then Exit Do
    Loop
End Sub

Public Sub SearchActionListForSkipActions(ActionCollection As Collection)
    Dim clsAction As Object
    Dim myListIndex As Long
    
    myListIndex = 0
    Do While Not Aborted And ActionCollection.Count > 0
        Set clsAction = ActionCollection.Item(myListIndex + 1)
        
        If clsAction.Selected Or Not RunSelectedActionsOnly Then
            Call AddSkipActionToSkipActionList(clsAction)
        End If
        
        myListIndex = myListIndex + 1
        
        If Aborted Then Exit Do
                        
        If myListIndex >= ActionCollection.Count Then Exit Do
    Loop
End Sub

Public Sub SearchAndDeleteSkipActions(ActionCollection As Collection)
    Dim clsAction As Object
    Dim myListIndex As Long
    
    myListIndex = 0
    Do While Not Aborted And ActionCollection.Count > 0
        Set clsAction = ActionCollection.Item(myListIndex + 1)
        
        If TypeName(clsAction) = "RunActionList" Then
            Call SearchAndDeleteSkipActions(clsAction.ActionCollection)
        ElseIf TypeName(clsAction) = "SkipAheadAtTimeAction" Or TypeName(clsAction) = "SkipAheadAtAltAction" Or TypeName(clsAction) = "SkipAheadAtHAAction" Then
            Call RemoveSkipToTime(clsAction.ActualSkipToTime)
        ElseIf TypeName(clsAction) = "AutoFlatAction" Then
            If clsAction.FlatLocation = DawnSkyFlat Or clsAction.FlatLocation = DuskSkyFlat Then
                Call RemoveSkipToTime(clsAction.ActualSkipToTime)
            End If
        End If
        
        myListIndex = myListIndex + 1
        
        If Aborted Then Exit Do
                        
        If myListIndex >= ActionCollection.Count Then Exit Do
    Loop
    
End Sub

Public Sub AddSkipActionToSkipActionList(clsAction As Object)
    Dim SkipTime As Date
    Dim tempTime As Date
    
    If TypeName(clsAction) = "RunActionList" Then
        Call SearchActionListForSkipActions(clsAction.ActionCollection)
    ElseIf TypeName(clsAction) = "SkipAheadAtTimeAction" Then
        If clsAction.Hour < 12 And CInt(Format(Time, "hh")) >= 12 Then
            SkipTime = (Format(Date + 1, "Short Date")) & " " & clsAction.Hour & ":" & clsAction.Minute & ":00"
        ElseIf clsAction.Hour >= 12 And CInt(Format(Time, "hh")) < 12 Then
            SkipTime = (Format(Date - 1, "Short Date")) & " " & clsAction.Hour & ":" & clsAction.Minute & ":00"
        Else
            SkipTime = Format(Date, "Short Date") & " " & clsAction.Hour & ":" & clsAction.Minute & ":00"
        End If
        
        clsAction.ActualSkipToTime = SkipTime
        
        Call AddSkipToTime(SkipTime, clsAction.SoftSkip)
    ElseIf TypeName(clsAction) = "SkipAheadAtAltAction" Then
        If clsAction.Name = "Sun" Then
            If clsAction.Rising Then
                Call Mount.ComputeTwilightStartTime(clsAction.Alt)
                
                If CInt(Format(Time, "hh")) > 12 And CInt(Format(Mount.TwilightStartTime, "hh")) < 12 Then
                    tempTime = Format(Date + 1, "Short Date") & " " & Mount.TwilightStartTime
                Else
                    tempTime = Format(Date, "Short Date") & " " & Mount.TwilightStartTime
                End If
            Else
                Call Mount.ComputeSunSetTime(clsAction.Alt)
                tempTime = Format(Date, "Short Date") & " " & Mount.SunSetTime
            End If
        ElseIf clsAction.Name = "Moon" Then
            If clsAction.Rising Then
                Call AstroFunctions.MoonRise(clsAction.Alt)
                tempTime = AstroFunctions.MoonRiseTime
            Else
                Call AstroFunctions.Moonset(clsAction.Alt)
                tempTime = AstroFunctions.MoonSetTime
            End If
        Else
            If clsAction.Rising Then
                tempTime = Misc.ComputeRiseTime(clsAction.RA, clsAction.Dec, clsAction.Alt, Mount.GetLatitude, Mount.GetLongitude)
            Else
                tempTime = Misc.ComputeSetTime(clsAction.RA, clsAction.Dec, clsAction.Alt, Mount.GetLatitude, Mount.GetLongitude)
            End If
        End If
        
        'Add 30s to time - this will round the time to the next minute
        tempTime = DateAdd("s", 30, tempTime)
        
        If Hour(tempTime) < 12 And CInt(Format(Time, "hh")) >= 12 Then
            SkipTime = (Format(DateAdd("d", 1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
        ElseIf Hour(tempTime) >= 12 And CInt(Format(Time, "hh")) < 12 Then
            SkipTime = (Format(DateAdd("d", -1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
        Else
            SkipTime = Format(Date, "Short Date") & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
        End If
        
        clsAction.ActualSkipToTime = SkipTime
        
        Call AddSkipToTime(SkipTime, clsAction.SoftSkip)
    ElseIf TypeName(clsAction) = "SkipAheadAtHAAction" Then
        tempTime = DateAdd("s", ((clsAction.HA + clsAction.RA) - Mount.GetSiderealTime()) * 3600, Now)
        
        'Add 30s to time - this will round the time to the next minute
        tempTime = DateAdd("s", 30, tempTime)
        
        If Hour(tempTime) < 12 And CInt(Format(Time, "hh")) >= 12 Then
            SkipTime = (Format(DateAdd("d", 1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
        ElseIf Hour(tempTime) >= 12 And CInt(Format(Time, "hh")) < 12 Then
            SkipTime = (Format(DateAdd("d", -1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
        Else
            SkipTime = Format(Date, "Short Date") & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
        End If
        
        clsAction.ActualSkipToTime = SkipTime
        
        Call AddSkipToTime(SkipTime, clsAction.SoftSkip)
    ElseIf TypeName(clsAction) = "AutoFlatAction" Then
        If clsAction.FlatLocation = DuskSkyFlat Then
            Call Mount.ComputeSunSetTime(clsAction.DuskSunAltitudeStart)
            
            tempTime = Format(Date, "Short Date") & " " & Mount.SunSetTime
            
            'Add 30s to time - this will round the time to the next minute
            tempTime = DateAdd("s", 30, tempTime)
            
            SkipTime = Format(tempTime, "Short Date") & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
            
            clsAction.ActualSkipToTime = SkipTime
            
            Call AddSkipToTime(SkipTime, False)
        ElseIf clsAction.FlatLocation = DawnSkyFlat Then
            Call Mount.ComputeTwilightStartTime(clsAction.DawnSunAltitudeStart)
            
            If CInt(Format(Time, "hh")) > 12 And CInt(Format(Mount.TwilightStartTime, "hh")) < 12 Then
                tempTime = Format(Date + 1, "Short Date") & " " & Mount.TwilightStartTime
            Else
                tempTime = Format(Date, "Short Date") & " " & Mount.TwilightStartTime
            End If
            
            'Add 30s to time - this will round the time to the next minute
            tempTime = DateAdd("s", 30, tempTime)
            
            SkipTime = Format(tempTime, "Short Date") & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
            
            clsAction.ActualSkipToTime = SkipTime
            
            Call AddSkipToTime(SkipTime, False)
        
        End If
    End If
End Sub

Public Sub CheckSkipToTimes()
    Dim mySkipAhead As Date
    Dim mySoftSkip As Boolean
    Dim Message As String
    
    If SkipAheadTimes.Count > 0 And SkipToNextSkipToAction = False Then
        mySkipAhead = SkipAheadTimes.Item(1)
        mySoftSkip = SkipAheadSoftSkips.Item(1)
        If DateDiff("s", Now, mySkipAhead) <= 0 Then
            Message = "Skip ahead at " & Format(mySkipAhead, "Short Time") & " action is active."
            If (mySoftSkip) Then
                Message = Message & " Soft skip will wait for a Take Image exposure to finish before skipping ahead."
            Else
                Message = Message & " Skipping ahead."
            End If
            
            Call AddToStatus(Message)
            If (mySoftSkip) Then
                SoftSkip = True
            End If
            Aborted = True
            SkipToNextSkipToAction = True
        End If
    End If
End Sub

Private Sub AddSkipToTime(SkipToTime As Date, SoftSkip As Boolean)
    Dim Counter As Integer
    
    If SkipAheadTimes.Count = 0 Then
        'just add it to the beginning
        Call SkipAheadTimes.Add(SkipToTime)
        Call SkipAheadSoftSkips.Add(SoftSkip)
        Exit Sub
    End If
    
    For Counter = 1 To SkipAheadTimes.Count
        If DateDiff("s", SkipAheadTimes.Item(Counter), SkipToTime) < 0 Then
            'add new time before the current one
            Call SkipAheadTimes.Add(SkipToTime, , Counter)
            Call SkipAheadSoftSkips.Add(SoftSkip, , Counter)
            Exit Sub
        End If
    Next Counter
    
    'add it to the end!
    Call SkipAheadTimes.Add(SkipToTime)
    Call SkipAheadSoftSkips.Add(SoftSkip)
End Sub

Public Function RemoveSkipToTime(SkipToTime As Date) As Integer
    Dim Counter As Integer
    
    For Counter = 1 To SkipAheadTimes.Count
        'Use minutes to compare - just in case seconds gives an inaccurate result
        If DateDiff("n", SkipToTime, SkipAheadTimes.Item(Counter)) = 0 Then
            Call SkipAheadTimes.Remove(Counter)
            Call SkipAheadSoftSkips.Remove(Counter)
            RemoveSkipToTime = Counter
            Exit Function
        End If
    Next Counter
    
    'Call AddToStatus("Error matching Skip Ahead Time to list.")
    RemoveSkipToTime = 0
End Function

Public Function RunActionList(ActionCollection As Collection, ByVal ActionListLevel As Integer) As Boolean
    Dim clsAction As Object
    Dim myListIndex As Long
    Dim CollectionIndex As Long
    Dim Temp As Integer
    Dim ReadyToRun As Boolean
    Dim FormList As ListView
    Dim RanOneAction As Boolean
    RanOneAction = False
    
    Set FormList = frmMain.lstAction(ActionListLevel)
    
    RunningActionListLevel = ActionListLevel
    
    myListIndex = 1
    Do While Not Aborted And ActionCollection.Count > 0
        ReadyToRun = False
        Do
            Set clsAction = ActionCollection.Item(myListIndex)
            
            If frmMain.EditingActionNumber = myListIndex And frmMain.EditingActionListLevel = ActionListLevel Then
                Call AddToStatus("Waiting for user to close the open action.")
            End If
            
            Do While frmMain.EditingActionNumber = myListIndex And frmMain.EditingActionListLevel = ActionListLevel And Not Aborted
                Call Wait(1)
            Loop
            
            If Not Aborted Then
                If (Not ActionCollection.Item(myListIndex).Selected) And RunSelectedActionsOnly Then
                    On Error Resume Next
                    'This could fail if the user closed the sub-action window - just ignore it if it does
                    If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
                        If Err.Number = 0 Then
                            FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString & " - Skipped"
                        End If
                    End If
                    On Error GoTo 0
                    ActionCollection.Item(myListIndex).RunTimeStatus = "Skipped"
                    
                    myListIndex = myListIndex + 1
                    
                    If myListIndex > ActionCollection.Count Then Exit Do
                    
                    ReadyToRun = False
                Else
                    On Error Resume Next
                    'This could fail if the user closed the sub-action window - just ignore it if it does
                    If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
                        If Err.Number = 0 Then
                            FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString & " - Running"
                        End If
                    End If
                    On Error GoTo 0
                    ActionCollection.Item(myListIndex).RunTimeStatus = "Running"
                    ReadyToRun = True
                End If
            End If
        Loop Until ReadyToRun Or Aborted
                
        If myListIndex > ActionCollection.Count Or Aborted Then Exit Do
            
        RanOneAction = True
        
        frmMain.txtCurrentAction.Text = clsAction.BuildActionListString()
            
        If TypeName(clsAction) = "ImagerAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call ImagerAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "MoveRADecAction" Then
            If SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                SkipToNextMoveAction = False
                RetryMoveAction = False
                Call MoveRADecAction(clsAction)
                If (RetryMoveAction) Then
                    Call AddToStatus("Retrying Move in 5s...")
                    Call Wait(5)
                    Call MoveRADecAction(clsAction)
                End If
            End If
        ElseIf TypeName(clsAction) = "FocusAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call FocusMaxAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "ImageLinkSyncAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                If Not TakeImageAndLink(clsAction) Then
                    If clsAction.AbortListOnFailure Then
                        Call AddToStatus("Aborting remainder of list/sub-list.")
                        If ActionListLevel > 1 Then
                            RunActionList = False
                            Exit Function
                        Else
                            Aborted = True
                        End If
                    End If
                End If
            End If
        ElseIf TypeName(clsAction) = "ParkMountAction" Then
            If SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                SkipToNextMoveAction = False
                Call ParkMount(clsAction)
            End If
        ElseIf TypeName(clsAction) = "DomeAction" Then
            If SkipToNextSkipToAction Or SkipToNextMoveAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call DomeControl.DomeAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "WaitForAltAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call WaitForAltAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "WaitForTimeAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call WaitForTimeAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "RunScriptAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call RunScript(clsAction)
            End If
        ElseIf TypeName(clsAction) = "RunActionList" Then
            If FollowRunningAction Then
                Call frmMain.RemoveSubActionTabsAfter(ActionListLevel)
            
                'Copy link to the ActionInfo object into the array so I can access the object when the user makes changes on the window
                Call frmMain.SubActionInfo.Add(clsAction, , , ActionListLevel)
                
                'Copy link to the Collection into the collection array
                Call frmMain.ActionCollections.Add(clsAction.ActionCollection, , , ActionListLevel)
                    
                Call frmMain.UpperActionIndexes.Add(myListIndex, , , ActionListLevel)
                
                If Not frmMain.AllowSubActionEdit(ActionListLevel) Or clsAction.LinkToFile Then
                    Call frmMain.AllowSubActionEdit.Add(False, , , ActionListLevel)
                Else
                    Call frmMain.AllowSubActionEdit.Add(True, , , ActionListLevel)
                End If
            End If
            
            Call ActionList.RunActionListAction(clsAction, ActionListLevel + 1)
            RunningActionListLevel = ActionListLevel
            
            If FollowRunningAction Then
                frmMain.ActionCollections.Remove ActionListLevel + 1
                frmMain.SubActionInfo.Remove ActionListLevel + 1
                frmMain.UpperActionIndexes.Remove ActionListLevel + 1
                frmMain.AllowSubActionEdit.Remove ActionListLevel + 1
            End If
        ElseIf TypeName(clsAction) = "RotatorAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call Rotator.RotateAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "IntelligentTempAction" Then
            If SkipToNextMoveAction Or SkipToNextSkipToAction Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call TempControl.IntelligentTempControl(clsAction)
            End If
        ElseIf TypeName(clsAction) = "CloudMonitorAction" Then
            Call CloudSensor.CloudMonitorAction(clsAction)
        ElseIf TypeName(clsAction) = "AutoFlatAction" Then
            If clsAction.FlatLocation = DuskSkyFlat Or clsAction.FlatLocation = DawnSkyFlat Then
                Temp = RemoveSkipToTime(clsAction.ActualSkipToTime)
            Else
                Temp = 1
            End If
            
            If Temp > 1 And SkipToNextSkipToAction = True Then
                Call AddToStatus("Skipping Auto Flat action - another action generated the skip!")
                Call AddToStatus("Auto Flat skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
            Else
                SkipToNextMoveAction = False
                SkipToNextSkipToAction = False
                
                'Could be paused due to clouds
                'Wait here until pause is false or we get another aborted signal
                If PauseBetweenActions = True And AbortButton = False And SkipToNextSkipToAction = False Then
                    Aborted = False
                    SoftSkip = False
                    
                    Call AddToStatus("Action list paused...")
                    Do While PauseBetweenActions = True And Not Aborted
                        Call Wait(1)
                    Loop
                    Call AddToStatus("Resuming action list.")
                End If
                
                If Not Aborted Then _
                    Call Camera.RunAutoFlatAction(clsAction)
            End If
        ElseIf TypeName(clsAction) = "SkipAheadAtTimeAction" Then
            frmMain.DisableSkipChecking = True
            
            If SkipToNextSkipToAction = False Then
                Call AddToStatus("Ignoring Skip To action.  Action reached before skip time.")
                Call AddToStatus("Skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
                Call RemoveSkipToTime(clsAction.ActualSkipToTime)
            Else
                If RemoveSkipToTime(clsAction.ActualSkipToTime) > 1 Then
                    Call AddToStatus("Skipping Skip To action - another action generated the skip!")
                    Call AddToStatus("Skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
                Else
                    Call AddToStatus("At Skip Ahead At Time Action.  Skip Time was " & Format(clsAction.ActualSkipToTime, "Short Time"))
                    SkipToNextSkipToAction = False
                End If
            End If
            
            frmMain.DisableSkipChecking = False
        ElseIf TypeName(clsAction) = "SkipAheadAtAltAction" Then
            frmMain.DisableSkipChecking = True
        
            If SkipToNextSkipToAction = False Then
                Call AddToStatus("Ignoring Skip To action.  Action reached before skip time.")
                Call AddToStatus("Skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
                Call RemoveSkipToTime(clsAction.ActualSkipToTime)
            Else
                If RemoveSkipToTime(clsAction.ActualSkipToTime) > 1 Then
                    Call AddToStatus("Skipping Skip To action - another action generated the skip!")
                    Call AddToStatus("Skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
                Else
                    If clsAction.Name = "" Then
                        Call AddToStatus("Coordinates " & Misc.ConvertEquatorialToString(clsAction.RA, clsAction.Dec, True) & " passed " & clsAction.Alt & "d altitude.")
                    Else
                        Call AddToStatus(clsAction.Name & " passed " & clsAction.Alt & "d altitude.")
                    End If
                    SkipToNextSkipToAction = False
                End If
            End If
            
            frmMain.DisableSkipChecking = False
        ElseIf TypeName(clsAction) = "SkipAheadAtHAAction" Then
            frmMain.DisableSkipChecking = True
            
            If SkipToNextSkipToAction = False Then
                Call AddToStatus("Ignoring Skip To action.  Action reached before skip time.")
                Call AddToStatus("Skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
                Call RemoveSkipToTime(clsAction.ActualSkipToTime)
            Else
                If RemoveSkipToTime(clsAction.ActualSkipToTime) > 1 Then
                    Call AddToStatus("Skipping Skip To action - another action generated the skip!")
                    Call AddToStatus("Skip time was " & Format(clsAction.ActualSkipToTime, "Short Time") & ".")
                Else
                    If clsAction.Name = "" Then
                        Call AddToStatus("Coordinates " & Misc.ConvertRAToString(clsAction.RA, True) & " passed hour angle " & clsAction.HA & "h.")
                    Else
                        Call AddToStatus(clsAction.Name & " passed hour angle " & clsAction.HA & "h.")
                    End If
                    SkipToNextSkipToAction = False
                End If
            End If
            
            frmMain.DisableSkipChecking = False
        ElseIf TypeName(clsAction) = "clsCommentAction" Then
            If (SkipToNextMoveAction Or SkipToNextSkipToAction) And (Not clsAction.DoNotSkip) Then
                Call AddToStatus("Skipping Action: " & clsAction.BuildActionListString())
            Else
                Call Comment.RunCommentAction(clsAction)
            End If
        End If
        
        If Aborted = True And AbortButton = False And SkipToNextSkipToAction = True Then
            Call AddToStatus("Skipping remainder of Action: " & clsAction.BuildActionListString())
            
            Aborted = False
            SoftSkip = False
            Call MountAbort
            Call CameraAbort
            Call Wait(1)
        End If
        
        If PauseBetweenActions = True And AbortButton = False And SkipToNextSkipToAction = False Then
            Aborted = False
            SoftSkip = False
            
            Call AddToStatus("Action list paused...")
            
            Do While PauseBetweenActions = True And Not Aborted
                Call Wait(1)
            Loop
            
            If Not Aborted Then
                Call AddToStatus("Resuming action list.")
            End If
            
            'Got out of above loop due to a skip to action becoming active
            'Need to set aborted to false so I can continue the action list
            If SkipToNextSkipToAction = True Then
                If (TypeName(clsAction) <> "SkipAheadAtTimeAction") And (TypeName(clsAction) <> "SkipAheadAtAltAction") And (TypeName(clsAction) <> "SkipAheadAtHAAction") And (TypeName(clsAction) <> "CloudMonitorAction") Then
                    Call AddToStatus("Skipping remainder of Action: " & clsAction.BuildActionListString())
                End If
                
                Aborted = False
                SoftSkip = False
                On Error Resume Next
                'This could fail if the user closed the sub-action window - just ignore it if it does
                If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
                    If Err.Number = 0 Then
                        FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString & " - Skipped"
                    End If
                End If
                On Error GoTo 0
                clsAction.RunTimeStatus = "Skipped"
                
                myListIndex = myListIndex + 1
            ElseIf Not Aborted Then
                'Go to the next action only if it is not a wait for action - those we will try again
                If (TypeName(clsAction) <> "WaitForAltAction") And (TypeName(clsAction) <> "WaitForTimeAction") And _
                    (TypeName(clsAction) <> "AutoFlatAction") And (TypeName(clsAction) <> "CloudMonitorAction") And _
                    (frmOptions.chkWeatherMonitorRepeatAction.Value <> vbChecked) Then
                    
                    Call AddToStatus("Skipping remainder of Action: " & clsAction.BuildActionListString())
                    
                    On Error Resume Next
                    'This could fail if the user closed the sub-action window - just ignore it if it does
                    If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
                        If Err.Number = 0 Then
                            FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString & " - Skipped"
                        End If
                    End If
                    On Error GoTo 0
                    clsAction.RunTimeStatus = "Skipped"
                    
                    myListIndex = myListIndex + 1
                End If
            End If
        Else
            If SkipToNextMoveAction Or SkipToNextSkipToAction And TypeName(clsAction) <> "CloudMonitorAction" Then
                On Error Resume Next
                'This could fail if the user closed the sub-action window - just ignore it if it does
                If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
                    If Err.Number = 0 Then
                        FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString & " - Skipped"
                    End If
                End If
                On Error GoTo 0
                clsAction.RunTimeStatus = "Skipped"
            Else
                On Error Resume Next
                'This could fail if the user closed the sub-action window - just ignore it if it does
                If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
                    If Err.Number = 0 Then
                        FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString & " - Complete"
                    End If
                End If
                On Error GoTo 0
                clsAction.RunTimeStatus = "Complete"
            End If
                    
            myListIndex = myListIndex + 1
        End If
                
        If Aborted Then Exit Do
                        
        If myListIndex > ActionCollection.Count Then Exit Do
    Loop

'    If Not (FormList Is Nothing) Then
'        myListIndex = 1
'        For Each clsAction In ActionCollection
'            clsAction.RunTimeStatus = ""
'            On Error Resume Next
'            If frmMain.ActionCollections(ActionListLevel) Is ActionCollection Then
'                FormList.ListItems(myListIndex).Text = clsAction.BuildActionListString()
'            End If
'            On Error GoTo 0
'            myListIndex = myListIndex + 1
'        Next clsAction
'    End If
    
    RunActionList = RanOneAction
End Function

Public Sub CleanUpGUIActionList(ActionCollection As Collection, ActionLevel As Integer)
    Dim clsAction As Object
    Dim ListIndex As Integer
    
    On Error Resume Next
    If (frmMain.ActionCollections(ActionLevel) Is ActionCollection And ActionLevel > 1) And (frmMain.ActionCollections.Item(ActionLevel - 1).Item(frmMain.UpperActionIndexes.Item(ActionLevel)).RunTimeStatus <> "Running") Then
        If Err.Number = 0 Then
            frmMain.optRunAborted(ActionLevel).Enabled = True
            frmMain.optRunMultiple(ActionLevel).Enabled = True
            frmMain.optRunOnce(ActionLevel).Enabled = True
            frmMain.optRunPeriod(ActionLevel).Enabled = True
        End If
    End If
    On Error GoTo 0
    
    On Error GoTo CleanUpGUIActionListError
    
    ListIndex = 1
    For Each clsAction In ActionCollection
        clsAction.RunTimeStatus = ""
        If TypeName(clsAction) = "RunActionList" Then
            Call CleanUpGUIActionList(clsAction.ActionCollection, ActionLevel + 1)
        End If
        
        On Error Resume Next
        If frmMain.ActionCollections(ActionLevel) Is ActionCollection Then
            If Err.Number = 0 Then
                frmMain.lstAction(ActionLevel).ListItems(ListIndex).Text = clsAction.BuildActionListString
            End If
        End If
        On Error GoTo 0
        
        ListIndex = ListIndex + 1
    Next clsAction
    
CleanUpGUIActionListError:
    ' Just get out of here...
    On Error GoTo 0
    
End Sub

Public Sub JumpToRunningAction(ActionCollection As Collection, ActionListLevel As Integer)
    Dim clsAction As Object
    Dim myListIndex As Integer
    
    'Now start searching for "Running" status texts
    'stop when I find a running action that is not a sub-action
    myListIndex = 1
    For Each clsAction In ActionCollection
        If clsAction.RunTimeStatus = "Running" Then
            'check if it is a sub-action
            If TypeName(clsAction) = "RunActionList" Then
                Call frmMain.RemoveSubActionTabsAfter(ActionListLevel)
            
                'Copy link to the ActionInfo object into the array so I can access the object when the user makes changes on the window
                Call frmMain.SubActionInfo.Add(clsAction, , , ActionListLevel)
                
                'Copy link to the Collection into the collection array
                Call frmMain.ActionCollections.Add(clsAction.ActionCollection, , , ActionListLevel)
                    
                Call frmMain.UpperActionIndexes.Add(myListIndex, , , ActionListLevel)
                
                If Not frmMain.AllowSubActionEdit(ActionListLevel) Then
                    Call frmMain.AllowSubActionEdit.Add(False, , , ActionListLevel)
                Else
                    Call frmMain.AllowSubActionEdit.Add(True, , , ActionListLevel)
                End If
                
                Call frmMain.AddStuffForSubAction(clsAction)
                frmMain.optRunAborted(ActionListLevel + 1).Enabled = False
                frmMain.optRunMultiple(ActionListLevel + 1).Enabled = False
                frmMain.optRunOnce(ActionListLevel + 1).Enabled = False
                frmMain.optRunPeriod(ActionListLevel + 1).Enabled = False
                
                Call JumpToRunningAction(clsAction.ActionCollection, ActionListLevel + 1)
            End If
        End If
        
        myListIndex = myListIndex + 1
    Next clsAction
End Sub

Public Sub AbortAction()
    Call AddToStatus("User aborted! Stopping imager and autoguider...")
    Aborted = True
    AbortButton = True
    Call MountAbort
    Call CameraAbort
End Sub

Public Sub SetupStuff()
    Randomize Timer
    Set colAutoguiderErrors = New Collection
    
    Load frmTempGraph
    Load frmAutoguiderError
    Load frmOptions

    Call Camera.PutFilterDataIntoForms
        
    Rotator.RotatorConnected = False
    
    On Error Resume Next
    Call MkDir(App.Path & "\Logs")
    On Error GoTo 0
    
    On Error Resume Next
    Call MkDir(App.Path & "\Images")
    On Error GoTo 0
    
    Set SkipAheadTimes = New Collection
    Set SkipAheadSoftSkips = New Collection
End Sub

Public Sub UnloadStuff()
    Call FocusUnload
    Call CameraUnload
    Call MountUnload
    Call RotatorUnload
    
    Close StatusFile
    
    Unload frmOptions
    Unload frmAutoguiderError
    Unload frmTempGraph
End Sub

Public Function Wait(TimeInSec As Double)
    Dim StartTime As Double
    Dim StartDate As Date
    Dim TimeNow As Double
    
    StartTime = Timer
    StartDate = Date
    
    Do
        If Paused Then Pausing = True
        Do
' Disabling watchdog for initial open source release
'            If Not (Watchdog Is Nothing) Then
'                Watchdog.ResetWatchdogTimer
'            End If
            
            Call Sleep(50)
            DoEvents
        Loop While Paused
        Pausing = False
        
        On Error Resume Next
        Do
            Err.Clear
            TimeNow = CDbl(Timer)
        Loop While Err.Number <> 0
        On Error GoTo 0
        
        If GetMySetting("Test", "Aborted", "0") = "1" Then
            Call AddToStatus("Remote abort!!!")
            Aborted = True
            AbortButton = True
        End If
                        
    Loop Until TimeNow > (StartTime + TimeInSec) Or _
        (DateDiff("d", Date, StartDate) <> 0 And _
        (TimeNow + 86400) > (StartTime + TimeInSec)) _
        Or Aborted
    
    If Timer > (StartTime + TimeInSec) Then
        Wait = True
    Else
        Wait = False
    End If
End Function

Public Sub AddToStatus(Message As String)
    Message = Format(Now, "hh:mm:ss") & "  " & Message
    If StatusFile <> 0 And Not Exiting Then
        Print #StatusFile, Message
    End If
    
    If (MoveToRADecPlateSolveStatus) Then
        frmMoveToPlateSolve.txtStatus.Text = frmMoveToPlateSolve.txtStatus.Text & Message & vbCrLf
        frmMoveToPlateSolve.txtStatus.SelStart = Len(frmMoveToPlateSolve.txtStatus.Text)
        frmMoveToPlateSolve.txtStatus.SelLength = 0
    Else
        frmMain.txtStatus.Text = frmMain.txtStatus.Text & Message & vbCrLf
        frmMain.txtStatus.SelStart = Len(frmMain.txtStatus.Text)
        frmMain.txtStatus.SelLength = 0
    End If
    DoEvents
End Sub

Public Sub AddToMissedTargetList(TargetName As String, RA As Double, Dec As Double)
    If MissedTargetFile <> 0 And Not Exiting Then
        Print #MissedTargetFile, TargetName & ", ", RA & ", ", Dec
    End If
End Sub

Private Sub WriteHeaderToLogFile()
    Print #StatusFile, "CCD Commander Log File"
    Print #StatusFile, "Software v" & App.Major & "." & App.Minor & "." & App.Revision
    
    If Registration.Registered Then
        Print #StatusFile, "This copy is registered to:"
        Print #StatusFile, Registration.RegName
        Print #StatusFile, Registration.RegEMail
    Else
        Print #StatusFile, "This copy is unregistered."
        Print #StatusFile, "Trial period expires in " & Registration.TrialDaysLeft & " days."
    End If
    
    Print #StatusFile,
End Sub

Public Function GetMySetting(Section As String, Key As String, Optional Default As String = "") As String
    GetMySetting = GetSetting("CCDCommander", Section, Key, Default)
End Function

Public Sub SaveMySetting(Section As String, Key As String, Value As String)
    Call SaveSetting("CCDCommander", Section, Key, Value)
End Sub

Public Sub DeleteMySetting(Section As String)
    On Error Resume Next
    Call DeleteSetting("CCDCommander", Section)
    On Error GoTo 0
End Sub

Public Function GetActionFromName(strTypeName As String) As Object
    Dim myclsAction As Object
    
    If strTypeName = "ImagerAction" Then
        Set myclsAction = New ImagerAction
    ElseIf strTypeName = "MoveRADecAction" Then
        Set myclsAction = New MoveRADecAction
    ElseIf strTypeName = "FocusAction" Then
        Set myclsAction = New FocusAction
    ElseIf strTypeName = "ImageLinkSyncAction" Then
        Set myclsAction = New ImageLinkSyncAction
    ElseIf strTypeName = "ParkMountAction" Then
        Set myclsAction = New ParkMountAction
    ElseIf strTypeName = "WaitForAltAction" Then
        Set myclsAction = New WaitForAltAction
    ElseIf strTypeName = "WaitForTimeAction" Then
        Set myclsAction = New WaitForTimeAction
    ElseIf strTypeName = "RunScriptAction" Then
        Set myclsAction = New RunScriptAction
    ElseIf strTypeName = "RunActionList" Then
        Set myclsAction = New RunActionList
    ElseIf strTypeName = "RotatorAction" Then
        Set myclsAction = New RotatorAction
    ElseIf strTypeName = "IntelligentTempAction" Then
        Set myclsAction = New IntelligentTempAction
    ElseIf strTypeName = "AutoFlatAction" Then
        Set myclsAction = New AutoFlatAction
    ElseIf strTypeName = "SkipAheadAtTimeAction" Then
        Set myclsAction = New SkipAheadAtTimeAction
    ElseIf strTypeName = "SkipAheadAtAltAction" Then
        Set myclsAction = New SkipAheadAtAltAction
    ElseIf strTypeName = "SkipAheadAtHAAction" Then
        Set myclsAction = New SkipAheadAtHAAction
    ElseIf strTypeName = "DomeAction" Then
        Set myclsAction = New DomeAction
    ElseIf strTypeName = "CloudMonitorAction" Then
        Set myclsAction = New CloudMonitorAction
    ElseIf strTypeName = "clsCommentAction" Then
        Set myclsAction = New clsCommentAction
    Else
        Call MsgBox("Load Error!!!", vbCritical)
        Stop
    End If
    
    Set GetActionFromName = myclsAction
End Function

Public Sub LoadActionForRunActionList(FileName As String, ActionListCollection As Collection)
    Dim actFileNum As Integer
    Dim ByteData() As Byte
    Dim NumberOfBytes As Long

    On Error GoTo LoadActionForRunActionListError
    actFileNum = FreeFile
    Open FileName For Binary Access Read As #actFileNum

    NumberOfBytes = LOF(actFileNum)
    
    ReDim ByteData(0 To NumberOfBytes - 1)
    Get #actFileNum, , ByteData()
    Close actFileNum
    
    Call LoadActionLists(ActionListCollection, ByteData, 0, UBound(ByteData))
    
    On Error GoTo 0
    Exit Sub
    
LoadActionForRunActionListError:
    On Error GoTo 0
    Call AddToStatus("Error!  Cannot load sub action " & FileName & "!")
    Aborted = True
End Sub

Public Sub LoadAction(FileName As String)
    Dim actFileNum As Integer
    Dim ByteData() As Byte
    Dim NumberOfBytes As Long
    Dim ErrNumber As Integer
    Dim ErrSource As Variant
    Dim ErrDesc As Variant
    Dim ErrHelp As Variant
    Dim ErrHelpContext As Variant
    
    On Error GoTo LoadActionError
        
    'First try to open the file for "input" mode, this will throw an error if the file doesn't exist
    'I'm doing it this way since the open for binary will create the file if it doesn't exist
    actFileNum = FreeFile
    Open FileName For Input Access Read As #actFileNum
    'Open succeeded, close the file
    Close actFileNum
    
    'Now I can open it for real
    actFileNum = FreeFile
    Open FileName For Binary Access Read As #actFileNum
        
    Call MainMod.ClearAll
    
    NumberOfBytes = LOF(actFileNum)
    ReDim ByteData(0 To NumberOfBytes - 1)
    Get #actFileNum, , ByteData()
    Close actFileNum
    
    Call LoadActionLists(colAction, ByteData, 0, UBound(ByteData), , frmMain.lstAction(1))
    
    Call AddFileToPreviousFileList(FileName)
    
LoadActionError:
    ErrNumber = Err.Number
    ErrSource = Err.Source
    ErrDesc = Err.Description
    ErrHelp = Err.HelpFile
    ErrHelpContext = Err.HelpContext
    
    On Error GoTo 0
    
    If ErrNumber <> 0 Then
        'Ensure the file is closed
        Close actFileNum
        'toss the error back up to the calling function
        Call Err.Raise(ErrNumber, ErrSource, ErrDesc, ErrHelp, ErrHelpContext)
    End If
End Sub

Public Sub LoadActionLists(ActionCollection As Collection, ByRef ByteData() As Byte, ByRef Index As Long, ByVal StopIndex As Long, Optional ByVal ListIndex As Long = -1, Optional ByRef ListToAddTo As ListView = Nothing, Optional ByVal AddAfter As Boolean = False)
    Dim clsAction As Object
    Dim strTypeName As String
    Dim ReadLength As Long
    
    Do Until Index > StopIndex
        DoEvents
        Call CopyMemory(ReadLength, ByteData(Index), Len(ReadLength))
        Index = Index + Len(ReadLength)
        strTypeName = String(ReadLength, " ")
        Call CopyByteArrayToString(strTypeName, ByteData(), Index)  'Index is passed ByRef, so I don't need to increment it
        
        Set clsAction = GetActionFromName(strTypeName)
        Call clsAction.LoadActionByteArray(ByteData(), Index)
        
        If TypeName(clsAction) = "RunActionList" Then
            Set clsAction.ActionCollection = New Collection
            
            If Not clsAction.LinkToFile Then
                Call CopyMemory(ReadLength, ByteData(Index), Len(ReadLength))
                Index = Index + Len(ReadLength)
                
                Call LoadActionLists(clsAction.ActionCollection, ByteData, Index, Index + ReadLength - 1)
            End If
        End If
        
        If ListIndex > 0 Then
            If AddAfter Then
                Call ActionCollection.Add(clsAction, , , ListIndex)
            Else
                Call ActionCollection.Add(clsAction, , ListIndex)
            End If
        Else
            Call ActionCollection.Add(clsAction)
        End If
        
        If Not (ListToAddTo Is Nothing) Then
            If ListIndex >= 0 Then
                If AddAfter Then
                    Call ListToAddTo.ListItems.Add(ListIndex + 1, , clsAction.BuildActionListString)
                    ListToAddTo.ListItems(ListIndex + 1).Checked = clsAction.Selected
                Else
                    Call ListToAddTo.ListItems.Add(ListIndex, , clsAction.BuildActionListString)
                    ListToAddTo.ListItems(ListIndex).Checked = clsAction.Selected
                End If
            Else
                Call ListToAddTo.ListItems.Add(, , clsAction.BuildActionListString)
                ListToAddTo.ListItems(ListToAddTo.ListItems.Count).Checked = clsAction.Selected
            End If
        End If
        
        If ListIndex >= 0 Then
            ListIndex = ListIndex + 1
        End If
    Loop
End Sub

Public Function GetNumberOfBytesForActionList(ActionCollection As Collection) As Long
    Dim NumberOfBytes As Long
    Dim clsAction As Object
    Dim strTypeName As String
    Dim strLen As Long
    
    NumberOfBytes = 0
    For Each clsAction In ActionCollection
        strTypeName = TypeName(clsAction)
        strLen = Len(strTypeName)
        NumberOfBytes = NumberOfBytes + clsAction.ByteArraySize() + Len(strLen) + Len(strTypeName)
        
        If TypeName(clsAction) = "RunActionList" Then
            If Not clsAction.LinkToFile Then
                'not linked to another file, need to save the action list
                'first increment by the length of a long
                'this long will be equal to the length of the sub-action
                NumberOfBytes = NumberOfBytes + Len(NumberOfBytes)
                NumberOfBytes = NumberOfBytes + GetNumberOfBytesForActionList(clsAction.ActionCollection)
            End If
        End If
    Next clsAction
    
    GetNumberOfBytesForActionList = NumberOfBytes
End Function

Public Sub SaveActionLists(ActionCollection As Collection, ByRef ByteData() As Byte, ByRef ByteIndex As Long)
    Dim clsAction As Object
    Dim strTypeName As String
    Dim strLen As Long
    
    For Each clsAction In ActionCollection
        DoEvents
        strTypeName = TypeName(clsAction)
        strLen = Len(strTypeName)
        Call CopyMemory(ByteData(ByteIndex), strLen, Len(strLen))
        ByteIndex = ByteIndex + Len(strLen)
        
        Call CopyStringToByteArray(ByteData(), ByteIndex, strTypeName)  'NumberOfBytes is passed ByRef, so I don't need to increment it
        
        Call clsAction.SaveActionByteArray(ByteData(), ByteIndex)       'NumberOfBytes is passed ByRef, so I don't need to increment it
        
        If TypeName(clsAction) = "RunActionList" Then
            If Not clsAction.LinkToFile Then
                'not linked to another file, need to save the action list
                'First save the length of the sub-action list
                strLen = GetNumberOfBytesForActionList(clsAction.ActionCollection)
                Call CopyMemory(ByteData(ByteIndex), strLen, Len(strLen))
                ByteIndex = ByteIndex + Len(strLen)
                'now recurse the function to save the sub-action list
                Call SaveActionLists(clsAction.ActionCollection, ByteData, ByteIndex)
            End If
        End If
    Next clsAction
End Sub

Public Function SaveAction(FileName As String, ActionCollection As Collection) As Boolean
    Dim clsAction As Object
    Dim actFileNum As Long
    Dim ByteData() As Byte
    Dim NumberOfBytes As Long
    
    On Error GoTo SaveActionErrorHandler
    
    If Dir(FileName) <> "" Then
        If Dir(Mid(FileName, 1, Len(FileName) - 4) & ".bak") <> "" Then
            Kill Mid(FileName, 1, Len(FileName) - 4) & ".bak"
        End If
        
        Name FileName As Mid(FileName, 1, Len(FileName) - 4) & ".bak"
    End If
    
    actFileNum = FreeFile
    
    Open FileName For Binary Access Write Lock Read Write As #actFileNum

    NumberOfBytes = GetNumberOfBytesForActionList(ActionCollection)
    
    If NumberOfBytes > 0 Then
        ReDim ByteData(0 To NumberOfBytes - 1)
        
        NumberOfBytes = 0
        Call SaveActionLists(ActionCollection, ByteData, NumberOfBytes)
        Put #actFileNum, , ByteData()
    End If
        
    Close actFileNum
        
    Call AddFileToPreviousFileList(FileName)
    
    On Error GoTo 0
    SaveAction = True
    Exit Function
    
SaveActionErrorHandler:
    On Error GoTo 0
    MsgBox "Error Saving the Action.", vbCritical
    
    Close actFileNum
    
    If Dir(FileName) = "" And Dir(Mid(FileName, 1, Len(FileName) - 4) & ".bak") <> "" Then
        Name Mid(FileName, 1, Len(FileName) - 4) & ".bak" As FileName
    End If
    
    SaveAction = False
End Function

Public Sub CopyStringToByteArray(ByRef ptrByteArray() As Byte, Index As Long, SourceString As String)
    Dim strIndex As Long
    
    For strIndex = 1 To Len(SourceString)
        ptrByteArray(Index) = AscB(Mid(SourceString, strIndex, 1))
        Index = Index + 1
    Next strIndex
End Sub

Public Sub CopyByteArrayToString(DestinationString As String, ByRef ptrByteArray() As Byte, Index As Long)
    Dim strIndex As Long
    
    For strIndex = 1 To Len(DestinationString)
        Mid(DestinationString, strIndex, 1) = Chr(ptrByteArray(Index))
        Index = Index + 1
    Next strIndex
End Sub

Public Sub ClearCollection(myCollection As Collection)
    Do While myCollection.Count > 0
        If TypeName(myCollection.Item(1)) = "RunActionList" Then
            Call ClearSubActions(myCollection.Item(1))
        End If
        Call myCollection.Remove(1)
    Loop
End Sub

Public Sub ClearSubActions(myAction As RunActionList)
    Dim clsAction As Object
    
    For Each clsAction In myAction.ActionCollection
        If TypeName(clsAction) = "RunActionList" Then
            Call ClearSubActions(clsAction)
        End If
    Next clsAction
    
    Set myAction.ActionCollection = Nothing
End Sub

Public Sub ClearAll()
    'There was a bug where the Sub-Action List "name" was not being validated until after things were cleared below.
    'The validation routine would then generate an error.
    'This causes the tab strip to get focus, firing off the validation, before things are cleared out below.  Seems to fix the problem.
    frmMain.TabStrip.Tabs.Item(1).Selected = True
    
    Call MainMod.ClearCollection(colAction)
    Call frmMain.lstAction(1).ListItems.Clear
    Call frmMain.RemoveSubActionTabsAfter(1)
End Sub

Public Sub SetOnTopMode(thisForm As Form)
    'thisForm.Visible = False
    thisForm.ScaleMode = 3
    
    If frmMain.mnuWindowItem(0).Checked = True Then
        SetWindowPos thisForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos thisForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    'thisForm.Visible = False
End Sub

Public Sub AddFileToPreviousFileList(FileName As String)
    'First save the file name to the registry
    'First check if the file name already exists, then I just move it to the top
    Dim i As Integer
    Dim j As Integer
    
    For i = 5 To 1 Step -1
        If GetMySetting("PreviousFiles", "PreviousFile" & i, "") = FileName Then
            Exit For
        End If
    Next i
    
    If i = 1 Then
        'i don't need to do anything
        'this file is already at the top of the list
    ElseIf i > 1 Then
        'Move down the others in the list
        For j = i To 2 Step -1
            Call SaveMySetting("PreviousFiles", "PreviousFile" & j, GetMySetting("PreviousFiles", "PreviousFile" & j - 1, ""))
        Next j
    Else
        'Move down the others in the list
        For j = 5 To 2 Step -1
            Call SaveMySetting("PreviousFiles", "PreviousFile" & j, GetMySetting("PreviousFiles", "PreviousFile" & j - 1, ""))
        Next j
    End If

    Call SaveMySetting("PreviousFiles", "PreviousFile1", FileName)
    
    'Now update the menu
    Call frmMain.UpdatePreviousFileMenu
End Sub

Public Sub RemoveFileFromPreviousFileList(Index As Integer)
    Dim i As Integer
    
    If Index < 5 Then
        For i = Index + 1 To 5
            Call SaveMySetting("PreviousFiles", "PreviousFile" & i - 1, GetMySetting("PreviousFiles", "PreviousFile" & i, ""))
        Next i
    End If
    
    Call SaveMySetting("PreviousFiles", "PreviousFile5", "")

    'Now update the menu
    Call frmMain.UpdatePreviousFileMenu
End Sub

