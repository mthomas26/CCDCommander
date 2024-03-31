VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDWControl"
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

Private objDDW As Object
Private DomeCoupled As Boolean

Public Function IsDomeOpen(Optional Force As Boolean = False, Optional State As Boolean = False) As Boolean
    Dim StartTime As Date
    Dim RetryCount As Integer
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
        
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            Call AddToStatus("Retrying...")
        End If
    Next RetryCount
        
    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "DDW reports always busy - unable to get shutter status." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
        
        Call Err.Raise(-1)
    End If
        
    If objDDW.statIsShutterOpen Then
        IsDomeOpen = True
    Else
        IsDomeOpen = False
    End If
End Function

Public Property Get IsDomeCoupled() As Boolean
    IsDomeCoupled = DomeCoupled
    
    'Below doesn't work for some reason - just keep track of the state of things
'    Do While objDDW.statIsBusy And Not Aborted
'        Call Wait(0.1)
'    Loop
'    Call AddToStatus("objDDW.optSlaveMode = " & objDDW.optSlaveMode)
'    IsDomeCoupled = objDDW.optSlaveMode
'    Call AddToStatus("IsDomeCoupled = " & IsDomeCoupled)
End Property

Public Sub DomeOpen()
    Dim StartTime As Date
    Dim RetryCount As Integer
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
    
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
        
    If (RetryCount <= Settings.DDWRetryCount) Then
        For RetryCount = 1 To Settings.DDWRetryCount
            Call objDDW.actOpenShutter
        
            StartTime = Now
            Do While Not objDDW.statIsShutterOpen() And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
                Call Wait(1)
            Loop

            If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
                'Didn't timeout...continue on
                Exit For
            Else
                Call AddToStatus("Timed out waiting for DDW.")
                If objDDW.statIsBusy Then
                    Call AddToStatus("Stopping current DDW process.")
                    Call objDDW.actStop
                    Call Wait(10)
                End If
                If (RetryCount < Settings.DDWRetryCount) Then
                    Call AddToStatus("Retrying...")
                Else
                    Call AddToStatus("Giving up.")
                End If
            End If
        Next RetryCount
    End If
    
    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "CCD Commander unable to open dome." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
        
        Call Err.Raise(-1)
    End If
        
End Sub

Public Sub DomeClose()
    Dim StartTime As Date
    Dim RetryCount As Integer
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
    
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
        
    If (RetryCount <= Settings.DDWRetryCount) Then
        'Make sure Dome is at the Home position!
        Call Me.DomeHome
    
        'Wait 10s to ensure DDW is not busy with the home command
        Call Wait(10)
    
        For RetryCount = 1 To Settings.DDWRetryCount
            Call objDDW.actCloseShutter

            StartTime = Now
            Do While Not objDDW.statIsShutterClosed() And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
                Call Wait(1)
            Loop
            If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
                'Didn't timeout...continue on
                Exit For
            Else
                Call AddToStatus("Timed out waiting for DDW.")
                If objDDW.statIsBusy Then
                    Call AddToStatus("Stopping current DDW process.")
                    Call objDDW.actStop
                    Call Wait(10)
                End If
                If (RetryCount < Settings.DDWRetryCount) Then
                    Call AddToStatus("Retrying...")
                Else
                    Call AddToStatus("Giving up.")
                End If
            End If
        Next RetryCount
    End If
    
    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "CCD Commander unable to close dome." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
    
        Call Err.Raise(-1)
    End If
End Sub

Public Sub ConnectToDome()
    Dim StartTime As Date
    Dim RetryCount As Integer
    
    Set objDDW = CreateObject("TIDigitalDomeWorks.DomeControl")
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(1)
        Loop
    
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
    
    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "DDW reports always busy - unable to connect to dome." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
    
        Call Err.Raise(-1)
    End If
        
    objDDW.optCommandBlockingMode = False
End Sub

Public Sub DomeCouple()
    Dim StartTime As Date
    Dim RetryCount As Integer
    Dim Counter As Integer
    Dim ErrNo As Long
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
        
    If (RetryCount <= Settings.DDWRetryCount) Then
        For RetryCount = 1 To Settings.DDWRetryCount
            On Error Resume Next
            Do
                Call Wait(1)
                Counter = Counter + 1
                objDDW.optSlaveMode = True
                If Err.Number = 0 Then Exit Do
            Loop While Counter < 5 And Not Aborted
            ErrNo = Err.Number
            On Error GoTo 0
                
            Call Wait(1)
        
            Counter = 0
            Do While Not objDDW.statIsBusy() And Not Aborted And Counter < 60
                Call Wait(0.1)
                Counter = Counter + 1
            Loop

            StartTime = Now
            Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
                Call Wait(1)
            Loop
        
            If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
                'Didn't timeout...continue on
                Exit For
            Else
                Call AddToStatus("Timed out waiting for DDW.")
                If objDDW.statIsBusy Then
                    Call AddToStatus("Stopping current DDW process.")
                    Call objDDW.actStop
                    Call Wait(10)
                End If
                If (RetryCount < Settings.DDWRetryCount) Then
                    Call AddToStatus("Retrying...")
                Else
                    Call AddToStatus("Giving up.")
                End If
            End If
        Next RetryCount
        
        If (RetryCount < Settings.DDWRetryCount) Then
            DomeCoupled = True
        End If
    End If

    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "CCD Commander unable to couple the dome to the mount." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
    
        Call Err.Raise(-1)
    End If
        
End Sub

Public Sub DomeUnCouple()
    Dim StartTime As Date
    Dim RetryCount As Integer
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
        
    If (RetryCount <= Settings.DDWRetryCount) Then
        objDDW.optSlaveMode = False
    
        DomeCoupled = False
    End If

    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "CCD Commander unable to uncouple the dome from the mount." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
    
        Call Err.Raise(-1)
    End If
End Sub

Public Sub DomeHome()
    Dim StartTime As Date
    Dim RetryCount As Integer
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
        
    If (RetryCount <= Settings.DDWRetryCount) Then
        For RetryCount = 1 To Settings.DDWRetryCount
            Call objDDW.actGoHome
            
            StartTime = Now
            Do While Not objDDW.statDomeIsHome And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
                Call Wait(1)
            Loop
            If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
                'Didn't timeout...continue on
                Exit For
            Else
                Call AddToStatus("Timed out waiting for DDW.")
                If objDDW.statIsBusy Then
                    Call AddToStatus("Stopping current DDW process.")
                    Call objDDW.actStop
                    Call Wait(10)
                End If
                If (RetryCount < Settings.DDWRetryCount) Then
                    Call AddToStatus("Retrying...")
                Else
                    Call AddToStatus("Giving up.")
                End If
            End If
        Next RetryCount
    End If

    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "CCD Commander unable to home the dome." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
        
        Call Err.Raise(-1)
    End If
End Sub

Public Sub DomePark()
    'no park function for DDW
    Call AddToStatus("DDW does not support a park function - please use Home instead.")
End Sub

Public Sub DisconnectFromDome()
    Set objDDW = Nothing
End Sub

Public Sub SlewToAzimuth(Azimuth As Integer)
    Dim StartTime As Date
    Dim RetryCount As Integer
    Dim Counter As Integer
    Dim ErrNo As Long
    
    For RetryCount = 1 To Settings.DDWRetryCount
        StartTime = Now
        Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
            Call Wait(0.1)
        Loop
        If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
            'Didn't timeout...continue on
            Exit For
        Else
            Call AddToStatus("Timed out waiting for DDW.")
            If objDDW.statIsBusy Then
                Call AddToStatus("Stopping current DDW process.")
                Call objDDW.actStop
                Call Wait(10)
            End If
            If (RetryCount < Settings.DDWRetryCount) Then
                Call AddToStatus("Retrying...")
            Else
                Call AddToStatus("Giving up.")
            End If
        End If
    Next RetryCount
        
    If (RetryCount <= Settings.DDWRetryCount) Then
        
        Call DomeUnCouple   'Cannot slew if coupled
        
        For RetryCount = 1 To Settings.DDWRetryCount
            On Error Resume Next
            Do
                Call Wait(1)
                Counter = Counter + 1
                
                Call objDDW.actGotoAzimuth(Azimuth)
                
                If Err.Number = 0 Then Exit Do
            Loop While Counter < 5 And Not Aborted
            ErrNo = Err.Number
            On Error GoTo 0
                
            Call Wait(1)
        
            Counter = 0
            Do While Not objDDW.statIsBusy() And Not Aborted And Counter < 60
                Call Wait(0.1)
                Counter = Counter + 1
            Loop

            StartTime = Now
            Do While objDDW.statIsBusy And Not Aborted And (DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60))
                Call Wait(1)
            Loop
        
            If DateDiff("s", StartTime, Now) < (Settings.DDWTimeout * 60) Then
                'Didn't timeout...continue on
                Exit For
            Else
                Call AddToStatus("Timed out waiting for DDW.")
                If objDDW.statIsBusy Then
                    Call AddToStatus("Stopping current DDW process.")
                    Call objDDW.actStop
                    Call Wait(10)
                End If
                If (RetryCount < Settings.DDWRetryCount) Then
                    Call AddToStatus("Retrying...")
                Else
                    Call AddToStatus("Giving up.")
                End If
            End If
        Next RetryCount
    End If

    If RetryCount > Settings.DDWRetryCount Then
        If frmOptions.chkEMailAlert(EMailAlertIndexes.DomeOpFailed).Value = vbChecked Then
            'Send e-mail!
            Call EMail.SendEMail(frmMain, "CCD Commander DDW Operation Failed", "CCD Commander unable to couple the dome to the mount." & vbCrLf & "Event occurred at: " & Format(Now, "hh:mm:ss"))
        End If
    
        Call Err.Raise(-1)
    End If
End Sub

