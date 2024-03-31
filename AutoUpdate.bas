Attribute VB_Name = "AutoUpdate"
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

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Public Enum CheckForNewVersionResult
    NoNewVersionAvailable
    NewVersionAvailable
    DownloadError
End Enum

Public Function CheckForNewVersion() As CheckForNewVersionResult
    Dim sSourceUrl As String
    Dim sLocalFile As String
    Dim hfile As Long
    Dim Major As Integer
    Dim Minor As Integer
    Dim Revision As Integer
    
    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/Version.txt"
    sLocalFile = App.Path & "\Version.txt"
    
    If DownloadFile(sSourceUrl, sLocalFile) Then
    
        hfile = FreeFile
        Open sLocalFile For Input As #hfile
    
        Input #hfile, Major, Minor, Revision
    
        If (Major > App.Major) Or (Major = App.Major And Minor > App.Minor) Or (Major = App.Major And Minor = App.Minor And Revision > App.Revision) Then
            CheckForNewVersion = NewVersionAvailable
        Else
            CheckForNewVersion = NoNewVersionAvailable
        End If
    
        Close #hfile
    Else
        CheckForNewVersion = DownloadError
    End If
End Function

Public Sub RemoveFileVersion()
    On Error Resume Next
    Kill App.Path & "\Version.txt"
    On Error GoTo 0
End Sub

Public Function GetVersionHistory(myTextBox As RichTextBox) As Boolean
    Dim sSourceUrl As String
    Dim sLocalFile As String
    
    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/VersionHistory.txt"
    sLocalFile = App.Path & "\VersionHistory.txt"
    
    If DownloadFile(sSourceUrl, sLocalFile) Then
        myTextBox.LoadFile (sLocalFile)
        GetVersionHistory = True
    Else
        GetVersionHistory = False
    End If
End Function

Public Function CheckForNewHelpVersion() As CheckForNewVersionResult
    Dim sSourceUrl As String
    Dim sLocalFile As String
    Dim hfile As Long
    Dim inetMajor As Integer
    Dim inetMinor As Integer
    Dim inetRevision As Integer
    Dim myMajor As Integer
    Dim myMinor As Integer
    Dim myRevision As Integer
        
    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/HelpVersion.txt"
    sLocalFile = App.Path & "\HelpVersion.txt"
    
    'open local HelpVersion.txt file
    'if file doesn't exist, then I must need to get the new verison
    'otherwise compare it to the version on the web
    On Error Resume Next
    hfile = FreeFile
    Open sLocalFile For Input As #hfile
    If Err.Number = 0 Then
        Input #hfile, myMajor, myMinor, myRevision
        Close #hfile
        
        If DownloadFile(sSourceUrl, sLocalFile) Then
            hfile = FreeFile
            Open sLocalFile For Input As #hfile
        
            Input #hfile, inetMajor, inetMinor, inetRevision
        
            If (inetMajor > myMajor) Or (inetMajor = myMajor And inetMinor > myMinor) Or (inetMajor = myMajor And inetMinor = myMinor And inetRevision > myRevision) Then
                CheckForNewHelpVersion = NewVersionAvailable
            Else
                CheckForNewHelpVersion = NoNewVersionAvailable
            End If
        
            Close #hfile
        Else
            CheckForNewHelpVersion = DownloadError
        End If
    Else
        Close #hfile
        
        'Couldn't get HelpVersion.txt file, must be a new version available
        CheckForNewHelpVersion = NewVersionAvailable
        
        'Need to download HelpVersion.txt anyway
        If Not DownloadFile(sSourceUrl, sLocalFile) Then
            'problem, signal error
            CheckForNewHelpVersion = DownloadError
        End If
    End If
    On Error GoTo 0
End Function

Public Sub RemoveHelpFileVersion()
    On Error Resume Next
    Kill App.Path & "\HelpVersion.txt"
    On Error GoTo 0
End Sub

Public Function GetNewHelpFile() As Boolean
    Dim sSourceUrl As String
    Dim sLocalFile As String
    
    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/CCDCommander.chm"
    sLocalFile = App.Path & "\CCDCommander.chm"
    
    GetNewHelpFile = DownloadFile(sSourceUrl, sLocalFile)
End Function

Public Function CheckForNewUpdaterVersion() As CheckForNewVersionResult
    Dim sSourceUrl As String
    Dim sLocalFile As String
    Dim hfile As Long
    Dim inetMajor As Integer
    Dim inetMinor As Integer
    Dim inetRevision As Integer
    Dim myMajor As Integer
    Dim myMinor As Integer
    Dim myRevision As Integer
        
    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/UpdaterVersion.txt"
    sLocalFile = App.Path & "\UpdaterVersion.txt"
    
    'open local UpdaterVersion.txt file
    'if file doesn't exist, then I must need to get the new verison
    'otherwise compare it to the version on the web
    On Error Resume Next
    hfile = FreeFile
    Open sLocalFile For Input As #hfile
    If Err.Number = 0 Then
        Input #hfile, myMajor, myMinor, myRevision
        Close #hfile
        
        If DownloadFile(sSourceUrl, sLocalFile) Then
            hfile = FreeFile
            Open sLocalFile For Input As #hfile
        
            Input #hfile, inetMajor, inetMinor, inetRevision
        
            If (inetMajor > myMajor) Or (inetMajor = myMajor And inetMinor > myMinor) Or (inetMajor = myMajor And inetMinor = myMinor And inetRevision > myRevision) Then
                CheckForNewUpdaterVersion = NewVersionAvailable
            Else
                CheckForNewUpdaterVersion = NoNewVersionAvailable
            End If
        
            Close #hfile
        Else
            CheckForNewUpdaterVersion = DownloadError
        End If
    Else
        Close #hfile
        
        'Couldn't get UpdaterVersion.txt file, must be a new version available
        CheckForNewUpdaterVersion = NewVersionAvailable
        
        'Need to download UpdaterVersion.txt anyway
        If Not DownloadFile(sSourceUrl, sLocalFile) Then
            'problem, signal error
            CheckForNewUpdaterVersion = DownloadError
        End If
    End If
    On Error GoTo 0
End Function

Public Sub RemoveUpdaterFileVersion()
    On Error Resume Next
    Kill App.Path & "\UpdaterVersion.txt"
    On Error GoTo 0
End Sub

Public Function GetNewUpdaterFile() As Boolean
    Dim sSourceUrl As String
    Dim sLocalFile As String
    
    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/CCDCommanderUpdater.exe"
    sLocalFile = App.Path & "\CCDCommanderUpdater.exe"
    GetNewUpdaterFile = DownloadFile(sSourceUrl, sLocalFile)

    sSourceUrl = "http://ccdcommander.astromatt.com/AutoUpdate/Download.exe"
    sLocalFile = App.Path & "\Download.exe"
    GetNewUpdaterFile = DownloadFile(sSourceUrl, sLocalFile)
End Function

Private Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean
  
  'Download the file. BINDF_GETNEWESTVERSION forces
  'the API to download from the specified source.
  'Passing 0& as dwReserved causes the locally-cached
  'copy to be downloaded, if available. If the API
  'returns ERROR_SUCCESS (0), DownloadFile returns True.
   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS
   
End Function

