VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImageLinkSyncAction"
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

Public Enum SyncModes
    MountSync
    Offset
    NoSync
End Enum

Public ExpTime As Double
Public DelayTime As Double
Public Bin As Integer
Public Filter As Integer
Public PixelScale As Double
Public NorthAngle As Double
Public FrameSize As FrameSizes
Public SyncMode As SyncModes
Public SlewToOriginalLocation As Integer

Public AbortListOnFailure As Boolean
Public RetryPlateSolveOnFailure As Boolean
Public SlewMountForRetry As Boolean
Public ArcminutesToSlew As Long
Public SkipIfRetrySucceeds As Boolean

Public AutosaveExposure As Integer
Public FileNamePrefix As String
Public UseGlobalSaveToLocation As Integer
Public FileSavePath As String

Public Selected As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 7
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(ExpTime)
    MySize = MySize + Len(Bin)
    MySize = MySize + Len(Filter)
    MySize = MySize + Len(PixelScale)
    MySize = MySize + Len(NorthAngle)
    MySize = MySize + Len(FrameSize)
    MySize = MySize + Len(SyncMode)
    MySize = MySize + Len(SlewToOriginalLocation)
    MySize = MySize + Len(DelayTime)
    
    MySize = MySize + Len(Selected)
        
    MySize = MySize + Len(AbortListOnFailure)
    MySize = MySize + Len(RetryPlateSolveOnFailure)
    MySize = MySize + Len(SlewMountForRetry)
    MySize = MySize + Len(ArcminutesToSlew)
    MySize = MySize + Len(SkipIfRetrySucceeds)
    
    MySize = MySize + Len(AutosaveExposure)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(FileNamePrefix)
    MySize = MySize + Len(UseGlobalSaveToLocation)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(FileSavePath)
        
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), ExpTime, Len(ExpTime))
    Index = Index + Len(ExpTime)
    Call CopyMemory(ptrByteArray(Index), Bin, Len(Bin))
    Index = Index + Len(Bin)
    Call CopyMemory(ptrByteArray(Index), Filter, Len(Filter))
    Index = Index + Len(Filter)
    Call CopyMemory(ptrByteArray(Index), PixelScale, Len(PixelScale))
    Index = Index + Len(PixelScale)
    Call CopyMemory(ptrByteArray(Index), NorthAngle, Len(NorthAngle))
    Index = Index + Len(NorthAngle)
    Call CopyMemory(ptrByteArray(Index), FrameSize, Len(FrameSize))
    Index = Index + Len(FrameSize)
    Call CopyMemory(ptrByteArray(Index), SyncMode, Len(SyncMode))
    Index = Index + Len(SyncMode)
    Call CopyMemory(ptrByteArray(Index), SlewToOriginalLocation, Len(SlewToOriginalLocation))
    Index = Index + Len(SlewToOriginalLocation)
    Call CopyMemory(ptrByteArray(Index), DelayTime, Len(DelayTime))
    Index = Index + Len(DelayTime)

    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)
    
    Call CopyMemory(ptrByteArray(Index), AbortListOnFailure, Len(AbortListOnFailure))
    Index = Index + Len(AbortListOnFailure)
    Call CopyMemory(ptrByteArray(Index), RetryPlateSolveOnFailure, Len(RetryPlateSolveOnFailure))
    Index = Index + Len(RetryPlateSolveOnFailure)
    Call CopyMemory(ptrByteArray(Index), SlewMountForRetry, Len(SlewMountForRetry))
    Index = Index + Len(SlewMountForRetry)
    Call CopyMemory(ptrByteArray(Index), ArcminutesToSlew, Len(ArcminutesToSlew))
    Index = Index + Len(ArcminutesToSlew)
    Call CopyMemory(ptrByteArray(Index), SkipIfRetrySucceeds, Len(SkipIfRetrySucceeds))
    Index = Index + Len(SkipIfRetrySucceeds)
    
    Call CopyMemory(ptrByteArray(Index), AutosaveExposure, Len(AutosaveExposure))
    Index = Index + Len(AutosaveExposure)
    StringLen = Len(FileNamePrefix)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray, Index, FileNamePrefix)       'Index is passed ByRef, so I don't need to increment it
    Call CopyMemory(ptrByteArray(Index), UseGlobalSaveToLocation, Len(UseGlobalSaveToLocation))
    Index = Index + Len(UseGlobalSaveToLocation)
    StringLen = Len(FileSavePath)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray, Index, FileSavePath)       'Index is passed ByRef, so I don't need to increment it
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long
    
    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(ExpTime, ptrByteArray(Index), Len(ExpTime))
    Index = Index + Len(ExpTime)
    Call CopyMemory(Bin, ptrByteArray(Index), Len(Bin))
    Index = Index + Len(Bin)
    Call CopyMemory(Filter, ptrByteArray(Index), Len(Filter))
    Index = Index + Len(Filter)
    Call CopyMemory(PixelScale, ptrByteArray(Index), Len(PixelScale))
    Index = Index + Len(PixelScale)
    Call CopyMemory(NorthAngle, ptrByteArray(Index), Len(NorthAngle))
    Index = Index + Len(NorthAngle)
    Call CopyMemory(FrameSize, ptrByteArray(Index), Len(FrameSize))
    Index = Index + Len(FrameSize)

    If SavedVersionNumber >= 3 Then
        Call CopyMemory(SyncMode, ptrByteArray(Index), Len(SyncMode))
        Index = Index + Len(SyncMode)
        Call CopyMemory(SlewToOriginalLocation, ptrByteArray(Index), Len(SlewToOriginalLocation))
        Index = Index + Len(SlewToOriginalLocation)
    Else
        SyncMode = MountSync
        SlewToOriginalLocation = 0
    End If
    
    If SavedVersionNumber >= 4 Then
        Call CopyMemory(DelayTime, ptrByteArray(Index), Len(DelayTime))
        Index = Index + Len(DelayTime)
    Else
        DelayTime = 0
    End If

    If SavedVersionNumber >= 5 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If
    
    If SavedVersionNumber >= 6 Then
        Call CopyMemory(AbortListOnFailure, ptrByteArray(Index), Len(AbortListOnFailure))
        Index = Index + Len(AbortListOnFailure)
        Call CopyMemory(RetryPlateSolveOnFailure, ptrByteArray(Index), Len(RetryPlateSolveOnFailure))
        Index = Index + Len(RetryPlateSolveOnFailure)
        Call CopyMemory(SlewMountForRetry, ptrByteArray(Index), Len(SlewMountForRetry))
        Index = Index + Len(SlewMountForRetry)
        Call CopyMemory(ArcminutesToSlew, ptrByteArray(Index), Len(ArcminutesToSlew))
        Index = Index + Len(ArcminutesToSlew)
        Call CopyMemory(SkipIfRetrySucceeds, ptrByteArray(Index), Len(SkipIfRetrySucceeds))
        Index = Index + Len(SkipIfRetrySucceeds)
    Else
        Me.AbortListOnFailure = False
        Me.RetryPlateSolveOnFailure = False
        Me.SlewMountForRetry = False
        Me.ArcminutesToSlew = 1
        Me.SkipIfRetrySucceeds = False
    End If
    
    If SavedVersionNumber >= 7 Then
        Call CopyMemory(AutosaveExposure, ptrByteArray(Index), Len(AutosaveExposure))
        Index = Index + Len(AutosaveExposure)
        Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
        Index = Index + Len(StringLen)
        FileNamePrefix = String(StringLen, " ")
        Call CopyByteArrayToString(FileNamePrefix, ptrByteArray(), Index)
        Call CopyMemory(UseGlobalSaveToLocation, ptrByteArray(Index), Len(UseGlobalSaveToLocation))
        Index = Index + Len(UseGlobalSaveToLocation)
        Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
        Index = Index + Len(StringLen)
        FileSavePath = String(StringLen, " ")
        Call CopyByteArrayToString(FileSavePath, ptrByteArray(), Index)
    Else
        Me.AutosaveExposure = vbUnchecked
        Me.FileNamePrefix = ""
        Me.UseGlobalSaveToLocation = vbChecked
        Me.FileSavePath = ""
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    ListString = "Plate Solve"
    
    If SyncMode = MountSync And SlewToOriginalLocation = 0 Then
        ListString = ListString & " and Sync"
    ElseIf SyncMode = MountSync And SlewToOriginalLocation = 1 Then
        ListString = ListString & ", Sync, and ReSlew"
    ElseIf SyncMode = Offset And SlewToOriginalLocation = 1 Then
        ListString = ListString & " and ReSlew"
    End If
    
    If Me.RetryPlateSolveOnFailure Then
        ListString = ListString & ", Retry"
        If Me.SlewMountForRetry Then
            ListString = ListString & " after " & Me.ArcminutesToSlew & " arcmin slew"
        End If
    End If
    
    If Me.AbortListOnFailure Then
        ListString = ListString & ", Abort on failure"
    End If
    
    If Filter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
        ListString = ListString & " with " & frmOptions.lstFilters.List(Filter) & " filter"
    End If
    
    ListString = ListString & "."
    
    BuildActionListString = ListString
End Function


