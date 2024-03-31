VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagerAction"
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

Public ImagerExpTime As Double
Public ImagerDelayTime As Double
Public ImagerNumExp As Long
Public ImagerBin As Integer
Public ImagerFilter As Integer
Public ImagerType As Integer
Public AutosaveExposure As Integer
Public FileNamePrefix As String
Public AutoguiderExpTime As Double
Public AutoguiderBin As Long
Public AutoguiderDitherStep As Double
Public AutoguiderDitherFreq As Long
Public AutoguiderDitherAmount As Double
Public AutoguiderXPos As Long
Public AutoguiderYPos As Long
Public AutoguiderEnabled As Integer
Public AutoguiderMinError As Double
Public ImageLinkAfterMeridianFlip As Boolean
Public clsImageLinkAction As New ImageLinkSyncAction
Public DoubleImageLink As Boolean
Public RotateAfterFlip As Integer
Public FrameSize As FrameSizes
Public UseGlobalSaveToLocation As Integer
Public FileSavePath As String
Public CenterAO As Integer
Public CalibrateImages As Integer
Public CustomFrameWidth As Long
Public CustomFrameHeight As Long
Public SyncToCurrentAtEnd As Integer

Public Selected As Boolean

Public UnguidedDither As Integer
Public MaintainDitherOnFilterChange As Integer

Public AutoguiderMaxGuideCycles As Integer

Public AutoguiderDelay As Integer

Public XAxisDither As Boolean
Public YAxisDither As Boolean

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 16
End Sub

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    
    MySize = 0
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(ImagerExpTime)
    MySize = MySize + Len(ImagerNumExp)
    MySize = MySize + Len(ImagerBin)
    MySize = MySize + Len(ImagerFilter)
    MySize = MySize + Len(ImagerType)
    MySize = MySize + Len(AutosaveExposure)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(FileNamePrefix)
    MySize = MySize + Len(AutoguiderExpTime)
    MySize = MySize + Len(AutoguiderBin)
    MySize = MySize + Len(AutoguiderDitherFreq)
    MySize = MySize + Len(StringLen)            'This is actually a place holder - used to have something here and need this for backwards compatibility
    MySize = MySize + Len(AutoguiderXPos)
    MySize = MySize + Len(AutoguiderYPos)
    MySize = MySize + Len(AutoguiderEnabled)
    MySize = MySize + Len(AutoguiderMinError)
    MySize = MySize + Len(ImageLinkAfterMeridianFlip)
    MySize = MySize + Len(AutoguiderDitherAmount)
    MySize = MySize + Len(AutoguiderDitherStep)
    MySize = MySize + clsImageLinkAction.ByteArraySize()
    MySize = MySize + Len(RotateAfterFlip)
    MySize = MySize + Len(ImagerDelayTime)
    MySize = MySize + Len(FrameSize)
    MySize = MySize + Len(UseGlobalSaveToLocation)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(FileSavePath)
    MySize = MySize + Len(CenterAO)
    MySize = MySize + Len(CalibrateImages)
    MySize = MySize + Len(CustomFrameWidth)
    MySize = MySize + Len(CustomFrameHeight)
    
    MySize = MySize + Len(Selected)
    
    MySize = MySize + Len(UnguidedDither)
    MySize = MySize + Len(MaintainDitherOnFilterChange)
    
    MySize = MySize + Len(AutoguiderMaxGuideCycles)
    
    MySize = MySize + Len(DoubleImageLink)
    
    MySize = MySize + Len(AutoguiderDelay)
    
    MySize = MySize + Len(XAxisDither)
    MySize = MySize + Len(YAxisDither)
    
    MySize = MySize + Len(SyncToCurrentAtEnd)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    
    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), ImagerExpTime, Len(ImagerExpTime))
    Index = Index + Len(ImagerExpTime)
    Call CopyMemory(ptrByteArray(Index), ImagerNumExp, Len(ImagerNumExp))
    Index = Index + Len(ImagerNumExp)
    Call CopyMemory(ptrByteArray(Index), ImagerBin, Len(ImagerBin))
    Index = Index + Len(ImagerBin)
    Call CopyMemory(ptrByteArray(Index), ImagerFilter, Len(ImagerFilter))
    Index = Index + Len(ImagerFilter)
    Call CopyMemory(ptrByteArray(Index), ImagerType, Len(ImagerType))
    Index = Index + Len(ImagerType)
    Call CopyMemory(ptrByteArray(Index), AutosaveExposure, Len(AutosaveExposure))
    Index = Index + Len(AutosaveExposure)
    
    StringLen = Len(FileNamePrefix)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray, Index, FileNamePrefix)     'Index is passed ByRef, so I don't need to increment it
    
    Call CopyMemory(ptrByteArray(Index), AutoguiderExpTime, Len(AutoguiderExpTime))
    Index = Index + Len(AutoguiderExpTime)
    Call CopyMemory(ptrByteArray(Index), AutoguiderBin, Len(AutoguiderBin))
    Index = Index + Len(AutoguiderBin)
    Call CopyMemory(ptrByteArray(Index), AutoguiderDitherFreq, Len(AutoguiderDitherFreq))
    Index = Index + Len(AutoguiderDitherFreq)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen)) 'This is actually a place holder - used to have something here and need this for backwards compatibility
    Index = Index + Len(StringLen)
    Call CopyMemory(ptrByteArray(Index), AutoguiderXPos, Len(AutoguiderXPos))
    Index = Index + Len(AutoguiderXPos)
    Call CopyMemory(ptrByteArray(Index), AutoguiderYPos, Len(AutoguiderYPos))
    Index = Index + Len(AutoguiderYPos)
    Call CopyMemory(ptrByteArray(Index), AutoguiderEnabled, Len(AutoguiderEnabled))
    Index = Index + Len(AutoguiderEnabled)
    Call CopyMemory(ptrByteArray(Index), AutoguiderMinError, Len(AutoguiderMinError))
    Index = Index + Len(AutoguiderMinError)
    Call CopyMemory(ptrByteArray(Index), ImageLinkAfterMeridianFlip, Len(ImageLinkAfterMeridianFlip))
    Index = Index + Len(ImageLinkAfterMeridianFlip)
    Call CopyMemory(ptrByteArray(Index), AutoguiderDitherAmount, Len(AutoguiderDitherAmount))
    Index = Index + Len(AutoguiderDitherAmount)
    Call CopyMemory(ptrByteArray(Index), AutoguiderDitherStep, Len(AutoguiderDitherStep))
    Index = Index + Len(AutoguiderDitherStep)
    
    Call clsImageLinkAction.SaveActionByteArray(ptrByteArray(), Index)  'Index is passed ByRef so I don't need to increment it
    
    Call CopyMemory(ptrByteArray(Index), RotateAfterFlip, Len(RotateAfterFlip))
    Index = Index + Len(RotateAfterFlip)
    Call CopyMemory(ptrByteArray(Index), ImagerDelayTime, Len(ImagerDelayTime))
    Index = Index + Len(ImagerDelayTime)
    Call CopyMemory(ptrByteArray(Index), FrameSize, Len(FrameSize))
    Index = Index + Len(FrameSize)
    Call CopyMemory(ptrByteArray(Index), UseGlobalSaveToLocation, Len(UseGlobalSaveToLocation))
    Index = Index + Len(UseGlobalSaveToLocation)
    
    StringLen = Len(FileSavePath)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray, Index, FileSavePath)       'Index is passed ByRef, so I don't need to increment it
    
    Call CopyMemory(ptrByteArray(Index), CenterAO, Len(CenterAO))
    Index = Index + Len(CenterAO)
    Call CopyMemory(ptrByteArray(Index), CalibrateImages, Len(CalibrateImages))
    Index = Index + Len(CalibrateImages)
    Call CopyMemory(ptrByteArray(Index), CustomFrameWidth, Len(CustomFrameWidth))
    Index = Index + Len(CustomFrameWidth)
    Call CopyMemory(ptrByteArray(Index), CustomFrameHeight, Len(CustomFrameHeight))
    Index = Index + Len(CustomFrameHeight)
    
    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)
    
    Call CopyMemory(ptrByteArray(Index), UnguidedDither, Len(UnguidedDither))
    Index = Index + Len(UnguidedDither)
    Call CopyMemory(ptrByteArray(Index), MaintainDitherOnFilterChange, Len(MaintainDitherOnFilterChange))
    Index = Index + Len(MaintainDitherOnFilterChange)
    
    Call CopyMemory(ptrByteArray(Index), AutoguiderMaxGuideCycles, Len(AutoguiderMaxGuideCycles))
    Index = Index + Len(AutoguiderMaxGuideCycles)

    Call CopyMemory(ptrByteArray(Index), DoubleImageLink, Len(DoubleImageLink))
    Index = Index + Len(DoubleImageLink)

    Call CopyMemory(ptrByteArray(Index), AutoguiderDelay, Len(AutoguiderDelay))
    Index = Index + Len(AutoguiderDelay)
    
    Call CopyMemory(ptrByteArray(Index), XAxisDither, Len(XAxisDither))
    Index = Index + Len(XAxisDither)
    Call CopyMemory(ptrByteArray(Index), YAxisDither, Len(YAxisDither))
    Index = Index + Len(YAxisDither)

    Call CopyMemory(ptrByteArray(Index), SyncToCurrentAtEnd, Len(SyncToCurrentAtEnd))
    Index = Index + Len(SyncToCurrentAtEnd)
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long
    Dim TempLong As Long

    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(ImagerExpTime, ptrByteArray(Index), Len(ImagerExpTime))
    Index = Index + Len(ImagerExpTime)
    Call CopyMemory(ImagerNumExp, ptrByteArray(Index), Len(ImagerNumExp))
    Index = Index + Len(ImagerNumExp)
    Call CopyMemory(ImagerBin, ptrByteArray(Index), Len(ImagerBin))
    Index = Index + Len(ImagerBin)
    Call CopyMemory(ImagerFilter, ptrByteArray(Index), Len(ImagerFilter))
    Index = Index + Len(ImagerFilter)
    Call CopyMemory(ImagerType, ptrByteArray(Index), Len(ImagerType))
    Index = Index + Len(ImagerType)
    Call CopyMemory(AutosaveExposure, ptrByteArray(Index), Len(AutosaveExposure))
    Index = Index + Len(AutosaveExposure)
    
    Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
    Index = Index + Len(StringLen)
    FileNamePrefix = String(StringLen, " ")
    Call CopyByteArrayToString(FileNamePrefix, ptrByteArray(), Index)
    
    Call CopyMemory(AutoguiderExpTime, ptrByteArray(Index), Len(AutoguiderExpTime))
    Index = Index + Len(AutoguiderExpTime)
    Call CopyMemory(AutoguiderBin, ptrByteArray(Index), Len(AutoguiderBin))
    Index = Index + Len(AutoguiderBin)
    Call CopyMemory(AutoguiderDitherFreq, ptrByteArray(Index), Len(AutoguiderDitherFreq))
    Index = Index + Len(AutoguiderDitherFreq)
    Call CopyMemory(TempLong, ptrByteArray(Index), Len(TempLong))
    Index = Index + Len(TempLong)
    Call CopyMemory(AutoguiderXPos, ptrByteArray(Index), Len(AutoguiderXPos))
    Index = Index + Len(AutoguiderXPos)
    Call CopyMemory(AutoguiderYPos, ptrByteArray(Index), Len(AutoguiderYPos))
    Index = Index + Len(AutoguiderYPos)
    Call CopyMemory(AutoguiderEnabled, ptrByteArray(Index), Len(AutoguiderEnabled))
    Index = Index + Len(AutoguiderEnabled)
    Call CopyMemory(AutoguiderMinError, ptrByteArray(Index), Len(AutoguiderMinError))
    Index = Index + Len(AutoguiderMinError)
    Call CopyMemory(ImageLinkAfterMeridianFlip, ptrByteArray(Index), Len(ImageLinkAfterMeridianFlip))
    Index = Index + Len(ImageLinkAfterMeridianFlip)
    
    If SavedVersionNumber >= 2 Then
        Call CopyMemory(AutoguiderDitherAmount, ptrByteArray(Index), Len(AutoguiderDitherAmount))
        Index = Index + Len(AutoguiderDitherAmount)
        Call CopyMemory(AutoguiderDitherStep, ptrByteArray(Index), Len(AutoguiderDitherStep))
        Index = Index + Len(AutoguiderDitherStep)
    Else
        AutoguiderDitherAmount = TempLong
        AutoguiderDitherStep = 1
    End If
    
    Call clsImageLinkAction.LoadActionByteArray(ptrByteArray, Index)
    
    If SavedVersionNumber >= 3 Then
        Call CopyMemory(RotateAfterFlip, ptrByteArray(Index), Len(RotateAfterFlip))
        Index = Index + Len(RotateAfterFlip)
    Else
        RotateAfterFlip = 0
    End If

    If SavedVersionNumber >= 4 Then
        Call CopyMemory(ImagerDelayTime, ptrByteArray(Index), Len(ImagerDelayTime))
        Index = Index + Len(ImagerDelayTime)
    Else
        ImagerDelayTime = 0
    End If
    
    If SavedVersionNumber >= 5 Then
        Call CopyMemory(FrameSize, ptrByteArray(Index), Len(FrameSize))
        Index = Index + Len(FrameSize)
    Else
        FrameSize = FullFrame
    End If
    
    If SavedVersionNumber >= 6 Then
        Call CopyMemory(UseGlobalSaveToLocation, ptrByteArray(Index), Len(UseGlobalSaveToLocation))
        Index = Index + Len(UseGlobalSaveToLocation)
        Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
        Index = Index + Len(StringLen)
        FileSavePath = String(StringLen, " ")
        Call CopyByteArrayToString(FileSavePath, ptrByteArray(), Index)
    Else
        UseGlobalSaveToLocation = 1
        FileSavePath = frmOptions.txtSaveTo.Text
    End If
    
    If SavedVersionNumber >= 7 Then
        Call CopyMemory(CenterAO, ptrByteArray(Index), Len(CenterAO))
        Index = Index + Len(CenterAO)
    Else
        CenterAO = 0
    End If
    
    If SavedVersionNumber >= 8 Then
        Call CopyMemory(CalibrateImages, ptrByteArray(Index), Len(CalibrateImages))
        Index = Index + Len(CalibrateImages)
    Else
        CalibrateImages = 0
    End If
    
    If SavedVersionNumber >= 9 Then
        Call CopyMemory(CustomFrameWidth, ptrByteArray(Index), Len(CustomFrameWidth))
        Index = Index + Len(CustomFrameWidth)
        Call CopyMemory(CustomFrameHeight, ptrByteArray(Index), Len(CustomFrameHeight))
        Index = Index + Len(CustomFrameHeight)
    Else
        CustomFrameWidth = 0
        CustomFrameHeight = 0
    End If

    If SavedVersionNumber >= 10 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If

    If SavedVersionNumber >= 11 Then
        Call CopyMemory(UnguidedDither, ptrByteArray(Index), Len(UnguidedDither))
        Index = Index + Len(UnguidedDither)
    Else
        UnguidedDither = vbUnchecked
    End If

    If SavedVersionNumber >= 12 Then
        Call CopyMemory(MaintainDitherOnFilterChange, ptrByteArray(Index), Len(MaintainDitherOnFilterChange))
        Index = Index + Len(MaintainDitherOnFilterChange)
    Else
        MaintainDitherOnFilterChange = vbUnchecked
    End If
    
    If SavedVersionNumber >= 13 Then
        Call CopyMemory(AutoguiderMaxGuideCycles, ptrByteArray(Index), Len(AutoguiderMaxGuideCycles))
        Index = Index + Len(AutoguiderMaxGuideCycles)
    Else
        AutoguiderMaxGuideCycles = 20
    End If

    If SavedVersionNumber >= 14 Then
        Call CopyMemory(DoubleImageLink, ptrByteArray(Index), Len(DoubleImageLink))
        Index = Index + Len(DoubleImageLink)
    
        Call CopyMemory(AutoguiderDelay, ptrByteArray(Index), Len(AutoguiderDelay))
        Index = Index + Len(AutoguiderDelay)
    Else
        DoubleImageLink = False
        
        AutoguiderDelay = 0
    End If
    
    If SavedVersionNumber >= 15 Then
        Call CopyMemory(XAxisDither, ptrByteArray(Index), Len(XAxisDither))
        Index = Index + Len(XAxisDither)
        
        Call CopyMemory(YAxisDither, ptrByteArray(Index), Len(YAxisDither))
        Index = Index + Len(YAxisDither)
    Else
        XAxisDither = True
        YAxisDither = True
    End If
        
    If SavedVersionNumber >= 16 Then
        Call CopyMemory(SyncToCurrentAtEnd, ptrByteArray(Index), Len(SyncToCurrentAtEnd))
        Index = Index + Len(SyncToCurrentAtEnd)
    Else
        SyncToCurrentAtEnd = vbUnchecked
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String

    If Camera.ImageTypes(ImagerType - 1) = "Bias" Then
        ListString = Camera.ImageTypes(ImagerType - 1) & " " & ImagerNumExp
    Else
        ListString = Camera.ImageTypes(ImagerType - 1) & " " & ImagerNumExp & "x" & ImagerExpTime & "s"
    End If

    If ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
        ListString = ListString & " " & frmOptions.lstFilters.List(ImagerFilter)
    End If
    
    If AutosaveExposure = vbChecked Then
        ListString = ListString & " AS "
        If Len(FileNamePrefix) <= 11 Then
            ListString = ListString & Chr(34) & FileNamePrefix & Chr(34)
        Else
            ListString = ListString & Chr(34) & Left(FileNamePrefix, 10) & "..." & Chr(34)
        End If
    End If
        
    If AutoguiderEnabled = vbChecked Then
        ListString = ListString & ", AG"
        If AutoguiderDitherAmount > 0 And AutoguiderDitherFreq > 0 Then
            ListString = ListString & " Dither every " & AutoguiderDitherFreq & " exp."
        End If
    ElseIf UnguidedDither = vbChecked Then
        If AutoguiderDitherAmount > 0 And AutoguiderDitherFreq > 0 Then
            ListString = ListString & " Dither every " & AutoguiderDitherFreq & " exp."
        End If
    End If
        
    If ImageLinkAfterMeridianFlip Then
        ListString = ListString & ", PM Plate Solve"
    End If
    
    BuildActionListString = ListString
End Function

