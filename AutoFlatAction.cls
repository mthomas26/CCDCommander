VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoFlatAction"
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

Public Enum FlatLocations
    FlatParkMount
    DuskSkyFlat
    DawnSkyFlat
    FixedLocation
    DoNotMove
End Enum

Public MinExpTime As Double
Public MaxExpTime As Double
Public AverageADU As Long 'This is now minimum ADU
Public DelayTime As Double
Public NumExp As Long
Public Bin As Integer
Public Filter As Integer
Public AutosaveExposure As Integer
Public FileNamePrefix As String
Public FlatLocation As FlatLocations
Public AltD As Double
Public AltM As Double
Public AltS As Double
Public AzimD As Double
Public AzimM As Double
Public AzimS As Double
Public ContinuouslyAdjust As Integer
Public FrameSize As FrameSizes

Public TwilightSelect As TwilightTypes

Public ActualSkipToTime As Date 'Note, this is for run time use only, no need to save it!

Private myFilters(MaxAutoFlatFilters) As Integer
Private myNumFilters As Long

Public UseGlobalSaveToLocation As Integer
Public FileSavePath As String

Public Selected As Boolean

Public MaximumADU As Long
Public TakeMatchingDarks As Integer
Public NumberOfDarksPerFlat As Integer
Public SetupFrameSize As FrameSizes
Private myRotations(MaxAutoFlatRotations) As Integer
Private myNumRotations As Long

Public DuskSunAltitudeStart As Double
Public DawnSunAltitudeStart As Double

Public DarkFrameTolerance As Double

Public FlipRotator As Integer

Private VersionNumber As Long

Private Sub Class_Initialize()
    VersionNumber = 13
End Sub

Public Property Let NumFilters(Value As Long)
    If Value > MaxAutoFlatFilters Then Value = MaxAutoFlatFilters
    
    myNumFilters = Value
End Property

Public Property Get NumFilters() As Long
    NumFilters = myNumFilters
End Property

Public Property Let Filters(Index As Long, Value As Integer)
    If Index >= 0 And Index < NumFilters Then
        myFilters(Index) = Value
    End If
End Property

Public Property Get Filters(Index As Long) As Integer
    If Index >= 0 And Index < NumFilters Then
        Filters = myFilters(Index)
    Else
        Filters = 99
    End If
End Property

Public Property Let NumRotations(Value As Long)
    If Value > MaxAutoFlatRotations Then Value = MaxAutoFlatRotations
    
    myNumRotations = Value
End Property

Public Property Get NumRotations() As Long
    NumRotations = myNumRotations
End Property

Public Property Let Rotations(Index As Long, Value As Integer)
    If Index >= 0 And Index < NumRotations Then
        myRotations(Index) = Value
    End If
End Property

Public Property Get Rotations(Index As Long) As Integer
    If Index >= 0 And Index < NumRotations Then
        Rotations = myRotations(Index)
    Else
        Rotations = 0
    End If
End Property

Public Function ByteArraySize() As Long
    Dim MySize As Long
    Dim StringLen As Long
    Dim Counter As Integer
    
    MySize = 0
    
    MySize = MySize + Len(VersionNumber)
    MySize = MySize + Len(MinExpTime)
    MySize = MySize + Len(MaxExpTime)
    MySize = MySize + Len(AverageADU)
    MySize = MySize + Len(DelayTime)
    MySize = MySize + Len(NumExp)
    MySize = MySize + Len(Bin)
    MySize = MySize + Len(Filter)
    MySize = MySize + Len(AutosaveExposure)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(FileNamePrefix)
    MySize = MySize + Len(FlatLocation)
    MySize = MySize + Len(AltD)
    MySize = MySize + Len(AltM)
    MySize = MySize + Len(AltS)
    MySize = MySize + Len(AzimD)
    MySize = MySize + Len(AzimM)
    MySize = MySize + Len(AzimS)
    
    MySize = MySize + Len(myNumFilters)
    For Counter = 0 To myNumFilters - 1
        MySize = MySize + Len(myFilters(Counter))
    Next Counter
    
    MySize = MySize + Len(UseGlobalSaveToLocation)
    MySize = MySize + Len(StringLen)
    MySize = MySize + Len(FileSavePath)
    MySize = MySize + Len(ContinuouslyAdjust)
    MySize = MySize + Len(FrameSize)
    MySize = MySize + Len(TwilightSelect)
    
    MySize = MySize + Len(Selected)
    
    MySize = MySize + Len(MaximumADU)
    MySize = MySize + Len(TakeMatchingDarks)
    MySize = MySize + Len(NumberOfDarksPerFlat)
    MySize = MySize + Len(SetupFrameSize)
    
    MySize = MySize + Len(myNumRotations)
    For Counter = 0 To myNumRotations - 1
        MySize = MySize + Len(myRotations(Counter))
    Next Counter
    
    MySize = MySize + Len(DuskSunAltitudeStart)
    MySize = MySize + Len(DawnSunAltitudeStart)
    MySize = MySize + Len(DarkFrameTolerance)
    
    MySize = MySize + Len(FlipRotator)
    
    ByteArraySize = MySize
End Function

Public Sub SaveActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim StringLen As Long
    Dim Counter As Integer

    Call CopyMemory(ptrByteArray(Index), VersionNumber, Len(VersionNumber))
    Index = Index + Len(VersionNumber)
    Call CopyMemory(ptrByteArray(Index), MinExpTime, Len(MinExpTime))
    Index = Index + Len(MinExpTime)
    Call CopyMemory(ptrByteArray(Index), MaxExpTime, Len(MaxExpTime))
    Index = Index + Len(MaxExpTime)
    Call CopyMemory(ptrByteArray(Index), AverageADU, Len(AverageADU))
    Index = Index + Len(AverageADU)
    Call CopyMemory(ptrByteArray(Index), DelayTime, Len(DelayTime))
    Index = Index + Len(DelayTime)
    Call CopyMemory(ptrByteArray(Index), NumExp, Len(NumExp))
    Index = Index + Len(NumExp)
    Call CopyMemory(ptrByteArray(Index), Bin, Len(Bin))
    Index = Index + Len(Bin)
    Call CopyMemory(ptrByteArray(Index), Filter, Len(Filter))
    Index = Index + Len(Filter)
    Call CopyMemory(ptrByteArray(Index), AutosaveExposure, Len(AutosaveExposure))
    Index = Index + Len(AutosaveExposure)
    
    StringLen = Len(FileNamePrefix)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray(), Index, FileNamePrefix)
    
    Call CopyMemory(ptrByteArray(Index), FlatLocation, Len(FlatLocation))
    Index = Index + Len(FlatLocation)
    Call CopyMemory(ptrByteArray(Index), AltD, Len(AltD))
    Index = Index + Len(AltD)
    Call CopyMemory(ptrByteArray(Index), AltM, Len(AltM))
    Index = Index + Len(AltM)
    Call CopyMemory(ptrByteArray(Index), AltS, Len(AltS))
    Index = Index + Len(AltS)
    Call CopyMemory(ptrByteArray(Index), AzimD, Len(AzimD))
    Index = Index + Len(AzimD)
    Call CopyMemory(ptrByteArray(Index), AzimM, Len(AzimM))
    Index = Index + Len(AzimM)
    Call CopyMemory(ptrByteArray(Index), AzimS, Len(AzimS))
    Index = Index + Len(AzimS)
    
    Call CopyMemory(ptrByteArray(Index), myNumFilters, Len(myNumFilters))
    Index = Index + Len(myNumFilters)
    For Counter = 0 To myNumFilters - 1
        Call CopyMemory(ptrByteArray(Index), myFilters(Counter), Len(myFilters(Counter)))
        Index = Index + Len(myFilters(Counter))
    Next Counter
    
    Call CopyMemory(ptrByteArray(Index), UseGlobalSaveToLocation, Len(UseGlobalSaveToLocation))
    Index = Index + Len(UseGlobalSaveToLocation)
    
    StringLen = Len(FileSavePath)
    Call CopyMemory(ptrByteArray(Index), StringLen, Len(StringLen))
    Index = Index + Len(StringLen)
    Call CopyStringToByteArray(ptrByteArray(), Index, FileSavePath)
    
    Call CopyMemory(ptrByteArray(Index), ContinuouslyAdjust, Len(ContinuouslyAdjust))
    Index = Index + Len(ContinuouslyAdjust)
    Call CopyMemory(ptrByteArray(Index), FrameSize, Len(FrameSize))
    Index = Index + Len(FrameSize)
    Call CopyMemory(ptrByteArray(Index), TwilightSelect, Len(TwilightSelect))
    Index = Index + Len(TwilightSelect)
    
    Call CopyMemory(ptrByteArray(Index), Selected, Len(Selected))
    Index = Index + Len(Selected)
    
    Call CopyMemory(ptrByteArray(Index), MaximumADU, Len(MaximumADU))
    Index = Index + Len(MaximumADU)
    Call CopyMemory(ptrByteArray(Index), TakeMatchingDarks, Len(TakeMatchingDarks))
    Index = Index + Len(TakeMatchingDarks)
    Call CopyMemory(ptrByteArray(Index), NumberOfDarksPerFlat, Len(NumberOfDarksPerFlat))
    Index = Index + Len(NumberOfDarksPerFlat)
    Call CopyMemory(ptrByteArray(Index), SetupFrameSize, Len(SetupFrameSize))
    Index = Index + Len(SetupFrameSize)
    
    Call CopyMemory(ptrByteArray(Index), myNumRotations, Len(myNumRotations))
    Index = Index + Len(myNumRotations)
    For Counter = 0 To myNumRotations - 1
        Call CopyMemory(ptrByteArray(Index), myRotations(Counter), Len(myRotations(Counter)))
        Index = Index + Len(myRotations(Counter))
    Next Counter
    
    Call CopyMemory(ptrByteArray(Index), DuskSunAltitudeStart, Len(DuskSunAltitudeStart))
    Index = Index + Len(DuskSunAltitudeStart)
    Call CopyMemory(ptrByteArray(Index), DawnSunAltitudeStart, Len(DawnSunAltitudeStart))
    Index = Index + Len(DawnSunAltitudeStart)
    Call CopyMemory(ptrByteArray(Index), DarkFrameTolerance, Len(DarkFrameTolerance))
    Index = Index + Len(DarkFrameTolerance)
    
    Call CopyMemory(ptrByteArray(Index), FlipRotator, Len(FlipRotator))
    Index = Index + Len(FlipRotator)
    
End Sub

Public Sub LoadActionByteArray(ByRef ptrByteArray() As Byte, ByRef Index As Long)
    Dim SavedVersionNumber As Long
    Dim StringLen As Long
    Dim Counter As Integer
    
    Call CopyMemory(SavedVersionNumber, ptrByteArray(Index), Len(SavedVersionNumber))
    Index = Index + Len(SavedVersionNumber)
    Call CopyMemory(MinExpTime, ptrByteArray(Index), Len(MinExpTime))
    Index = Index + Len(MinExpTime)
    Call CopyMemory(MaxExpTime, ptrByteArray(Index), Len(MaxExpTime))
    Index = Index + Len(MaxExpTime)
    Call CopyMemory(AverageADU, ptrByteArray(Index), Len(AverageADU))
    Index = Index + Len(AverageADU)
    Call CopyMemory(DelayTime, ptrByteArray(Index), Len(DelayTime))
    Index = Index + Len(DelayTime)
    Call CopyMemory(NumExp, ptrByteArray(Index), Len(NumExp))
    Index = Index + Len(NumExp)
    Call CopyMemory(Bin, ptrByteArray(Index), Len(Bin))
    Index = Index + Len(Bin)
    Call CopyMemory(Filter, ptrByteArray(Index), Len(Filter))
    Index = Index + Len(Filter)
    Call CopyMemory(AutosaveExposure, ptrByteArray(Index), Len(AutosaveExposure))
    Index = Index + Len(AutosaveExposure)
    
    Call CopyMemory(StringLen, ptrByteArray(Index), Len(StringLen))
    Index = Index + Len(StringLen)
    FileNamePrefix = String(StringLen, " ")
    Call CopyByteArrayToString(FileNamePrefix, ptrByteArray(), Index)
    
    If SavedVersionNumber >= 5 Then
        Call CopyMemory(FlatLocation, ptrByteArray(Index), Len(FlatLocation))
        Index = Index + Len(FlatLocation)
        Call CopyMemory(AltD, ptrByteArray(Index), Len(AltD))
        Index = Index + Len(AltD)
        Call CopyMemory(AltM, ptrByteArray(Index), Len(AltM))
        Index = Index + Len(AltM)
        Call CopyMemory(AltS, ptrByteArray(Index), Len(AltS))
        Index = Index + Len(AltS)
        Call CopyMemory(AzimD, ptrByteArray(Index), Len(AzimD))
        Index = Index + Len(AzimD)
        Call CopyMemory(AzimM, ptrByteArray(Index), Len(AzimM))
        Index = Index + Len(AzimM)
        Call CopyMemory(AzimS, ptrByteArray(Index), Len(AzimS))
        Index = Index + Len(AzimS)
    Else
        FlatLocation = FlatParkMount
        AltD = 0
        AltM = 0
        AltS = 0
        AzimD = 270
        AzimM = 0
        AzimS = 0
    End If
    
    If SavedVersionNumber >= 6 Then
        Call CopyMemory(myNumFilters, ptrByteArray(Index), Len(myNumFilters))
        Index = Index + Len(myNumFilters)
        For Counter = 0 To myNumFilters - 1
            Call CopyMemory(myFilters(Counter), ptrByteArray(Index), Len(myFilters(Counter)))
            Index = Index + Len(myFilters(Counter))
        Next Counter
    Else
        myNumFilters = 0
    End If

    If SavedVersionNumber >= 7 Then
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
    
    If SavedVersionNumber >= 8 Then
        Call CopyMemory(ContinuouslyAdjust, ptrByteArray(Index), Len(ContinuouslyAdjust))
        Index = Index + Len(ContinuouslyAdjust)
    Else
        ContinuouslyAdjust = 0
    End If
    
    If SavedVersionNumber >= 9 Then
        Call CopyMemory(FrameSize, ptrByteArray(Index), Len(FrameSize))
        Index = Index + Len(FrameSize)
        Call CopyMemory(TwilightSelect, ptrByteArray(Index), Len(TwilightSelect))
        Index = Index + Len(TwilightSelect)
    Else
        FrameSize = FullFrame
        TwilightSelect = Astronomical
    End If
    
    If SavedVersionNumber >= 10 Then
        Call CopyMemory(Selected, ptrByteArray(Index), Len(Selected))
        Index = Index + Len(Selected)
    Else
        Selected = True
    End If
    
    If SavedVersionNumber >= 11 Then
        Call CopyMemory(MaximumADU, ptrByteArray(Index), Len(MaximumADU))
        Index = Index + Len(MaximumADU)
        Call CopyMemory(TakeMatchingDarks, ptrByteArray(Index), Len(TakeMatchingDarks))
        Index = Index + Len(TakeMatchingDarks)
        Call CopyMemory(NumberOfDarksPerFlat, ptrByteArray(Index), Len(NumberOfDarksPerFlat))
        Index = Index + Len(NumberOfDarksPerFlat)
        Call CopyMemory(SetupFrameSize, ptrByteArray(Index), Len(SetupFrameSize))
        Index = Index + Len(SetupFrameSize)
        
        Call CopyMemory(myNumRotations, ptrByteArray(Index), Len(myNumRotations))
        Index = Index + Len(myNumRotations)
        For Counter = 0 To myNumRotations - 1
            Call CopyMemory(myRotations(Counter), ptrByteArray(Index), Len(myRotations(Counter)))
            Index = Index + Len(myRotations(Counter))
        Next Counter
    Else
        MaximumADU = AverageADU * 1.2
        TakeMatchingDarks = vbUnchecked
        NumberOfDarksPerFlat = 0
        SetupFrameSize = FullFrame
        myNumRotations = 0
    End If
    
    If SavedVersionNumber >= 12 Then
        Call CopyMemory(DuskSunAltitudeStart, ptrByteArray(Index), Len(DuskSunAltitudeStart))
        Index = Index + Len(DuskSunAltitudeStart)
        Call CopyMemory(DawnSunAltitudeStart, ptrByteArray(Index), Len(DawnSunAltitudeStart))
        Index = Index + Len(DawnSunAltitudeStart)
        Call CopyMemory(DarkFrameTolerance, ptrByteArray(Index), Len(DarkFrameTolerance))
        Index = Index + Len(DarkFrameTolerance)
    Else
        DuskSunAltitudeStart = -0.83
        If TwilightSelect = Civil Then
            DawnSunAltitudeStart = -6
        ElseIf TwilightSelect = Nautical Then
            DawnSunAltitudeStart = -12
        ElseIf TwilightSelect = Astronomical Then
            DawnSunAltitudeStart = -18
        End If
        
        DarkFrameTolerance = 1
    End If
    
    If SavedVersionNumber >= 13 Then
        Call CopyMemory(FlipRotator, ptrByteArray(Index), Len(FlipRotator))
        Index = Index + Len(FlipRotator)
    Else
        FlipRotator = False
    End If
End Sub

Public Function BuildActionListString() As String
    Dim ListString As String
    
    If FlatLocation = DawnSkyFlat Then
        ListString = "Automatic Dawn Flat, "
    ElseIf FlatLocation = DuskSkyFlat Then
        ListString = "Automatic Dusk Flat, "
    Else
        
        ListString = "Automatic Flat, "
    End If
    
    If (myNumFilters > 0) And (myNumRotations > 0) Then
        ListString = ListString & (NumExp * myNumFilters * myNumRotations) & " exposures"
    ElseIf (myNumRotations > 0) Then
        ListString = ListString & (NumExp * myNumRotations) & " exposures"
    ElseIf (myNumFilters > 0) Then
        ListString = ListString & (NumExp * myNumFilters) & " exposures"
    Else
        ListString = ListString & NumExp & " exposures"
    End If
    
    If (TakeMatchingDarks = vbChecked) Then
        ListString = ListString & ", " & NumberOfDarksPerFlat & " darks/flat"
    End If
    
    ListString = ListString & "."
    
    BuildActionListString = ListString
End Function
    
