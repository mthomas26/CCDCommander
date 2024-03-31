VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCCDSoftControlAO"
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
Option Base 0

Private objAutoguider As CCDSoftLib.CCDSoftCamera
Private objImager As CCDSoftLib.CCDSoftCamera
Attribute objImager.VB_VarHelpID = -1
Private AutoguiderEnabled As Boolean
Private myAutosaveState As Boolean
Private myAutoSavePath As String
Private myAutosaveFN As String
Private StartedImage As Boolean
Private myPixelSize As Double

Public Sub Abort()
    On Error Resume Next
    objImager.Abort
    objAutoguider.Abort
    On Error GoTo 0
End Sub

Public Property Let AGBinXY(Value As Integer)
    If AutoguiderEnabled Then
        objAutoguider.BinX = Value
        objAutoguider.BinY = Value
    End If
End Property

Public Property Let AGExposureTime(Value As Double)
    If AutoguiderEnabled Then
        objAutoguider.ExposureTime = Value
        objAutoguider.AutoguiderExposureTime = Value
    End If
End Property

Public Property Get AGExposureTime() As Double
    If AutoguiderEnabled Then
        AGExposureTime = objAutoguider.AutoguiderExposureTime
    End If
End Property

Public Property Let AGFilterNumber(Value As Integer)
    If AutoguiderEnabled Then
        objAutoguider.FilterIndexZeroBased = Value
    End If
End Property

Public Property Let AGImageType(Value As Integer)
    If AutoguiderEnabled Then
        objAutoguider.Frame = Value
    End If
End Property

Public Property Let AGReduction(Value As Integer)
    If AutoguiderEnabled Then
        objAutoguider.ImageReduction = Value
    End If
End Property

Public Sub AGTakeImage()
    If AutoguiderEnabled Then
        objAutoguider.TakeImage
    End If
End Sub

Public Property Get AGTakeImageComplete() As Boolean
    If AutoguiderEnabled Then
        If objAutoguider.State = cdStateNone Then
            AGTakeImageComplete = True
        Else
            AGTakeImageComplete = False
        End If
    Else
        AGTakeImageComplete = True
    End If
End Property

Public Property Get AutoguiderInterval() As Double
    If AutoguiderEnabled Then
        AutoguiderInterval = (objAutoguider.AutoguiderExposureTime + (objAutoguider.AutoguiderDelayAfterCorrection / 1000))
    Else
        AutoguiderInterval = 0
    End If
End Property

Public Sub AutoSave(OnOff As Boolean, Optional Path As String = "*", Optional FileName As String = "*")
    myAutosaveState = OnOff
'
'    If Path <> "" Then
'        myAutosavePath = Path
'    End If
'
'    If FileName <> "" Then
'        myAutosaveFN = FileName
'    End If

    If OnOff = True Then
        objImager.AutoSaveOn = 1
        If Path <> "*" Then
            If Right(Path, 1) = "\" Then
                objImager.AutoSavePath = Mid(Path, 1, Len(Path) - 1)
            Else
                objImager.AutoSavePath = Path
            End If
        End If
        If FileName <> "*" Then
            objImager.AutoSavePrefix = FileName
        End If
    Else
        objImager.AutoSaveOn = 0
    End If
End Sub

Public Property Get AutoSaveState() As Boolean
    If objImager.AutoSaveOn = 1 Then
        AutoSaveState = True
    Else
        AutoSaveState = False
    End If
End Property

Public Property Get AverageADUOfExposure(FrameSize As FrameSizes) As Double
    Dim myImage As Object
    Dim ImageWidth As Long
    Dim ImageHeight As Long
    Dim LocalImage As Variant
    Dim CurX As Long
    Dim CurY As Long
    Dim Total As Double
    
    Dim XStart As Long
    Dim YStart As Long
    Dim YStop As Long
    Dim XStop As Long
    
    Set myImage = CreateObject("CCDSoft.Image")
    Call myImage.AttachToActiveImager

    With myImage
        ImageWidth = .Width
        ImageHeight = .Height
    
        LocalImage = .DataArray
    End With
    
    If FrameSize = FullFrame Then
        XStart = 0
        YStart = 0
        XStop = ImageWidth
        YStop = ImageHeight
    ElseIf FrameSize = HalfFrame Then
        XStart = ImageWidth / 4
        YStart = ImageHeight / 4
        XStop = (ImageWidth * 3) / 4
        YStop = (ImageHeight * 3) / 4
    ElseIf FrameSize = QuarterFrame Then
        XStart = (ImageWidth * 3) / 8
        YStart = (ImageHeight * 3) / 8
        XStop = (ImageWidth * 5) / 8
        YStop = (ImageHeight * 5) / 8
    End If
    
    Total = 0
    For CurX = XStart To XStop - 1
        For CurY = YStart To YStop - 1
            Total = Total + LocalImage(CurX, CurY)
        Next CurY
    Next CurX
    
    Total = Total / ((XStop - XStart) * (YStop - YStart))
    
    AverageADUOfExposure = Total
End Property

Public Property Get BinX() As Integer
    BinX = objImager.BinX
End Property

Public Property Let BinXY(Value As Integer)
    objImager.BinX = Value
    objImager.BinY = Value
End Property

Public Property Get BinY() As Integer
    BinY = objImager.BinY
End Property

Public Sub ConnectToCamera()
    If objImager Is Nothing Then
        Set objImager = New CCDSoftCamera
    End If
    If objAutoguider Is Nothing Then
        Set objAutoguider = New CCDSoftCamera
    End If
    
    'Connect to the camera
    objImager.Connect
    objImager.Asynchronous = True
    objImager.ImageReduction = cdNone
    objImager.AutoSaveOn = 0
    myAutosaveState = False
    objImager.ToNewWindow = 0
    
    objAutoguider.Autoguider = 1
    On Error Resume Next
    objAutoguider.Connect
    If Err.Number <> 0 Then
        Call AddToStatus("No autoguider detected.")
        AutoguiderEnabled = False
    Else
        AutoguiderEnabled = True
    End If
    On Error GoTo 0
    
    If AutoguiderEnabled Then
        objAutoguider.Asynchronous = True
    End If

    objImager.Series = 1
    objImager.Delay = 0
End Sub

Public Sub ConnectToTelescope()
    'Do nothing - automatic with TheSky
End Sub

Public Sub DisconnectFromTelescope()
    'Do nothing - automatic with TheSky
End Sub

Public Sub CenterAO()
    If AutoguiderEnabled Then
        'just disconnect and reconnect the camera - but be sure to leave the cooling alone!
        Dim PreviousTempRegulatorSetting As Boolean
        
        PreviousTempRegulatorSetting = objAutoguider.ShutDownTemperatureRegulationOnDisconnect
        
        objAutoguider.ShutDownTemperatureRegulationOnDisconnect = False
        objAutoguider.Disconnect
        Call Wait(1)
        objAutoguider.Connect
        objAutoguider.ShutDownTemperatureRegulationOnDisconnect = PreviousTempRegulatorSetting
        objAutoguider.ExposureTime = 0.11
        objAutoguider.TakeImage
    
        'wait for image
        Do Until AGTakeImageComplete Or Aborted
            Call Wait(1)
        Loop
    End If
End Sub

Public Property Let CoolerState(Value As Boolean)
    If Value Then
        objImager.RegulateTemperature = 1
    Else
        objImager.RegulateTemperature = 0
    End If
End Property

Public Property Get CoolerState() As Boolean
    CoolerState = objImager.RegulateTemperature
End Property

Public Property Get CoolerPower() As Double
    On Error Resume Next
    CoolerPower = objImager.ThermalElectricCoolerPower
    If Err.Number <> 0 Then
        CoolerPower = -1
    End If
    On Error GoTo 0
End Property

Public Property Get ExposureTime() As Double
    ExposureTime = objImager.ExposureTime
End Property

Public Property Let ExposureTime(Value As Double)
    objImager.ExposureTime = Value
End Property

Public Property Let FanOn(Value As Boolean)
    objImager.SBIGFanOn = Value
End Property

Public Property Get FastDownloadEnabled() As Boolean
    FastDownloadEnabled = False
End Property

Public Property Let FastDownloadEnabled(Value As Boolean)
    'Do nothing
End Property

Public Property Get FilterNumber() As Integer
    FilterNumber = objImager.FilterIndexZeroBased
End Property

Public Property Let FilterNumber(Value As Integer)
    objImager.FilterIndexZeroBased = Value
End Property

Public Function FindGuideStar(MinADU As Long, MaxADU As Long, GuideBoxX As Long, GuideBoxY As Long, MaxDither As Double, MaxImageX As Long, MaxImageY As Long) As Long
    Dim CurrentXPos As Integer
    Dim CurrentYPos As Integer
    Dim BoxXStop As Integer
    Dim BoxYStop As Integer
    Dim Result As Variant
    Dim ResultXPos As Integer
    Dim ResultYPos As Integer
    Dim GuiderImage As Object
    Dim MaxVal As Variant
    Dim Counter As Long
    Dim XPos As Integer
    Dim YPos As Integer
    Dim ThisStarVal As Variant
    Dim ThisStarXPos As Integer
    Dim ThisStarYPos As Integer
    Dim ImageArray As Variant
    Dim MaxFieldVal As Variant
    Dim XSize As Integer
    Dim YSize As Integer
    Dim BoxXPos As Integer
    Dim BoxYPos As Integer
    Dim IgnoreStar As Boolean
    Dim IgnoredLocationsX As Collection
    Dim IgnoredLocationsY As Collection
    Dim IgnoredLocationCounter As Integer
    
    Dim AverageADU As Double
    Dim StandardDeviation As Double
    Dim StdDevCounter As Integer
        
    Set IgnoredLocationsX = New Collection
    Set IgnoredLocationsY = New Collection
    
    Set GuiderImage = CreateObject("CCDSoft.Image")
    
    If MainMod.UseFakeGuideExposure Then
        GuiderImage.Path = App.Path & "\GuiderTest.fit"
        Call GuiderImage.Open
    Else
        Call GuiderImage.AttachToActiveAutoguider
    End If
    
    'Moved these out of the for loop
    With GuiderImage
        XSize = .Width
        YSize = .Height
        ImageArray = .DataArray
    End With
    
    MaxVal = 0
    If Not Aborted Then
        'Guide Star MUST be more than 16 pixels away from the edge
        For CurrentXPos = (GuideBoxX + MaxDither) To XSize - ((GuideBoxX + 1) + MaxDither) Step GuideBoxX * 2
            If Aborted Then Exit For
            For CurrentYPos = (GuideBoxY + MaxDither) To YSize - ((GuideBoxY + 1) + MaxDither) Step GuideBoxY * 2
                
                If Aborted Then Exit For
                If XSize - ((GuideBoxX + 1) + MaxDither) - CurrentXPos < (GuideBoxX * 2) Then
                    BoxXStop = XSize - ((GuideBoxX + 1) + MaxDither)
                Else
                    BoxXStop = CurrentXPos + (GuideBoxX * 2)
                End If
                
                If YSize - ((GuideBoxY + 1) + MaxDither) - CurrentYPos < (GuideBoxY * 2) Then
                    BoxYStop = YSize - ((GuideBoxY + 1) + MaxDither)
                Else
                    BoxYStop = CurrentYPos + (GuideBoxY * 2)
                End If
                
                Result = 0
                For BoxXPos = CurrentXPos To BoxXStop
                    For BoxYPos = CurrentYPos To BoxYStop
                        If ImageArray(BoxXPos, BoxYPos) > Result Then
                            Result = ImageArray(BoxXPos, BoxYPos)
                            ResultXPos = BoxXPos
                            ResultYPos = BoxYPos
                        End If
                    Next BoxYPos
                Next BoxXPos
                
                'Check if the max pixel is below the max value
                If Result > MaxVal And Result < MaxADU Then
                    'Now search about where I think I found a good guide star
                    'I could be on the skirt of a star, so I want to find the max pixel
                    'Also I could be on the edge of my search area
                    ThisStarVal = 0
                    For XPos = ResultXPos - GuideBoxX To ResultXPos + GuideBoxX
                        For YPos = ResultYPos - GuideBoxX To ResultYPos + GuideBoxX
                            If ThisStarVal < ImageArray(XPos, YPos) Then
                                ThisStarVal = ImageArray(XPos, YPos)
                                ThisStarXPos = XPos
                                ThisStarYPos = YPos
                            End If
                        Next YPos
                    Next XPos
                End If
                
                'Check if I've already ruled out this star
                IgnoreStar = False
                For IgnoredLocationCounter = 1 To IgnoredLocationsX.Count
                    If (IgnoredLocationsX(IgnoredLocationCounter) = ThisStarXPos) And (IgnoredLocationsY(IgnoredLocationCounter) = ThisStarYPos) Then
                        IgnoreStar = True
                    End If
                Next IgnoredLocationCounter
                
                If (ThisStarXPos < (GuideBoxX + MaxDither)) Or (ThisStarXPos > XSize - ((GuideBoxX + 1) + MaxDither)) Or (ThisStarYPos < (GuideBoxY + MaxDither)) Or (ThisStarYPos > YSize - ((GuideBoxY + 1) + MaxDither)) Then
                    'outside my search range
                    IgnoreStar = True
                End If
                                
                'Since the above search could have found a brighter pixel (if I was on the edge of my search box),
                'check against the max value again
                If ThisStarVal > MaxVal And ThisStarVal < MaxADU And Not IgnoreStar Then
                    AverageADU = 0
                    'Look for brightest pixel, and compute Average
                    For XPos = ThisStarXPos - GuideBoxX To ThisStarXPos + GuideBoxX
                        For YPos = ThisStarYPos - GuideBoxY To ThisStarYPos + GuideBoxY
                            AverageADU = AverageADU + ImageArray(XPos, YPos)
                        Next YPos
                    Next XPos
                    AverageADU = AverageADU / CDbl((GuideBoxX * 2 + 1) * (GuideBoxY * 2 + 1))
                    
                    'Now compute Standard Deviation of pixels below the average
                    StandardDeviation = 0
                    StdDevCounter = 0
                    For XPos = ThisStarXPos - GuideBoxX To ThisStarXPos + GuideBoxX
                        For YPos = ThisStarYPos - GuideBoxY To ThisStarYPos + GuideBoxY
                            If (ImageArray(XPos, YPos) < AverageADU) Then
                                StandardDeviation = StandardDeviation + (ImageArray(XPos, YPos) - AverageADU) ^ 2
                                StdDevCounter = StdDevCounter + 1
                            End If
                        Next YPos
                    Next XPos
                    
                    StandardDeviation = Sqr(StandardDeviation / CDbl(StdDevCounter - 1))
                    
                    ' Check if 4 StdDev above background
                    If (ThisStarVal > (AverageADU + (StandardDeviation * 4))) Then
                        Call AddToStatus("Found a possible star at " & ThisStarXPos & "," & ThisStarYPos & ".  Max brightness = " & ThisStarVal)
                    
                        If ThisStarVal > MinADU Then
                            'this might be a winner - first check that it isn't a 1-pixel star (if that option is enabled)
                            If frmOptions.chkIgnore1PixelStars.Value = vbChecked Then
                                IgnoreStar = True
                                For XPos = ThisStarXPos - 1 To ThisStarXPos + 1
                                    DoEvents
                                    If Aborted Then Exit For
                                    
                                    For YPos = ThisStarYPos - 1 To ThisStarYPos + 1
                                        DoEvents
                                        If Aborted Then Exit For
                                        
                                        If XPos <> ThisStarXPos Or YPos <> ThisStarYPos Then
                                            If (ImageArray(XPos, YPos) - AverageADU) > ((ThisStarVal - AverageADU) * 0.5) Then
                                                'Not a 1 pixel star!
                                                IgnoreStar = False
                                            End If
                                        End If
                                    Next YPos
                                Next XPos
                            Else
                                IgnoreStar = False
                            End If
                                                    
                            If Not IgnoreStar Then
                                'now make sure there aren't any others in the same field
                                'look in a 32x32 box surrounding this object to see if there is one of similar brightness
                                MaxFieldVal = 0
                                For XPos = ThisStarXPos - GuideBoxX To ThisStarXPos + GuideBoxX
                                    DoEvents
                                    If Aborted Then Exit For
                                    
                                    For YPos = ThisStarYPos - GuideBoxY To ThisStarYPos + GuideBoxY
                                        DoEvents
                                        If Aborted Then Exit For
                                        
                                        If Abs(XPos - ThisStarXPos) > Settings.GuideStarFWHM Or Abs(YPos - ThisStarYPos) > Settings.GuideStarFWHM Then
                                            If XPos >= 0 And YPos >= 0 And XPos < XSize And YPos < YSize Then
                                                If ImageArray(XPos, YPos) > MaxFieldVal Then
                                                    MaxFieldVal = ImageArray(XPos, YPos)
                                                End If
                                            End If
                                        End If
                                    Next YPos
                                Next XPos
                                
                                If (ThisStarVal - AverageADU) >= ((MaxFieldVal - AverageADU) * 1.25) Then
                                    'good one - record it
                                    Call AddToStatus("Star is good.  Recording...")
                                    MaxVal = ThisStarVal
                                    MaxImageX = ThisStarXPos
                                    MaxImageY = ThisStarYPos
                                Else
                                    Call AddToStatus("Brighter star in field.  Ignoring star.")
                                End If
                            Else
                                Call AddToStatus("Star FWHM < 1 pixel. Ignoring.")
                                IgnoredLocationsX.Add ThisStarXPos
                                IgnoredLocationsY.Add ThisStarYPos
                            End If
                        Else
                            'not a final star, but record just in case it is the brightest
                            Call AddToStatus("Not bright enough, but brightest so far.")
                            MaxVal = ThisStarVal
                            MaxImageX = ThisStarXPos
                            MaxImageY = ThisStarYPos
                        End If
                    End If
                End If
            Next CurrentYPos
        Next CurrentXPos
    End If
    
    If Not Aborted Then
        FindGuideStar = MaxVal
    Else
        FindGuideStar = 0
    End If
End Function

Public Function FindGuideStarOld(MinADU As Long, MaxADU As Long, GuideBoxX As Long, GuideBoxY As Long, MaxDither As Double, MaxImageX As Long, MaxImageY As Long) As Long
    Dim GuiderImage As Object
    Dim InventoryIndex As Long
    Dim MaxImageVal As Long
    Dim XPos As Long
    Dim YPos As Long
    Dim MaxFieldVal As Long
    Dim ThisStarVal As Long
    Dim CentroidX As Double
    Dim CentroidY As Double
    Dim ImageWidth As Long
    Dim ImageHeight As Long
    Dim LocalImage As Variant
    Dim IgnoreStar As Boolean
    
    On Error Resume Next
    'This will remove any soultion files
    Call Kill(App.Path & "\*.SRC")
    On Error GoTo 0
    
    Call Wait(1)
    
    Set GuiderImage = CreateObject("CCDSoft.Image")
    
    If MainMod.UseFakeGuideExposure Then
        GuiderImage.Path = App.Path & "\GuiderTest.fit"
        Call GuiderImage.Open
    Else
        Call GuiderImage.AttachToActiveAutoguider
        GuiderImage.Path = App.Path & "\Guider.fit"
        Call GuiderImage.Save
    End If

    Call GuiderImage.ShowInventory
    
    MaxImageVal = 0
    
    'Moved these out of the for loop
    With GuiderImage
        ImageWidth = .Width
        ImageHeight = .Height
        LocalImage = .DataArray
    End With
    
    For InventoryIndex = 0 To UBound(GuiderImage.InventoryArray(cdInventoryX))
        DoEvents
        If Aborted Then Exit For
        
        'Check to see if it is over the minimum value and under the maximum value
        'also check around the star centroid
        With GuiderImage
            CentroidX = .InventoryArray(cdInventoryX)(InventoryIndex)
            CentroidY = .InventoryArray(cdInventoryY)(InventoryIndex)
                        
#If ExtraGuideStarDebug Then
            Call AddToStatus("Checking Guide Star @ " & CentroidX & "," & CentroidY)
#End If
        End With
        
        'star needs to be at least 16 pixels away from the edge
        'if not, the autoguider will bring the star in to this position - this could cause problems
        'unfortunatly I cannot think of a way to deal with this, so ignore these stars close to the edge
        'update! Can change the size of the guide box in CCDSoft.
        'pull the actual value from the setup form
        If CentroidX >= GuideBoxX And CentroidX <= (ImageWidth + 1 - GuideBoxX) And CentroidY >= GuideBoxY And CentroidY <= (ImageHeight + 1 - GuideBoxY) Then
            ThisStarVal = LocalImage(CentroidX, CentroidY)
            For XPos = CentroidX - 1 To CentroidX + 1
                DoEvents
                If Aborted Then Exit For
                
                For YPos = CentroidY - 1 To CentroidY + 1
                    DoEvents
                    If Aborted Then Exit For
                    
                    If XPos >= 0 And YPos >= 0 And XPos < GuiderImage.Width And YPos < GuiderImage.Height Then
                        If ThisStarVal < LocalImage(XPos, YPos) Then ThisStarVal = LocalImage(XPos, YPos)
                    End If
                Next YPos
            Next XPos
                
            If Aborted Then Exit For
            
#If ExtraGuideStarDebug Then
            Call AddToStatus("Guide Star max brightness = " & ThisStarVal)
#End If

            If ThisStarVal > MaxImageVal And ThisStarVal < MaxADU Then
                Call AddToStatus("Found a possible star at " & CentroidX & "," & CentroidY & ".  Max brightness = " & ThisStarVal)
                
                If ThisStarVal > MinADU Then
                    'this might be a winner - first check that it isn't a 1-pixel star (if that option is enabled)
                    If frmOptions.chkIgnore1PixelStars.Value = vbChecked Then
                        IgnoreStar = True
                        For XPos = CentroidX - 1 To CentroidX + 1
                            DoEvents
                            If Aborted Then Exit For
                            
                            For YPos = CentroidY - 1 To CentroidY + 1
                                DoEvents
                                If Aborted Then Exit For
                                
                                If XPos <> CentroidX Or YPos <> CentroidY Then
                                    If LocalImage(XPos, YPos) > (ThisStarVal * 0.5) Then
                                        'Not a 1 pixel star!
                                        IgnoreStar = False
                                    End If
                                End If
                            Next YPos
                        Next XPos
                    Else
                        IgnoreStar = False
                    End If
                                            
                    If Not IgnoreStar Then
                        'this might be a winner - make sure there aren't any others in the same field
                        'look in a 32x32 box surrounding this object to see if there is one of similar brightness
                        'update - use the specified size instead of the 32x32 box.
                        MaxFieldVal = 0
                        For XPos = CentroidX - GuideBoxX To CentroidX + GuideBoxX
                            DoEvents
                            If Aborted Then Exit For
                            
                            For YPos = CentroidY - GuideBoxY To CentroidY + GuideBoxY
                                DoEvents
                                If Aborted Then Exit For
                                
                                If Abs(XPos - CentroidX) > Settings.GuideStarFWHM Or Abs(YPos - CentroidY) > Settings.GuideStarFWHM Then
                                    'ignore the pixels around the object in question - don't come in here
                                    'check if there is an object brighter or 25% dimmer then the object in question
                                    If XPos >= 0 And YPos >= 0 And XPos < ImageWidth And YPos < ImageHeight Then
                                        If LocalImage(XPos, YPos) > MaxFieldVal Then
                                            MaxFieldVal = LocalImage(XPos, YPos)
                                        End If
                                    End If
                                End If
                            Next YPos
                        Next XPos
                        
                        If ThisStarVal >= (MaxFieldVal * 1.25) Then
                            Call AddToStatus("Star is good.  Recording...")
                            'good one - record it
                            MaxImageVal = ThisStarVal
                            MaxImageX = CentroidX
                            MaxImageY = CentroidY
                            
                            'exit out of InventoryIndex loop - no need to check any others since this one will do
                            Exit For
                        Else
                            Call AddToStatus("Brighter star in field.  Ignoring star.")
                        End If
                    Else
                        Call AddToStatus("Star FWHM < 1 pixel.  Ignoring.")
                    End If
                Else
                    'not a final star, but record just in case it is the brightest
                    Call AddToStatus("Not bright enough, but brightest so far.")
                    MaxImageVal = ThisStarVal
                    MaxImageX = CentroidX
                    MaxImageY = CentroidY
                End If
            Else
#If ExtraGuideStarDebug Then
                Call AddToStatus("Star is either too bright, or dimmer than the best star detected thus far.")
#End If
            End If
        Else
#If ExtraGuideStarDebug Then
            Call AddToStatus("Guide Star to close to edge.")
#End If
        End If
    Next InventoryIndex
    
    FindGuideStarOld = MaxImageVal
End Function

Public Function GetFilters() As Variant
    Dim FilterArray() As String
    Dim Counter As Integer
    
    On Error Resume Next
    
    If objImager Is Nothing Then
        Set objImager = New CCDSoftCamera
    End If
    
    objImager.Connect
    
    If (objImager.lNumberFilters < 1) Or (Err.Number <> 0) Then
        Call MsgBox("Error connecting to the camera or filter wheel.")
        GetFilters = Null
    Else
        ReDim FilterArray(objImager.lNumberFilters - 1)
        For Counter = 0 To objImager.lNumberFilters - 1
            FilterArray(Counter) = objImager.szFilterName(Counter)
        Next Counter
        
        GetFilters = FilterArray()
    End If
    
    On Error GoTo 0
End Function

Public Property Get GuideErrorX() As Double
    GuideErrorX = objAutoguider.GuideErrorX
End Property

Public Property Get GuideErrorY() As Double
    GuideErrorY = objAutoguider.GuideErrorY
End Property

Public Property Get GuiderRunning() As Boolean
    If (objAutoguider.State = cdStateAutoGuide) Then
        GuiderRunning = True
    Else
        GuiderRunning = False
    End If
End Property

Public Property Get GuideStarX() As Double
    If objAutoguider Is Nothing Then
        Call Me.ConnectToCamera
    End If
    
    If AutoguiderEnabled Then
        GuideStarX = objAutoguider.GuideStarX / objAutoguider.BinX
    End If
End Property

Public Property Get GuideStarY() As Double
    If objAutoguider Is Nothing Then
        Call Me.ConnectToCamera
    End If
    
    If AutoguiderEnabled Then
        GuideStarY = objAutoguider.GuideStarY / objAutoguider.BinY
    End If
End Property

Public Property Get HeightInPixels() As Long
    HeightInPixels = objImager.HeightInPixels
End Property

Public Property Let ImageReduction(Value As ReductionType)
    If Value = NoReduction Then
        objImager.ImageReduction = cdNone
    ElseIf Value = AutoDark Then
        objImager.ImageReduction = cdAutoDark
    ElseIf Value = FullReduction Then
        objImager.ImageReduction = cdBiasDarkFlat
    End If
End Property

Public Property Get ImageReduction() As ReductionType
    If objImager.ImageReduction = cdNone Then
        ImageReduction = NoReduction
    ElseIf objImager.ImageReduction = cdAutoDark Then
        ImageReduction = AutoDark
    ElseIf objImager.ImageReduction = cdBiasDarkFlat Then
        ImageReduction = FullReduction
    End If
End Property

Public Property Let ImageType(Value As Integer)
    objImager.Frame = Value
End Property

Public Property Get ImageType() As Integer
    ImageType = objImager.Frame
End Property

Public Property Get PixelSize() As Double
    PixelSize = myPixelSize
End Property

Public Sub RecomputeGuiderCalibration(NewAngle As Double, CurrentAngle As Double)
    'CCDSoft stores the calibration for each relay (+X, +Y, -X, -Y) as a pair of X & Y vectors.
    'So, I need to read the current vector values, convert to Magnitude and Angle, rotate to the new angle,
    'and then convert back to X & Y vectors.
    
    'CCDSoft calibration angles are measured with 0 degrees at the +X (on the detector) and goes counter clockwise.
    'The position angle goes clockwise with 0 degrees at North, like a compass.
    'So, I compute the change in angle (NewAngle - CurrentAngle) and then invert this since CCDSoft's angles are reverse.
    
    Dim VectorX As Double
    Dim VectorY As Double
    
    Dim RGH As Boolean
    
    If objAutoguider.PropLng("m_bSBIGUseExternalCCD") = 0 Then
        RGH = False
    Else
        RGH = True
    End If
    
    'First do +X
    If Not RGH Then
        VectorX = objAutoguider.CalibrationVectorXPositiveXComponent
        VectorY = objAutoguider.CalibrationVectorXPositiveYComponent
    Else
        VectorX = objAutoguider.PropDbl("m_CRExternalCCD.vXPlus.dXComponent")
        VectorY = objAutoguider.PropDbl("m_CRExternalCCD.vXPlus.dYComponent")
    End If
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    If Not RGH Then
        objAutoguider.CalibrationVectorXPositiveXComponent = VectorX
        objAutoguider.CalibrationVectorXPositiveYComponent = VectorY
    Else
        objAutoguider.PropDbl("m_CRExternalCCD.vXPlus.dXComponent") = VectorX
        objAutoguider.PropDbl("m_CRExternalCCD.vXPlus.dYComponent") = VectorY
    End If
    VectorX = objImager.CalibrationVectorXPositiveXComponent
    VectorY = objImager.CalibrationVectorXPositiveYComponent
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    objImager.CalibrationVectorXPositiveXComponent = VectorX
    objImager.CalibrationVectorXPositiveYComponent = VectorY

    'Now -X
    If Not RGH Then
        VectorX = objAutoguider.CalibrationVectorXNegativeXComponent
        VectorY = objAutoguider.CalibrationVectorXNegativeYComponent
    Else
        VectorX = objAutoguider.PropDbl("m_CRExternalCCD.vXMinus.dXComponent")
        VectorY = objAutoguider.PropDbl("m_CRExternalCCD.vXMinus.dYComponent")
    End If
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    If Not RGH Then
        objAutoguider.CalibrationVectorXNegativeXComponent = VectorX
        objAutoguider.CalibrationVectorXNegativeYComponent = VectorY
    Else
        objAutoguider.PropDbl("m_CRExternalCCD.vXMinus.dXComponent") = VectorX
        objAutoguider.PropDbl("m_CRExternalCCD.vXMinus.dYComponent") = VectorY
    End If
    VectorX = objImager.CalibrationVectorXNegativeXComponent
    VectorY = objImager.CalibrationVectorXNegativeYComponent
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    objImager.CalibrationVectorXNegativeXComponent = VectorX
    objImager.CalibrationVectorXNegativeYComponent = VectorY

    '+Y
    If Not RGH Then
        VectorX = objAutoguider.CalibrationVectorYPositiveXComponent
        VectorY = objAutoguider.CalibrationVectorYPositiveYComponent
    Else
        VectorX = objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dXComponent")
        VectorY = objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dYComponent")
    End If
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    If Not RGH Then
        objAutoguider.CalibrationVectorYPositiveXComponent = VectorX
        objAutoguider.CalibrationVectorYPositiveYComponent = VectorY
    Else
        objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dXComponent") = VectorX
        objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dYComponent") = VectorY
    End If
    VectorX = objImager.CalibrationVectorYPositiveXComponent
    VectorY = objImager.CalibrationVectorYPositiveYComponent
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    objImager.CalibrationVectorYPositiveXComponent = VectorX
    objImager.CalibrationVectorYPositiveYComponent = VectorY

    '-Y
    If Not RGH Then
        VectorX = objAutoguider.CalibrationVectorYNegativeXComponent
        VectorY = objAutoguider.CalibrationVectorYNegativeYComponent
    Else
        VectorX = objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dXComponent")
        VectorY = objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dYComponent")
    End If
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    If Not RGH Then
        objAutoguider.CalibrationVectorYNegativeXComponent = VectorX
        objAutoguider.CalibrationVectorYNegativeYComponent = VectorY
    Else
        objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dXComponent") = VectorX
        objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dYComponent") = VectorY
    End If
    VectorX = objImager.CalibrationVectorYNegativeXComponent
    VectorY = objImager.CalibrationVectorYNegativeYComponent
    If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(CurrentAngle - NewAngle, 360))
    Else
        Call RotateVector(VectorX, VectorY, Misc.DoubleModulus(NewAngle - CurrentAngle, 360))
    End If
    objImager.CalibrationVectorYNegativeXComponent = VectorX
    objImager.CalibrationVectorYNegativeYComponent = VectorY

    'that should be it!
End Sub

Public Property Let ReverseX(Value As Boolean)
'    objAutoguider.ReverseX = Value
    'If objAutoguider.MoveVia <> cdAutoguideViaDirectGuide Then
    If (InStr(Mount.GetTelescopeType(), "Paramount ME") = 0) Then
        If Value Then
            objAutoguider.ReverseX = 1
        Else
            objAutoguider.ReverseX = 0
        End If
    End If
End Property

Public Property Get ReverseX() As Boolean
    If objAutoguider.ReverseX = 1 Then
        ReverseX = True
    Else
        ReverseX = False
    End If
End Property

Public Property Get ReverseXNecessary() As Boolean
    If objAutoguider.MoveVia <> cdAutoguideViaDirectGuide Then
        ReverseXNecessary = True
    Else
        ReverseXNecessary = False
    End If
End Property

Public Sub ReverseYGuiderDirections()
    Dim VectorX As Double
    Dim VectorY As Double
    
    Dim RGH As Boolean
    
    If objAutoguider.PropLng("m_bSBIGUseExternalCCD") = 0 Then
        RGH = False
    Else
        RGH = True
    End If
    
    'This just needs to reverse the +Y and -Y vectors.
    'Reversing is the same as rotating 180 degrees!
    
    '+Y
    If Not RGH Then
        VectorX = objAutoguider.CalibrationVectorYPositiveXComponent
        VectorY = objAutoguider.CalibrationVectorYPositiveYComponent
    Else
        VectorX = objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dXComponent")
        VectorY = objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dYComponent")
    End If
    Call RotateVector(VectorX, VectorY, 180)
    If Not RGH Then
        objAutoguider.CalibrationVectorYPositiveXComponent = VectorX
        objAutoguider.CalibrationVectorYPositiveYComponent = VectorY
    Else
        objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dXComponent") = VectorX
        objAutoguider.PropDbl("m_CRExternalCCD.vYPlus.dYComponent") = VectorY
    End If
    VectorX = objImager.CalibrationVectorYPositiveXComponent
    VectorY = objImager.CalibrationVectorYPositiveYComponent
    Call RotateVector(VectorX, VectorY, 180)
    objImager.CalibrationVectorYPositiveXComponent = VectorX
    objImager.CalibrationVectorYPositiveYComponent = VectorY

    '-Y
    If Not RGH Then
        VectorX = objAutoguider.CalibrationVectorYNegativeXComponent
        VectorY = objAutoguider.CalibrationVectorYNegativeYComponent
    Else
        VectorX = objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dXComponent")
        VectorY = objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dYComponent")
    End If
    Call RotateVector(VectorX, VectorY, 180)
    If Not RGH Then
        objAutoguider.CalibrationVectorYNegativeXComponent = VectorX
        objAutoguider.CalibrationVectorYNegativeYComponent = VectorY
    Else
        objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dXComponent") = VectorX
        objAutoguider.PropDbl("m_CRExternalCCD.vYMinus.dYComponent") = VectorY
    End If
    VectorX = objImager.CalibrationVectorYNegativeXComponent
    VectorY = objImager.CalibrationVectorYNegativeYComponent
    Call RotateVector(VectorX, VectorY, 180)
    objImager.CalibrationVectorYNegativeXComponent = VectorX
    objImager.CalibrationVectorYNegativeYComponent = VectorY
End Sub

Public Sub SaveImagerImage(FileName As String)
    Dim objImage As CCDSoftLib.Image 'Object
    
    Set objImage = CreateObject("CCDSoft.Image")
    'save the image
    Call AddToStatus("Saving Image...")
    objImage.AttachToActiveImager
    
    objImage.FITSKeyword("OBJCTRA") = Misc.FormatRAForFITSHeader(Mount.CurrentRA)
    objImage.FITSKeyword("OBJCTDEC") = Misc.FormatDecForFITSHeader(Mount.CurrentDec)
    objImage.FITSKeyword("OBJECT") = Mount.CurrentTargetName
    
    myPixelSize = objImage.FITSKeyword("XPIXSZ")
    
    objImage.Path = FileName
    Call objImage.Save
    
    Call objImage.Close
End Sub

Public Sub StartAutoguider()
    If frmOptions.chkDisableDecComp = vbChecked Then
        objAutoguider.TelescopeDeclination = 0
    Else
        objAutoguider.TelescopeDeclination = Mount.CurrentDec
    End If
    Call objAutoguider.Autoguide
    
    If objAutoguider.PropLng("m_bSBIGUseExternalCCD") <> 0 Then
        'Special code to help RGH shutter problem
        Call Wait(5)
        Call Me.StopAutoguider
        Call objAutoguider.Autoguide
    End If
End Sub

Public Sub StopAutoguider()
    Dim Counter1 As Integer
    Dim Counter2 As Integer
    
    Counter2 = 0
    Do
        On Error Resume Next
        objAutoguider.Abort
        On Error GoTo 0
    
        Counter1 = 0
        Do Until (objAutoguider.State = cdStateNone) Or (Counter1 > 15)
            Call Wait(1)
            Counter1 = Counter1 + 1
        Loop
        
        If Counter1 > 15 And Counter2 < 4 Then
            Call AddToStatus("Problem stopping autoguider, retrying")
        ElseIf Counter1 > 15 And Counter2 > 4 Then
            Call AddToStatus("Could not stop autoguider.")
        End If
        
        Counter2 = Counter2 + 1
    Loop While Counter1 > 15 And Counter2 < 5
End Sub

Public Sub SubFrame(OnOff As Boolean, Optional Left As Long, Optional Right As Long, Optional Top As Long, Optional Bottom As Long)
    If OnOff = True Then
        objImager.SubFrame = 1
        objImager.SubframeLeft = Left * objImager.BinX
        objImager.SubframeRight = Right * objImager.BinX
        objImager.SubframeTop = Top * objImager.BinY
        objImager.SubframeBottom = Bottom * objImager.BinY
    Else
        objImager.SubFrame = 0
    End If
End Sub

Public Sub TakeDummyImage(Filter As Integer)
    Dim SavedExposureTime As Double
    Dim SavedImageReduction As ReductionType
    Dim SavedAutosaveState As Boolean
    Dim objImage As CCDSoftLib.Image
    Dim TakeImageRetry As Boolean
    Dim RetryCounter As Integer
    
    If objImager.State <> cdStateNone Then
        'Something wrong, imager should be idle
        'Try aborting the camera
        On Error Resume Next
        objImager.Abort
        On Error GoTo 0
    End If
    
    SavedAutosaveState = Me.AutoSaveState
    Call Me.AutoSave(False)
    
    objImager.Series = 1
    If Filter >= 0 Then
        objImager.FilterIndexZeroBased = Filter
    End If
    
    Call SubFrame(True, 0, 2, 0, 2)
    
    SavedExposureTime = objImager.ExposureTime
    
    SavedImageReduction = Me.ImageReduction
    
    objImager.ExposureTime = 0.11
    RetryCounter = 0
    Do
        On Error Resume Next
        Call objImager.TakeImage
        If Err.Number = -2147220285 Then
            'Command in progress error - try aborting and try waiting a bit before retrying
            On Error GoTo 0
            objImager.Abort
            Call Wait(10)
            TakeImageRetry = True
            RetryCounter = RetryCounter + 1
        ElseIf Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        Else
            On Error GoTo 0
            TakeImageRetry = False
        End If
    Loop While TakeImageRetry = True
    
    Do Until objImager.State = cdStateNone Or Aborted
        Call Wait(0.1)
    Loop
        
    Call Wait(1)
    'close the dummy image
    Set objImage = CreateObject("CCDSoft.Image")
    'get the handle to the image
    objImage.AttachToActive
    'objImage.Visible = False
    objImage.Close
    Set objImage = Nothing
        
    Call SubFrame(False)
    objImager.ExposureTime = SavedExposureTime
    Me.ImageReduction = SavedImageReduction
    Call Me.AutoSave(SavedAutosaveState)
End Sub

Public Sub TakeImage()
    objImager.TakeImage
    StartedImage = True
End Sub

Public Property Get TakeImageComplete() As Boolean
    Dim objImage As Object
    Dim SequenceNum As Long
    Dim Alt As Double
    Dim Az As Double
    Dim ImagePath As String
    Dim ImagePathOnly As String
    Dim ImageName As String
    Dim StartPos As Long
    Dim Counter As Integer
    Dim Done As Boolean
    
    If objImager.State = cdStateNone Then
        TakeImageComplete = True
    
        If StartedImage = True Then
            Call Wait(1)
            StartedImage = False
            If myAutosaveState = True Then
                Set objImage = CreateObject("CCDSoft.Image")
                
                'get the handle to the image
                objImage.AttachToActiveImager
                
                'setup the FITS headers
                If (Mount.CurrentTargetJ2000) Then
                    objImage.FITSKeyword("OBJCTRA") = Misc.FormatRAForFITSHeader(Mount.CurrentTargetJ2000RA)
                    objImage.FITSKeyword("OBJCTDEC") = Misc.FormatDecForFITSHeader(Mount.CurrentTargetJ2000Dec)
                Else
                    objImage.FITSKeyword("OBJCTRA") = Misc.FormatRAForFITSHeader(Mount.CurrentRA)
                    objImage.FITSKeyword("OBJCTDEC") = Misc.FormatDecForFITSHeader(Mount.CurrentDec)
                End If
                objImage.FITSKeyword("OBJECT") = Mount.CurrentTargetName
                Call Mount.GetTelescopeAltAz(Alt, Az)
                objImage.FITSKeyword("CENTAZ") = Format(Az, "0.0000")
                objImage.FITSKeyword("CENTALT") = Format(Alt, "0.0000")
                objImage.FITSKeyword("HISTORY") = "Automated by CCD Commander"
                
                'save the image again with the headers
                On Error Resume Next
                Counter = 0
                Done = False
                Do
                    Call objImage.Save
                    If Err.Number <> 0 Then
                        Counter = Counter + 1
                        Call Wait(1)
                    Else
                        Done = True
                    End If
                Loop Until Done Or Counter > 10
                On Error GoTo 0
                
                If Counter > 10 Then
                    Call AddToStatus("Error saving image in CCDSoft.")
                    Exit Property
                End If
                
                ImagePath = objImage.Path
                
                'extract path and file name
                ImagePathOnly = Left(ImagePath, InStrRev(ImagePath, "\"))
                ImageName = Mid(ImagePath, InStrRev(ImagePath, "\") + 1)
                
                Call objImage.Close
                
                If Me.ImageReduction = FullReduction Then
                    'I need to do all of this again, but with the RAW image now
                    
                    'Get a new handle to an image
                    Set objImage = CreateObject("CCDSoft.Image")
                    
                    'Replace the REDUCED text with RAW to get the file name
                    StartPos = InStr(ImagePath, "REDUCED")
                    Mid(ImagePath, StartPos) = "RAW"
                    Mid(ImagePath, StartPos + 3) = Mid(ImagePath, StartPos + 7) & "    "
                    
                    objImage.Path = ImagePath
                    objImage.Open
                    
                    'setup the FITS headers
                    objImage.FITSKeyword("OBJCTRA") = Misc.FormatRAForFITSHeader(Mount.CurrentRA)
                    objImage.FITSKeyword("OBJCTDEC") = Misc.FormatDecForFITSHeader(Mount.CurrentDec)
                    objImage.FITSKeyword("OBJECT") = Mount.CurrentTargetName
                    Call Mount.GetTelescopeAltAz(Alt, Az)
                    objImage.FITSKeyword("CENTAZ") = Format(Az, "0.0000")
                    objImage.FITSKeyword("CENTALT") = Format(Alt, "0.0000")
                    objImage.FITSKeyword("HISTORY") = "Automated by CCD Commander"
                    
                    'save the image again with the headers
                    Call objImage.Save
                
                    Call objImage.Close
                    
                    'Finally move the reduced image into a sub-folder
                    On Error Resume Next
                    Call MkDir(ImagePathOnly & "Reduced")
                    On Error GoTo 0
                
                    'copy to the reduced folder
                    Call FileCopy(ImagePathOnly & ImageName, ImagePathOnly & "Reduced\" & ImageName)
                    'delete the reduced file in the original location
                    Call Kill(ImagePathOnly & ImageName)
                End If
            End If
        End If
    Else
        TakeImageComplete = False
    End If
End Property

Public Property Let Temperature(Value As Double)
    objImager.TemperatureSetPoint = Value
    Me.CoolerState = True
End Property

Public Property Get Temperature() As Double
    On Error Resume Next
    Temperature = objImager.Temperature
    On Error GoTo 0
End Property

Public Property Get TemperatureSetPoint() As Double
    TemperatureSetPoint = objImager.TemperatureSetPoint
End Property

Public Sub UseGuideStarAt(X As Double, Y As Double)
    objAutoguider.GuideStarX = CLng(X) * objAutoguider.BinX
    objAutoguider.GuideStarY = CLng(Y) * objAutoguider.BinY
End Sub

Public Property Get WidthInPixels() As Long
    WidthInPixels = objImager.WidthInPixels
End Property

Public Property Get MultiStarGuidingEnabled() As Boolean
    MultiStarGuidingEnabled = False
End Property

Public Property Let MultiStarGuidingEnabled(Value As Boolean)
    'Nothing to do
End Property

