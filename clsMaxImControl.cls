VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaxImControl"
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
 
Private WithEvents objMaxImImager As MaxIm.CCDCamera
Attribute objMaxImImager.VB_VarHelpID = -1
Private myAGExposureTime As Double
Private myAutosaveState As Boolean
Private myAutoSavePath As String
Private myAutosaveFN As String
Private myExposureTime As Double
Private myImageType As Integer
Private myXYBinning As Boolean
Private myGuiderX As Double
Private myGuiderY As Double
Private myReductionType As ReductionType
Private StartedImage As Boolean
Private ExposureComplete As Boolean
Private AutoguiderEnabled As Boolean
Private myPixelSize As Double

Private IAmAborting As Boolean

Private Version As Single

Public Sub Abort()
    On Error Resume Next
    IAmAborting = True
    objMaxImImager.AbortExposure
    If AutoguiderEnabled Then
        Do
            IAmAborting = True
            objMaxImImager.GuiderStop
        Loop While objMaxImImager.GuiderRunning
    End If
    On Error GoTo 0
    
    ExposureComplete = False
    StartedImage = False
End Sub

Public Property Let AGBinXY(Value As Integer)
    On Error GoTo AGBinXYError
    If (Version >= 5.08) Then
        objMaxImImager.GuiderBinning = Value
    End If
    
    On Error GoTo 0
    Exit Property
    
AGBinXYError:
    Call AddToStatus("MaxIm AGBinXY Error.  Version = " & Version)
    On Error GoTo 0
End Property

Public Property Let AGExposureTime(Value As Double)
    If AutoguiderEnabled Then
        myAGExposureTime = Value
    End If
End Property

Public Property Get AGExposureTime() As Double
    If AutoguiderEnabled Then
        AGExposureTime = myAGExposureTime
    End If
End Property

Public Property Let AGFilterNumber(Value As Integer)
    If AutoguiderEnabled Then
        objMaxImImager.Filter = Value
    End If
End Property

Public Property Let AGImageType(Value As Integer)
    'not supported
End Property

Public Property Let AGReduction(Value As Integer)
    'not supported
End Property

Public Sub AGTakeImage()
    Dim ExposureSuccessful As Boolean
    Dim Counter As Integer
    
    If AutoguiderEnabled Then
        Counter = 0
        Do
            On Error Resume Next
            Call objMaxImImager.GuiderExpose(myAGExposureTime)
            If Err.Number = 65535 Then
                On Error GoTo 0
                Call AddToStatus("Error taking autoguider image - retrying.")
                objMaxImImager.GuiderStop
                ExposureSuccessful = False
                Counter = Counter + 1
            Else
                ExposureSuccessful = True
            End If
        Loop Until ExposureSuccessful Or Counter > 5
    End If
End Sub

Public Property Get AGTakeImageComplete() As Boolean
    If AutoguiderEnabled Then
        AGTakeImageComplete = Not objMaxImImager.GuiderRunning
    Else
        AGTakeImageComplete = True
    End If
End Property

Public Property Get AutoguiderInterval() As Double
    If AutoguiderEnabled Then
        AutoguiderInterval = myAGExposureTime
    End If
End Property

Public Sub AutoSave(OnOff As Boolean, Optional Path As String = "*", Optional FileName As String = "*")
    myAutosaveState = OnOff
    
    If Path <> "*" Then
        myAutoSavePath = Path
    End If
    
    If FileName <> "*" Then
        myAutosaveFN = FileName
    End If
    
'    If myAutosaveState Then
        'objMaxImImager.EventMask = objMaxImImager.EventMask Or (2 ^ MaxIm.ceExposureCompleted)
'    Else
'        objMaxImImager.EventMask = objMaxImImager.EventMask And (Not (2 ^ MaxIm.ceExposureCompleted))
'    End If
End Sub

Public Property Get AutoSaveState() As Boolean
    AutoSaveState = myAutosaveState
End Property

Public Property Get AverageADUOfExposure(FrameSize As FrameSizes) As Double
    Dim objDoc As MaxIm.Document
    Dim Result As Variant
    Dim XStart As Long
    Dim YStart As Long
    Dim YStop As Long
    Dim XStop As Long
    
    If MainMod.UseFakeFlatImage Then
        Set objDoc = New MaxIm.Document
        Call objDoc.OpenFile(App.Path & "\FlatTest.fit")
    Else
        Set objDoc = objMaxImImager.Document
    End If
    
    If FrameSize = FullFrame Then
        XStart = 0
        YStart = 0
        XStop = objDoc.XSize
        YStop = objDoc.YSize
    ElseIf FrameSize = HalfFrame Then
        XStart = objDoc.XSize / 4
        YStart = objDoc.YSize / 4
        XStop = (objDoc.XSize * 3) / 4
        YStop = (objDoc.YSize * 3) / 4
    ElseIf FrameSize = QuarterFrame Then
        XStart = (objDoc.XSize * 3) / 8
        YStart = (objDoc.YSize * 3) / 8
        XStop = (objDoc.XSize * 5) / 8
        YStop = (objDoc.YSize * 5) / 8
    End If
    
    Result = objDoc.CalcAreaInfo(XStart, YStart, XStop - 1, YStop - 1, 0)
    
    AverageADUOfExposure = Result(2)
End Property

Public Property Get BinX() As Integer
    BinX = objMaxImImager.BinX
End Property

Public Property Let BinXY(Value As Integer)
    If myXYBinning Then
        objMaxImImager.BinX = Value
        objMaxImImager.BinY = Value
    Else
        objMaxImImager.BinX = Value
    End If
End Property

Public Property Get BinY() As Integer
    BinY = objMaxImImager.BinY
End Property

Public Sub ConnectToCamera()
    Dim MyApp As Object
    Dim TempObj As Object
    
    IAmAborting = False
    
    Set MyApp = CreateObject("MaxIm.Application")
    
    Version = MyApp.Version
    
    If objMaxImImager Is Nothing Then
        Set objMaxImImager = New MaxIm.CCDCamera
    End If
    
    'Connect to the camera
    objMaxImImager.LinkEnabled = True
    Call Wait(1)
    
    If (Not objMaxImImager.LinkEnabled) Then
        Call Err.Raise(65535, "CCDCommander:MaxImControl", "Could not connect to camera.")
    End If
    
    If (frmOptions.chkIgnoreStarFaded.Value = vbChecked) Then
        objMaxImImager.EventMask = 2 ^ MaxIm.ceGuiderExposureCompleted + 2 ^ MaxIm.ceExposureReadoutCompleted + 2 ^ MaxIm.ceExposureAborted
    Else
        objMaxImImager.EventMask = 2 ^ MaxIm.ceGuiderExposureCompleted + 2 ^ MaxIm.ceExposureReadoutCompleted + 2 ^ MaxIm.ceGuiderExposureStarFaded + 2 ^ MaxIm.ceExposureAborted
    End If

    objMaxImImager.AutoDownload = True
    myXYBinning = objMaxImImager.XYBinning
    myAutosaveState = False
    objMaxImImager.DisableAutoShutdown = True
    
    If objMaxImImager.GuiderName = "No Camera" Then
        Call AddToStatus("No autoguider detected.")
        AutoguiderEnabled = False
    Else
        Set TempObj = objMaxImImager
        TempObj.AOMirrorHome = False
        AutoguiderEnabled = True
    End If
    
    myReductionType = NoReduction
    
    On Error GoTo ConnectToCameraError
    If (Version >= 5.08) Then
        If ((objMaxImImager.GuideControlVia = 2) Or (objMaxImImager.GuideControlVia = 3)) Then
            If MyApp.TelescopeConnected = False Then
                MyApp.TelescopeConnected = True
                If MyApp.TelescopeConnected = False Then
                    On Error GoTo 0
                    Call Err.Raise(65534, "CCDCommander:MaxImControl", "Could not connect to telescope.")
                End If
            End If
        End If
    End If
    
    On Error GoTo 0
    Set TempObj = Nothing
    Exit Sub
    
ConnectToCameraError:
    On Error GoTo 0
    Call AddToStatus("MaxIm Error in ConnectToCamera.  Version = " & Version)
End Sub

Public Sub ConnectToTelescope()
    'Not needed anymore since I am writing the RA/Dec to the FITS header and the autoguider
    
    'Dim MyApp As New MaxIm.Application

    'MyApp.TelescopeConnected = True
    'If MyApp.TelescopeConnected = False Then
    '    Call Err.Raise(65534, "CCDCommander:MaxImControl", "Could not connect to telescope.")
    'End If
End Sub

Public Sub DisconnectFromTelescope()
    'Not needed anymore since I am writing the RA/Dec to the FITS header and the autoguider
'    Dim MyApp As New MaxIm.Application
'
'    MyApp.TelescopeConnected = False
End Sub

Public Sub CenterAO()
    Dim TempObj As Object
    Set TempObj = objMaxImImager
    TempObj.AOTiltX = 2048
    TempObj.AOTiltY = 2048
    Call AGTakeImage
    'wait for image
    Do Until AGTakeImageComplete Or Aborted
        Call Wait(1)
    Loop
    Set TempObj = Nothing
End Sub

Public Property Let CoolerState(Value As Boolean)
    Call Wait(1)
    objMaxImImager.CoolerOn = Value
    Call Wait(1)
End Property

Public Property Get CoolerState() As Boolean
    CoolerState = objMaxImImager.CoolerOn
End Property

Public Property Get CoolerPower() As Double
    On Error Resume Next
    CoolerPower = objMaxImImager.CoolerPower
    If Err.Number <> 0 Then
        CoolerPower = -1
    End If
    On Error GoTo 0
End Property

Public Property Get ExposureTime() As Double
    If myImageType = 2 Then
        ExposureTime = 0
    Else
        ExposureTime = myExposureTime
    End If
End Property

Public Property Let ExposureTime(Value As Double)
    myExposureTime = Value
End Property

Public Property Let FanOn(Value As Boolean)
    'not available in MaxIm
    Dim Cam As Object
    
    Set Cam = objMaxImImager
    On Error Resume Next
    Cam.FanEnabled = Value
    If Err.Number <> 0 Then
        On Error GoTo 0
        Call AddToStatus("Fan control not available with your camera.")
        If Value Then
            Call AddToStatus("Unable to turn fan on.")
        Else
            Call AddToStatus("Unable to turn fan off.")
        End If
    Else
        On Error GoTo 0
    End If
    Set Cam = Nothing
End Property

Public Property Get FastDownloadEnabled() As Boolean
    FastDownloadEnabled = objMaxImImager.FastReadout
End Property

Public Property Let FastDownloadEnabled(Value As Boolean)
    objMaxImImager.FastReadout = Value
End Property

Public Property Get FilterNumber() As Integer
    FilterNumber = objMaxImImager.Filter
End Property

Public Property Let FilterNumber(Value As Integer)
    objMaxImImager.Filter = Value
End Property

Public Function FindGuideStar(MinADU As Long, MaxADU As Long, GuideBoxX As Long, GuideBoxY As Long, MaxDither As Double, MaxImageX As Long, MaxImageY As Long) As Long
    Dim CurrentXPos As Integer
    Dim CurrentYPos As Integer
    Dim BoxXStop As Integer
    Dim BoxYStop As Integer
    Dim Rings(0 To 2) As Integer
    Dim Result As Variant
    Dim ResultXPos As Integer
    Dim ResultYPos As Integer
    Dim GuiderImage As MaxIm.Document
    Dim AllDocs As MaxIm.Documents
    Dim MaximApp As New MaxIm.Application
    Dim MaxVal As Variant
    Dim Counter As Long
    Dim XPos As Integer
    Dim YPos As Integer
    Dim ThisStarVal As Variant
    Dim ThisStarXPos As Integer
    Dim ThisStarYPos As Integer
    Dim ImageArray() As Single
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
    
    Rings(0) = 16
    Rings(1) = 0
    Rings(2) = 1
        
    If Not Aborted Then
        Set AllDocs = MaximApp.Documents
    End If
    
    If MainMod.UseFakeGuideExposure Then
        Set GuiderImage = New MaxIm.Document
        Call GuiderImage.OpenFile(App.Path & "\GuiderTest.fit")
    Else
        If Not Aborted Then
            For Each GuiderImage In AllDocs
                If GuiderImage.DisplayName = "Autoguider Image" Or Aborted Then
                    Exit For
                End If
            Next GuiderImage
        End If
    End If
    
    If Not Aborted Then
        ImageArray = GuiderImage.ImageArray
    End If
    
    If Not Aborted Then
        With GuiderImage
            XSize = .XSize
            YSize = .YSize
        End With
    End If
    
    MaxVal = 0
    If Not Aborted Then
        'Guide Star MUST be more than 16 pixels away from the edge
#If ExtraGuideStarDebug Then
        Call AddToStatus("Searching for a guide star from:")
        Call AddToStatus("X = " & (GuideBoxX + MaxDither) & " To " & (XSize - ((GuideBoxX + 1) + MaxDither)))
        Call AddToStatus("Y = " & (GuideBoxY + MaxDither) & " To " & (YSize - ((GuideBoxY + 1) + MaxDither)))
#End If

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
                'Look for brightest pixel, and compute Average
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
#If ExtraGuideStarDebug Then
                    Call AddToStatus("Found something at X = " & XPos & ", Y = " & YPos & ", @ " & Result & " ADU.")
#End If
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
                    
#If ExtraGuideStarDebug Then
                    Call AddToStatus("Actual Star @ X = " & ThisStarXPos & ", Y = " & ThisStarYPos & ", @ " & ThisStarVal & " ADU.")
#End If
                                              
                End If
                                              
                'Check if I've already ruled out this star
                IgnoreStar = False
                For IgnoredLocationCounter = 1 To IgnoredLocationsX.Count
                    If (IgnoredLocationsX(IgnoredLocationCounter) = ThisStarXPos) And (IgnoredLocationsY(IgnoredLocationCounter) = ThisStarYPos) Then
                        IgnoreStar = True
#If ExtraGuideStarDebug Then
                        Call AddToStatus("Star already on Ignore list - ignoring.")
#End If
                    End If
                Next IgnoredLocationCounter
                
                If (ThisStarXPos < (GuideBoxX + MaxDither)) Or (ThisStarXPos > XSize - ((GuideBoxX + 1) + MaxDither)) Or (ThisStarYPos < (GuideBoxY + MaxDither)) Or (ThisStarYPos > YSize - ((GuideBoxY + 1) + MaxDither)) Then
                    'outside my search range
                    IgnoreStar = True
#If ExtraGuideStarDebug Then
                        Call AddToStatus("Star center outside of search area - ignoring.")
#End If
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
                    
#If ExtraGuideStarDebug Then
                    Call AddToStatus("Average ADU = " & AverageADU)
#End If
                    
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
                    
#If ExtraGuideStarDebug Then
                    Call AddToStatus("Standard Deviation of pixels below the average = " & StandardDeviation)
#End If

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
                                
#If ExtraGuideStarDebug Then
                                If IgnoreStar Then
                                    Call AddToStatus("Star looks like only a 1-pixel star - ignoring.")
                                End If
#End If
                                
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
                                
#If ExtraGuideStarDebug Then
                                Call AddToStatus("Brightest location away from star is = " & MaxFieldVal & " ADU.")
#End If
                                
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
                    Else
#If ExtraGuideStarDebug Then
                        Call AddToStatus("Star not bright enough above background.")
#End If
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

Public Function GetFilters() As Variant
    Dim FilterArray() As String
    Dim Counter As Integer
    
    On Error Resume Next
        
    If objMaxImImager Is Nothing Then
        Set objMaxImImager = New MaxIm.CCDCamera
    End If
    
    'Connect to the camera
    objMaxImImager.LinkEnabled = True
    Call Wait(1)
    
    If (Not objMaxImImager.LinkEnabled) Or (Err.Number <> 0) Then
        Call MsgBox("Error connecting to the camera or filter wheel.")
        GetFilters = Null
    Else
        ReDim FilterArray(UBound(objMaxImImager.FilterNames))
        For Counter = 0 To UBound(objMaxImImager.FilterNames)
            FilterArray(Counter) = objMaxImImager.FilterNames(Counter)
        Next Counter
        
        GetFilters = FilterArray()
    End If

    On Error GoTo 0
End Function

Public Property Get GuideErrorX() As Double
    GuideErrorX = objMaxImImager.GuiderXError
End Property

Public Property Get GuideErrorY() As Double
    GuideErrorY = objMaxImImager.GuiderYError
End Property

Public Property Get GuiderRunning() As Boolean
    GuiderRunning = objMaxImImager.GuiderRunning
End Property

Public Property Get GuideStarX() As Double
    If objMaxImImager Is Nothing Then
        Call Me.ConnectToCamera
    End If
    
    If AutoguiderEnabled Then
        GuideStarX = objMaxImImager.GuiderXStarPosition
    End If
End Property

Public Property Get GuideStarY() As Double
    If objMaxImImager Is Nothing Then
        Call Me.ConnectToCamera
    End If
    
    If AutoguiderEnabled Then
        GuideStarY = objMaxImImager.GuiderYStarPosition
    End If
End Property

Public Property Get HeightInPixels() As Long
    HeightInPixels = objMaxImImager.CameraYSize / objMaxImImager.BinY
End Property

Public Property Let ImageType(Value As Integer)
    myImageType = Value
End Property

Public Property Get ImageType() As Integer
    ImageType = myImageType
End Property

Public Property Let ImageReduction(Value As ReductionType)
    myReductionType = Value
End Property

Public Property Get ImageReduction() As ReductionType
    ImageReduction = myReductionType
End Property

Public Property Get PixelSize() As Double
    PixelSize = myPixelSize
End Property

Public Sub RecomputeGuiderCalibration(NewAngle As Double, CurrentAngle As Double)
    Dim NewGuiderAngle As Double
    Dim TempObj As Object
    Dim MirrorImage As Boolean
    Dim WestSideOfMount As Boolean
    
    If objMaxImImager.CameraName = "SBIG w/ AO" Then
        On Error GoTo RecomputeGuiderCalibrationError
        
        Set TempObj = objMaxImImager
        
        WestSideOfMount = frmOptions.optGuiderCal(0).Value
        
        If TempObj.GuiderReverseX = False And TempObj.AOSwapMotorAxes = False Then
            'PA is between 315 and 45 for non-mirrored image
            'PA is between 135 and 235 for mirrored image
            'Give some wiggle room
            If CurrentAngle > 310 Or CurrentAngle < 50 Then
                'no mirror image
                Call AddToStatus("MaxIm settings indicate current PA = 0.")
                MirrorImage = False
            Else
                Call AddToStatus("MaxIm settings indicate current PA = 180.")
                MirrorImage = True
            End If
        ElseIf TempObj.GuiderReverseX = True And TempObj.AOSwapMotorAxes = False Then
            'PA is between 135 and 225 for non-mirrored image
            'PA is between 315 and 45 for mirrored image
            'Give some wiggle room
            If CurrentAngle > 130 And CurrentAngle < 230 Then
                'no mirror image
                Call AddToStatus("MaxIm settings indicate current PA = 180.")
                MirrorImage = False
            Else
                Call AddToStatus("MaxIm settings indicate current PA = 0.")
                MirrorImage = True
            End If
        ElseIf TempObj.GuiderReverseX = False And TempObj.AOSwapMotorAxes = True Then
            'Can't use PA here - Both non-mirrored and mirrored have the same PA with these settings
            'Need to use ReverseY and MountSide
            Call AddToStatus("MaxIm settings indicate current PA = 90.")
            If (((TempObj.GuiderReverseY = True) And (Not WestSideOfMount)) Or ((TempObj.GuiderReverseY = False) And (WestSideOfMount))) Then
                'no mirror image
                MirrorImage = False
            Else
                MirrorImage = True
            End If
        ElseIf TempObj.GuiderReverseX = True And TempObj.AOSwapMotorAxes = True Then
            'Can't use PA here - Both non-mirrored and mirrored have the same PA with these settings
            'Need to use ReverseY and MountSide
            Call AddToStatus("MaxIm settings indicate current PA = 270.")
            If (((TempObj.GuiderReverseY = False) And (Not WestSideOfMount)) Or ((TempObj.GuiderReverseY = True) And (WestSideOfMount))) Then
                'no mirror image
                MirrorImage = False
            Else
                MirrorImage = True
            End If
        End If
        
        'Now just change the settings to match the new angle
        If MirrorImage Then
            Call AddToStatus("Detected mirror image.")
            If Not WestSideOfMount Then
                Call AddToStatus("Guider calibrated on east side of mount (looking at the western sky).")
                If NewAngle > 315 Or NewAngle <= 45 Then
                    Call AddToStatus("Setting MaxIm to PA = 0.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 45 And NewAngle <= 135 Then
                    Call AddToStatus("Setting MaxIm to PA = 90.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = True
                ElseIf NewAngle > 135 And NewAngle <= 225 Then
                    Call AddToStatus("Setting MaxIm to PA = 180.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 225 And NewAngle <= 315 Then
                    Call AddToStatus("Setting MaxIm to PA = 270.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = True
                End If
            Else
                Call AddToStatus("Guider calibrated on west side of mount (looking at the eastern sky).")
                If NewAngle > 315 Or NewAngle <= 45 Then
                    Call AddToStatus("Setting MaxIm to PA = 0.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 45 And NewAngle <= 135 Then
                    Call AddToStatus("Setting MaxIm to PA = 90.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = True
                ElseIf NewAngle > 135 And NewAngle <= 225 Then
                    Call AddToStatus("Setting MaxIm to PA = 180.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 225 And NewAngle <= 315 Then
                    Call AddToStatus("Setting MaxIm to PA = 270.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = True
                End If
            End If
        Else
            If Not WestSideOfMount Then
                Call AddToStatus("Guider calibrated on east side of mount.")
                If NewAngle > 315 Or NewAngle <= 45 Then
                    Call AddToStatus("Setting MaxIm to PA = 0.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 45 And NewAngle <= 135 Then
                    Call AddToStatus("Setting MaxIm to PA = 90.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = True
                ElseIf NewAngle > 135 And NewAngle <= 225 Then
                    Call AddToStatus("Setting MaxIm to PA = 180.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 225 And NewAngle <= 315 Then
                    Call AddToStatus("Setting MaxIm to PA = 270.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = True
                End If
            Else
                Call AddToStatus("Guider calibrated on west side of mount.")
                If NewAngle > 315 Or NewAngle <= 45 Then
                    Call AddToStatus("Setting MaxIm to PA = 0.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 45 And NewAngle <= 135 Then
                    Call AddToStatus("Setting MaxIm to PA = 90.")
                    TempObj.GuiderReverseX = False
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = True
                ElseIf NewAngle > 135 And NewAngle <= 225 Then
                    Call AddToStatus("Setting MaxIm to PA = 180.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = False
                    TempObj.AOSwapMotorAxes = False
                ElseIf NewAngle > 225 And NewAngle <= 315 Then
                    Call AddToStatus("Setting MaxIm to PA = 270.")
                    TempObj.GuiderReverseX = True
                    TempObj.GuiderReverseY = True
                    TempObj.AOSwapMotorAxes = True
                End If
            End If
        End If

    Else
        'MaxIm wants a number from -180 to +180
        If frmOptions.chkGuiderMirrorImage.Value = vbChecked Then
            NewGuiderAngle = Misc.DoubleModulus(objMaxImImager.GuiderAngle + (NewAngle - CurrentAngle), 360)
        Else
            NewGuiderAngle = Misc.DoubleModulus(objMaxImImager.GuiderAngle - (NewAngle - CurrentAngle), 360)
        End If
        If NewGuiderAngle > 180 Then
            NewGuiderAngle = NewGuiderAngle - 360
        End If
        objMaxImImager.GuiderAngle = NewGuiderAngle
    End If
    
    Exit Sub
    
RecomputeGuiderCalibrationError:
    On Error GoTo 0
    Call AddToStatus("This version of MaxIm/DL is not capable of supporting automatic guider rotation.")
    Call AddToStatus("You need to upgrate to MaxIm/DL v4.58 or later to use this feature.")
End Sub

Public Property Let ReverseX(Value As Boolean)
    If AutoguiderEnabled And (objMaxImImager.CameraName <> "SBIG w/ AO" Or frmOptions.lstRotator.ListIndex = RotatorControl.None) Then
        objMaxImImager.GuiderReverseX = Value

        'gvTelescopePulseGuide = 2
        'gvTelescopeMicroGuide = 3
        'These aren't defined in MaxIm v4
        On Error GoTo ReverseXError
        If (Version >= 5.08) Then
            If ((InStr(Mount.GetTelescopeType(), "Paramount ME") <> 0) And _
                ((objMaxImImager.GuideControlVia = 2) Or (objMaxImImager.GuideControlVia = 3))) Then
                objMaxImImager.GuiderReverseY = Value
            End If
        End If
    End If
    On Error GoTo 0
    Exit Property
    
ReverseXError:
    On Error GoTo 0
    Call AddToStatus("MaxIm ReverseX Error.  Version = " & Version)
End Property

Public Property Get ReverseX() As Boolean
    ReverseX = objMaxImImager.GuiderReverseX
End Property

Public Property Get ReverseXNecessary() As Boolean
'gvTelescopePulseGuide = 2
'gvTelescopeMicroGuide = 3
'These aren't defined in MaxIm v4
    ReverseXNecessary = True
    
    On Error GoTo ReverseXNecessaryError
    If (Version >= 5.08) Then
        If (((InStr(Mount.GetTelescopeType(), "Paramount") <> 0) And _
            ((objMaxImImager.GuideControlVia = 2) Or (objMaxImImager.GuideControlVia = 3)))) Or _
            (objMaxImImager.GuiderAutoPierFlip = True) Then
            ReverseXNecessary = False
        End If
    End If
    On Error GoTo 0
    Exit Property
    
ReverseXNecessaryError:
    On Error GoTo 0
    Call AddToStatus("MaxIm ReverseXNecessary Error.  Version = " & Version)
End Property

Public Sub ReverseYGuiderDirections()
    If objMaxImImager.CameraName = "SBIG w/ AO" Then
        objMaxImImager.GuiderReverseY = Not objMaxImImager.GuiderReverseY
    Else
        objMaxImImager.GuiderYSpeed = -objMaxImImager.GuiderYSpeed
    End If
End Sub

Public Sub SaveImagerImage(FileName As String)
    Dim MaxImDoc As MaxIm.Document

    Call AddToStatus("Saving image...")
    Set MaxImDoc = objMaxImImager.Document
    Call MaxImDoc.SetFITSKey("OBJCTRA", Misc.FormatRAForFITSHeader(Mount.CurrentRA))
    Call MaxImDoc.SetFITSKey("OBJCTDEC", Misc.FormatDecForFITSHeader(Mount.CurrentDec))
    Call MaxImDoc.SetFITSKey("OBJECT", Mount.CurrentTargetName)
    
    myPixelSize = MaxImDoc.GetFITSKey("XPIXSZ")

    Set MaxImDoc = Nothing
    Call objMaxImImager.SaveImage(FileName)
End Sub

Public Sub StartAutoguider()
    Dim Counter As Integer
    Dim Guiding As Boolean
    Dim MultiStarGuiding As Boolean
    
    MultiStarGuiding = False
    On Error Resume Next
    MultiStarGuiding = objMaxImImager.MultiStarGuiding
    On Error GoTo 0
    
    If AutoguiderEnabled Then
        If frmOptions.chkDisableDecComp = vbChecked Then
            objMaxImImager.GuiderDeclination = 0
        Else
            objMaxImImager.GuiderDeclination = Mount.CurrentDec
        End If
                
        Do
            On Error Resume Next
            Guiding = objMaxImImager.GuiderTrack(myAGExposureTime)
            If Err.Number = 65535 Or Guiding = False Then
                On Error GoTo 0
                
                Guiding = False
            
                Call AddToStatus("Error starting autoguider - MaxIm.GuiderTrack.")
                Call AddToStatus("Guiding = " & Guiding)
                Call AddToStatus("Error Source = " & Err.Source)
                Call AddToStatus("Error Number = " & Err.Number)
                Call AddToStatus("Error Description = " & Err.Description)
                Call AddToStatus("Attempting to stop autoguider in case it is still running.")
                
                Call objMaxImImager.GuiderStop
            Else
                On Error GoTo 0
                
                Call Wait(1)
                    
                If (CDbl(CInt(myGuiderX)) <> myGuiderX) Or (CDbl(CInt(myGuiderY)) <> myGuiderY) Then
                    On Error Resume Next
                    
                    If Not MultiStarGuiding Then
                        Guiding = objMaxImImager.GuiderMoveStar(myGuiderX, myGuiderY)
                    End If
            
                    If Err.Number = 65535 Or Guiding = False Then
                        On Error GoTo 0
                        
                        Guiding = False
                    
                        Call AddToStatus("Error starting autoguider - MaxIm.GuiderMoveStar.")
                        Call AddToStatus("Attempting to stop autoguider in case it is still running.")
                        Call objMaxImImager.GuiderStop
                    Else
                        On Error GoTo 0
                    End If
                End If
            End If
            
            Counter = Counter + 1
        Loop While Not Guiding And Counter < 5
        
        On Error GoTo 0
    End If
End Sub

Public Sub StopAutoguider()
    If objMaxImImager.GuiderRunning Then
        Call Wait(1)
        Do While Not objMaxImImager.GuiderStop
            Call AddToStatus("GuiderStop failed, delaying.")
            Call Wait(1)
            Call AddToStatus("Attempting to stop autoguider again.")
        Loop
        Call Wait(1)
    End If
End Sub

Public Sub SubFrame(OnOff As Boolean, Optional Left As Long, Optional Right As Long, Optional Top As Long, Optional Bottom As Long)
    If OnOff = True Then
        objMaxImImager.StartX = Left
        objMaxImImager.NumX = Right - Left
        objMaxImImager.StartY = Top
        objMaxImImager.NumY = Bottom - Top
        
        'Call AddToStatus("StartX = " & objMaxImImager.StartX & ", NumX = " & objMaxImImager.NumX & ", StartY = " & objMaxImImager.StartY & ", NumY = " & objMaxImImager.NumY)
    Else
        objMaxImImager.SetFullFrame
    End If
End Sub

Public Sub TakeDummyImage(Filter As Integer)
    Dim CurrentAutosave As Boolean
    Dim Counter As Integer
    Dim Counter2 As Integer
        
    CurrentAutosave = myAutosaveState
    
    Counter2 = 1
    Do
        Call AutoSave(False)
        
        Call SubFrame(True, 0, 50, 0, 50)
        
        Me.FilterNumber = Filter
        
        Call objMaxImImager.Expose(0.11, 1, Filter)
        On Error Resume Next
        'ImageReady will return an error if aborted
        Counter = 0
        Do Until CBool(objMaxImImager.ImageReady) Or Counter > 30
            If Err.Number <> 0 And Not Aborted Then
                Call AddToStatus("MaxIm.ImageReady returned an error: " & Err.Number & ", " & Err.Description)
                Counter = 31
                Exit Do
            End If
            
            If Aborted Then Exit Do
            
            Call Wait(1)
            Counter = Counter + 1
        Loop
        On Error GoTo 0
        Call Wait(1)
        
        If Aborted Then Exit Do
            
        If Counter > 30 And Counter2 <= 5 And Not Aborted Then
            Counter2 = Counter2 + 1
            If Counter2 <= 5 Then
                Call AddToStatus("Error changing filter in MaxIm.  Trying again, attempt #" & Counter2)
            End If
            Call objMaxImImager.AbortExposure
            Call Wait(1)
        End If
    Loop Until Counter <= 30 Or Counter2 > 5
    
    If Counter > 30 And Counter2 > 5 Then
        Call objMaxImImager.AbortExposure
        Call Wait(1)
        Call AddToStatus("Filter change failed.  Check your hardware.")
        
        Call Err.Raise(65533, "CCDCommander:MaxImControl", "Filter change error.")
    End If
    
    Call AutoSave(CurrentAutosave)

    Call SubFrame(False)
End Sub

Public Sub TakeImage()
    If myImageType = 1 Or myImageType = 4 Then
        Call objMaxImImager.Expose(myExposureTime, 1)
    ElseIf myImageType = 3 Then
        Call objMaxImImager.Expose(myExposureTime, 0)
    ElseIf myImageType = 2 Then
        Call objMaxImImager.Expose(0, 0)
    End If
    StartedImage = True
    ExposureComplete = False
End Sub

Public Property Get TakeImageComplete() As Boolean
    Dim MaxImDoc As MaxIm.Document
    Dim SequenceNum As Long
    Dim Alt As Double
    Dim Az As Double
    Dim Compression As Integer
    
    If ExposureComplete Then
        On Error Resume Next
        TakeImageComplete = objMaxImImager.ImageReady
        On Error GoTo 0
        If TakeImageComplete And StartedImage = True Then
            Call Wait(1)
            StartedImage = False
            If myAutosaveState = True Then
                On Error Resume Next
                Call AddToStatus("Saving image...")
                SequenceNum = CLng(GetMySetting("ProgramSettings", "MaxImSequenceNum", "0"))
                Set MaxImDoc = objMaxImImager.Document
                
                If (Mount.CurrentTargetJ2000) Then
                    Call MaxImDoc.SetFITSKey("OBJCTRA", Misc.FormatRAForFITSHeader(Mount.CurrentTargetJ2000RA))
                    Call MaxImDoc.SetFITSKey("OBJCTDEC", Misc.FormatDecForFITSHeader(Mount.CurrentTargetJ2000Dec))
                Else
                    Call MaxImDoc.SetFITSKey("OBJCTRA", Misc.FormatRAForFITSHeader(Mount.CurrentRA))
                    Call MaxImDoc.SetFITSKey("OBJCTDEC", Misc.FormatDecForFITSHeader(Mount.CurrentDec))
                End If
                Call MaxImDoc.SetFITSKey("OBJECT", Mount.CurrentTargetName)
                Call Mount.GetTelescopeAltAz(Alt, Az)
                Call MaxImDoc.SetFITSKey("CENTAZ", Format(Az, "0.0000"))
                Call MaxImDoc.SetFITSKey("CENTALT", Format(Alt, "0.0000"))
                
                If myImageType = 4 Then
                    Call MaxImDoc.SetFITSKey("IMAGETYP", "Flat Field")
                End If
                
                Call MaxImDoc.SetFITSKey("HISTORY", "Automated by CCD Commander")
                
                If frmOptions.chkMaxImCompression.Value = vbChecked Then
                    Compression = 1
                Else
                    Compression = 0
                End If
                
                If myReductionType = FullReduction Then
                    Call MaxImDoc.SaveFile(myAutoSavePath & myAutosaveFN & "_" & Format(SequenceNum, "000000000") & ".RAW.FIT", mxFITS, False, mx16BitPF, Compression)
                    If Err.Number <> 0 Then
                        Call AddToStatus("Error saving image: Error #" & Err.Number & ", " & Err.Description)
                        Aborted = True
                    End If
                    
                    If Not Aborted Then
                        On Error Resume Next
                        Call MaxImDoc.Calibrate
                        If Err.Number <> 0 Then
                            Call AddToStatus("Error while calibrating image.")
                            Call AddToStatus("Error #" & Err.Number & ", " & Err.Description)
                        Else
                            On Error Resume Next
                            Call MkDir(myAutoSavePath & "Reduced")
                            On Error GoTo 0
                        
                            Call MaxImDoc.SaveFile(myAutoSavePath & "Reduced\" & myAutosaveFN & "_" & Format(SequenceNum, "000000000") & ".REDUCED.FIT", mxFITS, False, mx16BitPF, Compression)
                            If Err.Number <> 0 Then
                                Call AddToStatus("Error saving image: Error #" & Err.Number & ", " & Err.Description)
                                Aborted = True
                            End If
                        End If
                    End If
                Else
                    Call MaxImDoc.SaveFile(myAutoSavePath & myAutosaveFN & "_" & Format(SequenceNum, "000000000") & ".FIT", mxFITS, False, mx16BitPF, Compression)
                    If Err.Number <> 0 Then
                        Call AddToStatus("Error saving image: Error #" & Err.Number & ", " & Err.Description)
                        Aborted = True
                    End If
                End If
                
                Set MaxImDoc = Nothing
                On Error GoTo 0
                SequenceNum = SequenceNum + 1
                Call SaveMySetting("ProgramSettings", "MaxImSequenceNum", Format(SequenceNum, "000000000"))
            Else
                If myReductionType <> NoReduction Then
                    Set MaxImDoc = objMaxImImager.Document
                    On Error Resume Next
                    Call MaxImDoc.Calibrate
                    If Err.Number <> 0 Then
                        Call AddToStatus("Error while calibrating image.")
                        Call AddToStatus("Error #" & Err.Number & ", " & Err.Description)
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Else
        TakeImageComplete = False
    End If
End Property

Public Property Let Temperature(Value As Double)
    objMaxImImager.TemperatureSetPoint = Value
End Property

Public Property Get Temperature() As Double
    On Error Resume Next
    Temperature = objMaxImImager.Temperature
    On Error GoTo 0
End Property

Public Property Get TemperatureSetPoint() As Double
    TemperatureSetPoint = objMaxImImager.TemperatureSetPoint
End Property

Public Sub UseGuideStarAt(X As Double, Y As Double)
    Dim MultiStarGuiding As Boolean
    
    MultiStarGuiding = False
    On Error Resume Next
    MultiStarGuiding = objMaxImImager.MultiStarGuiding
    On Error GoTo 0
    
    If AutoguiderEnabled Then
        myGuiderX = X
        myGuiderY = Y
        If Not MultiStarGuiding Then
            Call objMaxImImager.GuiderSetStarPosition(CInt(X), CInt(Y))
        End If
    End If
End Sub

Public Property Get WidthInPixels() As Long
    WidthInPixels = objMaxImImager.CameraXSize / objMaxImImager.BinX
End Property

Public Property Get MultiStarGuidingEnabled() As Boolean
    Dim MultiStarGuiding As Boolean
    
    MultiStarGuiding = False
    On Error Resume Next
    MultiStarGuiding = objMaxImImager.MultiStarGuiding
    On Error GoTo 0
    
    MultiStarGuidingEnabled = MultiStarGuiding
End Property

Public Property Let MultiStarGuidingEnabled(Value As Boolean)
    On Error Resume Next
    objMaxImImager.MultiStarGuiding = Value
    
    If Err.Number = 0 Then
        On Error GoTo 0
        If Value Then
            'Multi-star guiding has been enabled - take an exposure with the autoguider so the multi-star guiding will work
            Call AGTakeImage
        End If
    Else
        On Error GoTo 0
        'There was a problem, probably older version - ignore
    End If
End Property

Private Sub objMaxImImager_Notify(ByVal EventCode As Integer)
    Dim ThisGuideError As AutoguiderError
    
    Select Case EventCode
        Case MaxIm.ceExposureReadoutCompleted
            ExposureComplete = True
        Case MaxIm.ceGuiderExposureCompleted
            Camera.AutoguiderGuideErrorEvent = True
            
            If Autoguiding And frmMain.mnuViewItem(2).Checked Then
                Set ThisGuideError = New AutoguiderError
                ThisGuideError.XError = objMaxImImager.GuiderXError
                ThisGuideError.YError = objMaxImImager.GuiderYError
                ThisGuideError.TimeStamp = Now
                colAutoguiderErrors.Add ThisGuideError
            End If
        Case MaxIm.ceGuiderExposureStarFaded
            Camera.AutoguiderGuideFailedEvent = True
            
            If Autoguiding And frmMain.mnuViewItem(2).Checked Then
                Set ThisGuideError = New AutoguiderError
                ThisGuideError.XError = 15
                ThisGuideError.YError = 15
                ThisGuideError.TimeStamp = Now
                colAutoguiderErrors.Add ThisGuideError
            End If
        Case MaxIm.ceExposureAborted
            If Not IAmAborting And frmOptions.chkIgnoreExposureAborted.Value = vbUnchecked Then
                Call AddToStatus("Error - MaxIm reported the exposure aborted!")
                Aborted = True
            Else
                IAmAborting = False
            End If
    End Select
End Sub


