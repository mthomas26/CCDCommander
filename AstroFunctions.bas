Attribute VB_Name = "AstroFunctions"
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

Public SunSetTime As Date
Public TwilightStartTime As Date
Public MoonSetTime As Date
Public MoonRiseTime As Date

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Public Enum cstDSTType
    cstDSTUnknown = 0
    cstDSTStandard = 1
    cstDSTDaylight = 2
End Enum

Public Enum TwilightTypes
    Astronomical = 0
    Nautical = 1
    Civil = 2
End Enum

Public Sub ComputeSunSetTime(Altitude As Double)
    Dim Latitude As Double
    Dim Longitude As Double
    Dim Elevation As Double
    Dim JD_t As Double
    Dim JD_0 As Double
    Dim N As Double
    Dim Ti As Double
    Dim Z0 As Double
    Dim MeanAnomaly As Double
    Dim SunLongitude As Double
    Dim SunRA As Double
    Dim SunDec As Double
    Dim SunLocalHourAngle As Double
    Dim LocalMeanTime As Double
    Dim UT As Double
    Dim TempX As Double
    Dim TempY As Double
    Dim TempAngle As Double
    Dim X0 As Double
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Long
    Dim LocalTime As Double
    Dim LocalTimeH As Integer
    Dim LocalTimeM As Integer
    Dim LocalTimeS As Integer
                
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                
    Latitude = objTele.Latitude
    Longitude = objTele.Longitude
    Elevation = objTele.Elevation
    
    Longitude = Longitude / 15
    JD_t = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")))
    JD_0 = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), 1, 0)
    
    N = JD_t - JD_0
    
    Ti = N + ((18 + Longitude) / 24)
    
    Z0 = 90 - Altitude
    
    MeanAnomaly = (0.9856 * Ti) - 3.289
    SunLongitude = MeanAnomaly + (1.916 * Sin(MeanAnomaly * PI / 180#)) + (0.02 * Sin(2 * MeanAnomaly * PI / 180#)) + 282.634
    SunLongitude = Misc.DoubleModulus(SunLongitude, 360)
    
    TempX = Cos(SunLongitude * PI / 180#)
    TempY = 0.91746 * Sin(SunLongitude * PI / 180#)
    
    If (TempX > 0) And (TempY > 0) Then
        TempAngle = Atn(TempY / TempX)
    ElseIf (TempX < 0) And (TempY > 0) Then
        TempAngle = PI + Atn(TempY / TempX)
    ElseIf (TempX < 0) And (TempY < 0) Then
        TempAngle = PI + Atn(TempY / TempX)
    ElseIf (TempX > 0) And (TempY < 0) Then
        TempAngle = (2 * PI) + Atn(TempY / TempX)
    End If
    
    SunRA = (180 / (PI * 15)) * TempAngle
    
    SunDec = (180 / PI) * Misc.ASin(0.39782 * Sin(SunLongitude * PI / 180))
    
    X0 = (Cos(Z0 * PI / 180) - (Sin(SunDec * PI / 180) * Sin(Latitude * PI / 180))) / (Cos(SunDec * PI / 180) * Cos(Latitude * PI / 180))
    
    On Error Resume Next
    SunLocalHourAngle = (180 / PI) * ACos(X0) / 15
    If Err.Number <> 0 Then
        ' Error probably means I'm trying to get above the maximum altitude (or really close to it)
        ' Just use the maximum hour angle
        SunLocalHourAngle = 24
    End If
    On Error GoTo 0
        
    LocalMeanTime = SunLocalHourAngle + SunRA - (0.06571 * Ti) - 6.622
    LocalMeanTime = Misc.DoubleModulus(LocalMeanTime, 24)
    
    UT = Misc.DoubleModulus(LocalMeanTime + Longitude, 24)
    
    If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
    Else
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
    End If
    
    LocalTime = Misc.DoubleModulus(UT - (LocalBias / 60) + (0.5 / 3600), 24)
    LocalTimeH = Int(LocalTime)
    LocalTimeM = Int((LocalTime - Int(LocalTime)) * 60)
    LocalTimeS = Int((((LocalTime - Int(LocalTime)) * 60) - LocalTimeM) * 60)
    
    SunSetTime = LocalTimeH & ":" & LocalTimeM & ":" & LocalTimeS
End Sub

Public Sub ComputeTwilightStartTime(Altitude As Double)
    Dim Latitude As Double
    Dim Longitude As Double
    Dim Elevation As Double
    Dim JD_t As Double
    Dim JD_0 As Double
    Dim N As Double
    Dim Ti As Double
    Dim Z3 As Double
    Dim MeanAnomaly As Double
    Dim SunLongitude As Double
    Dim SunRA As Double
    Dim SunDec As Double
    Dim SunLocalHourAngle As Double
    Dim LocalMeanTime As Double
    Dim UT As Double
    Dim TempX As Double
    Dim TempY As Double
    Dim TempAngle As Double
    Dim X3 As Double
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Long
    Dim LocalTime As Double
    Dim LocalTimeH As Integer
    Dim LocalTimeM As Integer
    Dim LocalTimeS As Integer
    
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                                
    Latitude = objTele.Latitude
    Longitude = objTele.Longitude
    Elevation = objTele.Elevation
    
    Longitude = Longitude / 15
    If CInt(Format(Time, "hh")) > 12 Then
        JD_t = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")) + 1)
    Else
        JD_t = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")))
    End If
    JD_0 = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), 1, 0)
    
    N = JD_t - JD_0
    
    Ti = N + ((6 + Longitude) / 24)
    
    Z3 = 90 - Altitude
    
    MeanAnomaly = (0.9856 * Ti) - 3.289
    SunLongitude = MeanAnomaly + (1.916 * Sin(MeanAnomaly * PI / 180)) + (0.02 * Sin(2 * MeanAnomaly * PI / 180)) + 282.634
    SunLongitude = Misc.DoubleModulus(SunLongitude, 360)
    
    TempX = Cos(SunLongitude * PI / 180)
    TempY = 0.91746 * Sin(SunLongitude * PI / 180)
    
    If (TempX > 0) And (TempY > 0) Then
        TempAngle = Atn(TempY / TempX)
    ElseIf (TempX < 0) And (TempY > 0) Then
        TempAngle = PI + Atn(TempY / TempX)
    ElseIf (TempX < 0) And (TempY < 0) Then
        TempAngle = PI + Atn(TempY / TempX)
    ElseIf (TempX > 0) And (TempY < 0) Then
        TempAngle = (2 * PI) + Atn(TempY / TempX)
    End If
    
    SunRA = (180 / (PI * 15)) * TempAngle
    
    SunDec = (180 / PI) * Misc.ASin(0.39782 * Sin(SunLongitude * PI / 180))
    
    X3 = (Cos(Z3 * PI / 180) - (Sin(SunDec * PI / 180) * Sin(Latitude * PI / 180))) / (Cos(SunDec * PI / 180) * Cos(Latitude * PI / 180))
    
    On Error Resume Next
    SunLocalHourAngle = (360 - ((180 / PI) * ACos(X3))) / 15
    If Err.Number <> 0 Then
        ' Error probably means I'm trying to get above the maximum altitude (or really close to it)
        ' Just use the maximum hour angle
        SunLocalHourAngle = 24
    End If
    On Error GoTo 0
    
    LocalMeanTime = SunLocalHourAngle + SunRA - (0.06571 * Ti) - 6.622
    LocalMeanTime = Misc.DoubleModulus(LocalMeanTime, 24)
    
    UT = Misc.DoubleModulus(LocalMeanTime + Longitude, 24)
    
    If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
    Else
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
    End If
    
    LocalTime = Misc.DoubleModulus(UT - (LocalBias / 60) + (0.5 / 3600), 24)
    LocalTimeH = Int(LocalTime)
    LocalTimeM = Int((LocalTime - Int(LocalTime)) * 60)
    LocalTimeS = Int((((LocalTime - Int(LocalTime)) * 60) - LocalTimeM) * 60)
    
    TwilightStartTime = LocalTimeH & ":" & LocalTimeM & ":" & LocalTimeS
End Sub


Public Sub MoonRise(Alt As Double)
    Dim RA As Double
    Dim Dec As Double
    Dim PI As Double
    Dim JD0 As Double
    Dim LST0 As Double
    Dim LST1 As Double
    Dim JDUT As Double
    Dim T As Double
    Dim GST0 As Double
    Dim LSTdegrees As Double
    Dim BRKT(0 To 36, 2) As Double
    Dim mrise As Long
    Dim mset As Long
    Dim flag As Long
    Dim pVal As Double
    Dim R1 As Double
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Double
    Dim i As Integer, k As Integer, j As Integer
    Dim u As Double, v As Double, e As Double, W As Double, a As Double, b As Double, c As Double
    Dim MaxAlt As Double
    Dim MaxAltHour As Integer
    Dim LoopHourStart As Integer
    Dim LoopHourStop As Integer
    
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                
    RA = 0: Dec = 0: PI = 0: LSTdegrees = 0: mrise = -1: mset = -1: flag = 0: pVal = 0: R1 = Misc.PI / 180
    
    If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
    Else
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
    End If
    
    LocalBias = LocalBias / 60
    
    JDUT = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")))
    
    JD0 = JDUT + (LocalBias / 24)
    
    T = (JDUT - 2415020#) / 36525
    GST0 = 24 + 6.6460656 + 2400.051262 * T + 0.00002581 * T * T
    GST0 = GST0 - (24 * (Fix(GST0 / 24)))
    LST0 = GST0 + LocalBias * 1.002737908 + (-objTele.Longitude / 15#) + 24#    '// Long = -West
    LST0 = LST0 - (24 * (Fix(LST0 / 24)))

    '// Calculate Moon's Position and Corrected Altitude for 12h - 36h LMT at 1h Intervals
    '// Store values in BRKT array. Correct altitude for parallax, refraction, and moon's semi-diameter.
    '// Run bisection algorithm on segments [i-1] to [i] that show a sign change.
    MaxAlt = -99
    If Hour(Now) < 12 Then
        LoopHourStart = 0
        LoopHourStop = 24
    Else
        LoopHourStart = 12
        LoopHourStop = 36
    End If
    For i = LoopHourStart To LoopHourStop
        BRKT(i, 0) = JD0 + i / 24#
        Call MoonPosition(BRKT(i, 0), RA, Dec, PI)
        LST1 = LST0 + 1.002737908 * i
        LST1 = LST1 - (24 * (Fix(LST1 / 24)))
        LSTdegrees = LST1 * 15

        BRKT(i, 1) = Sin(objTele.Latitude * R1) * Sin(Dec * R1) + Cos(objTele.Latitude * R1) * Cos(Dec * R1) * Cos((LSTdegrees - RA) * R1)
        'BRKT(i, 1) = (180 / Misc.PI) * Misc.ASin(BRKT(i, 1)) - 0.7275 * PI + 0.5667 -- Changed to the below, seems to work better
        BRKT(i, 1) = (180 / Misc.PI) * Misc.ASin(BRKT(i, 1)) - 0.7275 * PI

        If (BRKT(i, 1) > MaxAlt) Then
            MaxAlt = BRKT(i, 1)
            MaxAltHour = i
        End If

        '// Run bisection algorithm on interval to find zero crossing
        '// only if interval shows sign change between endpoints
        If (i > LoopHourStart) Then
            If (BRKT(i, 1) > Alt) And (BRKT(i - 1, 1) < Alt) Then
                c = 0
                u = BRKT(i - 1, 1)
                v = BRKT(i, 1)
                b = BRKT(i, 0)
                a = BRKT(i - 1, 0)
                e = b - a
                        
                For k = 0 To 20
                    e = e / 2#
                    c = a + e
                    Call MoonPosition(c, RA, Dec, PI)
                    LST1 = LST0 + 1.002737908 * (c - JD0) * 24# + 24
                    LST1 = LST1 - (24 * (Fix(LST1 / 24)))
                    LSTdegrees = LST1 * 15
                    W = Sin(objTele.Latitude * R1) * Sin(Dec * R1) + Cos(objTele.Latitude * R1) * Cos(Dec * R1) * Cos((LSTdegrees - RA) * R1)
                    'W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI + 0.5667 -- Changed to the below, seems to work better
                    W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI
    
                    If (Abs(W) < 0.001) Then Exit For
                    If (W > Alt) Then
                        b = c
                        v = W
                    Else
                        a = c
                        u = W
                    End If
                Next k
                
                BRKT(i, 2) = 24 * (c - JD0)
            End If
        End If
    Next i

    '// Search BRKT array for the transition event - if it exists
    For j = LoopHourStart To LoopHourStop
        If j > LoopHourStart Then
            If (BRKT(j, 1) > Alt) And (BRKT(j - 1, 1) < Alt) Then
                MoonRiseTime = DateAdd("s", BRKT(j, 2) * 3600, Date)
                Exit Sub
            End If
        End If
    Next j
    
    ' Never reached the desired altitude - check if desired alt is greater than max alt
    If Alt > MaxAlt Then
        'Desired alt greater than max alt, interpolate to actual max
        If BRKT(MaxAltHour - 1, 1) > BRKT(MaxAltHour + 1, 1) Then
            'Max is between MaxAltHour-1 and MaxAltHour
            b = BRKT(MaxAltHour - 1, 0)
            a = BRKT(MaxAltHour, 0)
        Else
            'Max is between MaxAltHour and MaxAltHour-1
            b = BRKT(MaxAltHour + 1, 0)
            a = BRKT(MaxAltHour, 0)
        End If
        
        u = BRKT(MaxAltHour, 1)
        
        e = (b - a) / 2
        i = 0
        j = 1
                
        c = a
        
        For k = 0 To 20
            c = c + e
            
            Call MoonPosition(c, RA, Dec, PI)
            LST1 = LST0 + 1.002737908 * (c - JD0) * 24# + 24
            LST1 = LST1 - (24 * (Fix(LST1 / 24)))
            LSTdegrees = LST1 * 15
            W = Sin(objTele.Latitude * R1) * Sin(Dec * R1) + Cos(objTele.Latitude * R1) * Cos(Dec * R1) * Cos((LSTdegrees - RA) * R1)
            'W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI + 0.5667 -- Changed to the below, seems to work better
            W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI

            If (Abs(W) < 0.001) Then Exit For
            If (W >= MaxAlt) Then
                MaxAlt = W
                u = c
                j = 1
            ElseIf j = 1 Then
                j = 0
                e = -e / 2#
            End If
        Next k
        
        MoonRiseTime = DateAdd("s", 24 * (u - JD0) * 3600, Date)
        Exit Sub
    End If
    
    MoonRiseTime = Now
End Sub

Public Sub Moonset(Alt As Double)
    Dim RA As Double
    Dim Dec As Double
    Dim PI As Double
    Dim JD0 As Double
    Dim LST0 As Double
    Dim LST1 As Double
    Dim JDUT As Double
    Dim T As Double
    Dim GST0 As Double
    Dim LSTdegrees As Double
    Dim BRKT(0 To 36, 2) As Double
    Dim mrise As Long
    Dim mset As Long
    Dim flag As Long
    Dim pVal As Double
    Dim R1 As Double
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Double
    Dim i As Integer, k As Integer, j As Integer
    Dim u As Double, v As Double, e As Double, W As Double, a As Double, b As Double, c As Double
    Dim MinAlt As Double
    Dim MinAltHour As Integer
    Dim LoopHourStart As Integer
    Dim LoopHourStop As Integer
    
    If (objTele Is Nothing) Then
        Call Mount.MountSetup
    End If
                
    RA = 0: Dec = 0: PI = 0: LSTdegrees = 0: mrise = -1: mset = -1: flag = 0: pVal = 0: R1 = Misc.PI / 180
    
    If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
    Else
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
    End If
    
    LocalBias = LocalBias / 60
    
    JDUT = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")))
    
    JD0 = JDUT + (LocalBias / 24)
    
    T = (JDUT - 2415020#) / 36525
    GST0 = 24 + 6.6460656 + 2400.051262 * T + 0.00002581 * T * T
    GST0 = GST0 - (24 * (Fix(GST0 / 24)))
    LST0 = GST0 + LocalBias * 1.002737908 + (-objTele.Longitude / 15#) + 24#    '// Long = -West
    LST0 = LST0 - (24 * (Fix(LST0 / 24)))

    '// Calculate Moon's Position and Corrected Altitude for 12h - 36h LMT at 1h Intervals
    '// Store values in BRKT array. Correct altitude for parallax, refraction, and moon's semi-diameter.
    '// Run bisection algorithm on segments [i-1] to [i] that show a sign change.
    MinAlt = 99
    
    If Hour(Now) < 12 Then
        LoopHourStart = 0
        LoopHourStop = 24
    Else
        LoopHourStart = 12
        LoopHourStop = 36
    End If
    For i = LoopHourStart To LoopHourStop
        BRKT(i, 0) = JD0 + i / 24#
        Call MoonPosition(BRKT(i, 0), RA, Dec, PI)
        LST1 = LST0 + 1.002737908 * i
        LST1 = LST1 - (24 * (Fix(LST1 / 24)))
        LSTdegrees = LST1 * 15

        BRKT(i, 1) = Sin(objTele.Latitude * R1) * Sin(Dec * R1) + Cos(objTele.Latitude * R1) * Cos(Dec * R1) * Cos((LSTdegrees - RA) * R1)
        'BRKT(i, 1) = (180 / Misc.PI) * Misc.ASin(BRKT(i, 1)) - 0.7275 * PI + 0.5667 -- Changed to the below, seems to work better
        BRKT(i, 1) = (180 / Misc.PI) * Misc.ASin(BRKT(i, 1)) - 0.7275 * PI

        If (BRKT(i, 1) < MinAlt) Then
            MinAlt = BRKT(i, 1)
            MinAltHour = i
        End If

        '// Run bisection algorithm on interval to find zero crossing
        '// only if interval shows sign change between endpoints
        If (i > LoopHourStart) Then
            If (BRKT(i, 1) < Alt) And (BRKT(i - 1, 1) > Alt) Then
                c = 0
                u = BRKT(i - 1, 1)
                v = BRKT(i, 1)
                b = BRKT(i, 0)
                a = BRKT(i - 1, 0)
                e = b - a
                        
                For k = 0 To 20
                    e = e / 2#
                    c = a + e
                    Call MoonPosition(c, RA, Dec, PI)
                    LST1 = LST0 + 1.002737908 * (c - JD0) * 24# + 24
                    LST1 = LST1 - (24 * (Fix(LST1 / 24)))
                    LSTdegrees = LST1 * 15
                    W = Sin(objTele.Latitude * R1) * Sin(Dec * R1) + Cos(objTele.Latitude * R1) * Cos(Dec * R1) * Cos((LSTdegrees - RA) * R1)
                    'W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI + 0.5667 -- Changed to the below, seems to work better
                    W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI
    
                    If (Abs(W) < 0.001) Then Exit For
                    If (W < Alt) Then
                        b = c
                        v = W
                    Else
                        a = c
                        u = W
                    End If
                Next k
                
                BRKT(i, 2) = 24 * (c - JD0)
            End If
        End If
    Next i

    '// Search BRKT array for the transition event - if it exists
    For j = LoopHourStart To LoopHourStop
        If j > LoopHourStart Then
            If (BRKT(j, 1) < Alt) And (BRKT(j - 1, 1) > Alt) Then
                MoonSetTime = DateAdd("s", BRKT(j, 2) * 3600, Date)
                Exit Sub
            End If
        End If
    Next j
    
    ' Never reached the desired altitude - check if desired alt is less than min alt
    If Alt < MinAlt Then
        'Desired alt less than min alt, interpolate to actual min
        If BRKT(MinAltHour - 1, 1) < BRKT(MinAltHour + 1, 1) Then
            'min is between MinAltHour-1 and MinAltHour
            b = BRKT(MinAltHour - 1, 0)
            a = BRKT(MinAltHour, 0)
        Else
            'min is between MinAltHour and MinAltHour-1
            b = BRKT(MinAltHour + 1, 0)
            a = BRKT(MinAltHour, 0)
        End If
        
        u = BRKT(MinAltHour, 1)
        
        e = (b - a) / 2
        i = 0
        j = 1
                
        c = a
        
        For k = 0 To 20
            c = c + e
            
            Call MoonPosition(c, RA, Dec, PI)
            LST1 = LST0 + 1.002737908 * (c - JD0) * 24# + 24
            LST1 = LST1 - (24 * (Fix(LST1 / 24)))
            LSTdegrees = LST1 * 15
            W = Sin(objTele.Latitude * R1) * Sin(Dec * R1) + Cos(objTele.Latitude * R1) * Cos(Dec * R1) * Cos((LSTdegrees - RA) * R1)
            'W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI + 0.5667 -- Changed to the below, seems to work better
            W = (180 / Misc.PI) * Misc.ASin(W) - 0.7275 * PI

            If (Abs(W) < 0.001) Then Exit For
            If (W <= MinAlt) Then
                MinAlt = W
                u = c
                j = 1
            ElseIf j = 1 Then
                j = 0
                e = -e / 2#
            End If
        Next k
        
        MoonSetTime = DateAdd("s", 24 * (u - JD0) * 3600, Date)
        Exit Sub
    End If
    
    MoonSetTime = Now
End Sub

Private Sub MoonPosition(JED As Double, ByRef RA As Double, ByRef Dec As Double, ByRef Par As Double)
    Dim T As Double, Lp As Double, m As Double, Mp As Double, D As Double, F As Double, Omega As Double, e As Double, lambda As Double, b As Double, beta As Double, epsilon As Double, PI As Double
    Dim tmp As Double, Y As Double, X As Double, R1 As Double
    
    R1 = Misc.PI / 180

    T = (JED - 2415020#) / 36525
    '// Moon's mean longitude:
    Lp = 270.434164 + 481267.8831 * T - 0.001133 * T * T + 0.0000019 * T * T * T

    '// Sun's mean anomaly:
    m = 358.475833 + 35999.0498 * T - 0.00015 * T * T - 0.0000033 * T * T * T

    '// Moon's mean anomaly:
    Mp = 296.104608 + 477198.8491 * T + 0.009192 * T * T + 0.0000144 * T * T * T

    '// Moon's mean elongation:
    D = 350.737486 + 445267.1142 * T - 0.001436 * T * T + 0.0000019 * T * T * T

    '// Moon's mean distance from ascending node:
    F = 11.250889 + 483202.0251 * T - 0.003211 * T * T - 0.0000003 * T * T * T

    '//Moon's longitude of ascending node:
    Omega = 259.183275 - 1934.142 * T + 0.002078 * T * T + 0.0000022 * T * T * T

    '//Eccentricity
    e = 1 - 0.002495 * T - 0.00000752 * T * T

    '// Obliquity of the date
    epsilon = 23.452294 - 0.0130125 * T - 0.00000164 * T * T + 0.000000503 * T * T * T

    '// Additive terms:
    tmp = Sin((51.2 + 20.2 * T) * R1)
    Lp = Lp + (0.000233 * tmp)
    m = m + (-0.001778 * tmp)
    Mp = Mp + (0.000817 * tmp)
    D = D + (0.002011 * tmp)

    tmp = 0.003964 * Sin((346.56 + 132.87 * T - 0.0091731 * T * T) * R1)
    Lp = Lp + tmp
    Mp = Mp + tmp
    D = D + tmp
    F = F + tmp

    tmp = Sin(Omega * R1)
    Lp = Lp + (0.001964 * tmp)
    Mp = Mp + (0.002541 * tmp)
    D = D + (0.001964 * tmp)
    F = F + (-0.024691 * tmp)

    F = F + (-0.004328 * Sin((Omega + 275.05 - 2.3 * T) * R1))

    '//Moon's geocentric longitude
    lambda = Lp + 6.28875 * Sin(Mp * R1) + 1.274018 * Sin((2 * D - Mp) * R1) + 0.658309 * Sin(2 * D * R1) + 0.213616 * Sin(2 * Mp * R1) _
                - 0.185596 * Sin(m * R1) * e - 0.114336 * Sin(2 * F * R1) + 0.058793 * Sin((2 * D - 2 * Mp) * R1) + 0.057212 * Sin((2 * D - m - Mp) * R1) * e _
                + 0.05332 * Sin((2 * D + Mp) * R1) + 0.045874 * Sin((2 * D - m) * R1) * e + 0.041024 * Sin((Mp - m) * R1) * e - 0.034718 * Sin(D * R1) _
                - 0.030465 * Sin((m + Mp) * R1) * e + 0.015326 * Sin((2 * D - 2 * F) * R1) - 0.012528 * Sin((2 * F + Mp) * R1) - 0.01098 * Sin((2 * F - Mp) * R1) _
                + 0.010674 * Sin((4 * D - Mp) * R1) + 0.010034 * Sin(3 * Mp * R1) + 0.008548 * Sin((4 * D - 2 * Mp) * R1) - 0.00791 * Sin((m - Mp + 2 * D) * R1) * e _
                - 0.006783 * Sin((2 * D + m) * R1) * e + 0.005162 * Sin((Mp - D) * R1) + 0.005 * Sin((m + D) * R1) * e + 0.004049 * Sin((Mp - m + 2 * D) * R1) * e _
                + 0.003996 * Sin((2 * Mp + 2 * D) * R1) + 0.003862 * Sin(4 * D * R1) + 0.003665 * Sin((2 * D - 3 * Mp) * R1) + 0.002695 * Sin((2 * Mp - m) * R1) * e _
                + 0.002602 * Sin((Mp - 2 * F - 2 * D) * R1) + 0.002396 * Sin((2 * D - m - 2 * Mp) * R1) * e - 0.002349 * Sin((Mp + D) * R1) + 0.002249 * Sin((2 * D - 2 * m) * R1) * e * e _
                - 0.002125 * Sin((2 * Mp + m) * R1) * e - 0.002079 * Sin(2 * m * R1) * e * e + 0.002059 * Sin((2 * D - Mp - 2 * m) * R1) * e * e - 0.001773 * Sin((Mp + 2 * D - 2 * F) * R1) _
                - 0.001595 * Sin((2 * F + 2 * D) * R1) + 0.00122 * Sin((4 * D - m - Mp) * R1) * e - 0.00111 * Sin((2 * Mp + 2 * F) * R1) + 0.000892 * Sin((Mp - 3 * D) * R1) _
                - 0.000811 * Sin((m + Mp + 2 * D) * R1) * e + 0.000761 * Sin((4 * D - m - 2 * Mp) * R1) * e + 0.000717 * Sin((Mp - 2 * m) * R1) * e * e + 0.000704 * Sin((Mp - 2 * m - 2 * D) * R1) * e * e _
                + 0.000693 * Sin((m - 2 * Mp + 2 * D) * R1) * e + 0.000598 * Sin((2 * D - m - 2 * F) * R1) * e + 0.00055 * Sin((Mp + 4 * D) * R1) + 0.000538 * Sin(4 * Mp * R1) _
                + 0.000521 * Sin((4 * D - m) * R1) * e + 0.000486 * Sin((2 * Mp - D) * R1)

    lambda = lambda - (360 * (Fix(lambda / 360)))

    b = 5.128189 * Sin(F * R1) + 0.280606 * Sin((Mp + F) * R1) + 0.277693 * Sin((Mp - F) * R1) + 0.173238 * Sin((2 * D - F) * R1) + 0.055413 * Sin((2 * D + F - Mp) * R1) _
        + 0.046272 * Sin((2 * D - F - Mp) * R1) + 0.032573 * Sin((2 * D + F) * R1) + 0.017198 * Sin((2 * Mp + F) * R1) + 0.009267 * Sin((2 * D + Mp - F) * R1) + 0.008823 * Sin((2 * Mp - F) * R1) _
        + 0.008247 * Sin((2 * D - m - F) * R1) * e + 0.004323 * Sin((2 * D - F - 2 * Mp) * R1) + 0.0042 * Sin((2 * D + F + Mp) * R1) + 0.003372 * Sin((F - m - 2 * D) * R1) * e + 0.002472 * Sin((2 * D + F - m - Mp) * R1) * e _
        + 0.002222 * Sin((2 * D + F - m) * R1) * e + 0.002072 * Sin((2 * D - F - m - Mp) * R1) * e + 0.001877 * Sin((F - m + Mp) * R1) * e + 0.001828 * Sin((4 * D - F - Mp) * R1) - 0.001803 * Sin((F + m) * R1) * e _
        - 0.00175 * Sin(3 * F * R1) + 0.00157 * Sin((Mp - m - F) * R1) * e - 0.001487 * Sin((F + D) * R1) - 0.001481 * Sin((F + m + Mp) * R1) * e + 0.001417 * Sin((F - m - Mp) * R1) * e _
        + 0.00135 * Sin((F - m) * R1) * e + 0.00133 * Sin((F - D) * R1) + 0.001106 * Sin((F + 3 * Mp) * R1) + 0.00102 * Sin((4 * D - F) * R1) + 0.000833 * Sin((F + 4 * D - Mp) * R1) _
        + 0.000781 * Sin((Mp - 3 * F) * R1) + 0.00067 * Sin((F + 4 * D - 2 * Mp) * R1) + 0.000606 * Sin((2 * D - 3 * F) * R1) + 0.000597 * Sin((2 * D + 2 * Mp - F) * R1) + 0.000492 * Sin((2 * D + Mp - m - F) * R1) * e _
        + 0.00045 * Sin((2 * Mp - F - 2 * D) * R1) + 0.000439 * Sin((3 * Mp - F) * R1) + 0.000423 * Sin((F + 2 * D + 2 * Mp) * R1) + 0.000422 * Sin((2 * D - F - 3 * Mp) * R1) - 0.000367 * Sin((m + F + 2 * D - Mp) * R1) * e _
        - 0.000353 * Sin((m + F + 2 * D) * R1) * e + 0.000331 * Sin((F + 4 * D) * R1) + 0.000317 * Sin((2 * D + F - m + Mp) * R1) * e + 0.000306 * Sin((2 * D - 2 * m - F) * R1) * e * e - 0.000283 * Sin((Mp + 3 * F) * R1)

    beta = b * (1 - 0.0004664 * Cos(Omega * R1) - 0.0000754 * Cos((Omega + 275.05 - 2.3 * T) * R1))

    beta = beta - (360 * (Fix(beta / 360)))

    PI = 0.950724 + 0.051818 * Cos(Mp * R1) + 0.009531 * Cos((2 * D - Mp) * R1) + 0.007843 * Cos(2 * D * R1) + 0.002824 * Cos(2 * Mp * R1) _
        + 0.000857 * Cos((2 * D + Mp) * R1) + 0.000533 * Cos((2 * D - m) * R1) * e + 0.000401 * Cos((2 * D - m - Mp) * R1) * e + 0.00032 * Cos((Mp - m) * R1) * e _
        - 0.000271 * Cos(D * R1) - 0.000264 * Cos((m + Mp) * R1) * e - 0.000198 * Cos((2 * F - Mp) * R1) + 0.000173 * Cos(3 * Mp * R1) + 0.000167 * Cos((4 * D - Mp) * R1) _
        - 0.000111 * Cos(m * R1) * e + 0.000103 * Cos((4 * D - 2 * Mp) * R1) - 0.000084 * Cos((2 * Mp - 2 * D) * R1) - 0.000083 * Cos((2 * D + m) * R1) * e _
        + 0.000079 * Cos((2 * D + 2 * Mp) * R1) + 0.000072 * Cos(4 * D * R1) + 0.000064 * Cos((2 * D - m + Mp) * R1) * e - 0.000063 * Cos((2 * D + m - Mp) * R1) * e _
        + 0.000041 * Cos((m + D) * R1) * e + 0.000035 * Cos((2 * Mp - m) * R1) * e - 0.000033 * Cos((3 * Mp - 2 * D) * R1) - 0.00003 * Cos((Mp + D) * R1) _
        - 0.000029 * Cos((2 * F - 2 * D) * R1) - 0.000029 * Cos((2 * Mp + m) * R1) * e + 0.000026 * Cos((2 * D - 2 * m) * R1) * e * e - 0.000023 * Cos((2 * F - 2 * D + Mp) * R1) _
        + 0.000019 * Cos((4 * D - m - Mp) * R1) * e


    Y = Sin(lambda * R1) * Cos(epsilon * R1) - Tan(beta * R1) * Sin(epsilon * R1)
    X = Cos(lambda * R1)
    tmp = (180 / Misc.PI) * Misc.Atn360(X, Y) + 360
    tmp = tmp - (360 * (Fix(tmp / 360)))

    If (tmp = 360) Then tmp = 0
    RA = tmp   '//RA

    Dec = (180 / Misc.PI) * ASin(Sin(beta * R1) * Cos(epsilon * R1) + Cos(beta * R1) * Sin(epsilon * R1) * Sin(lambda * R1))

    Par = PI
End Sub


