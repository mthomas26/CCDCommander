Attribute VB_Name = "Misc"
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
Private Enum cstDSTType
    cstDSTUnknown = 0
    cstDSTStandard = 1
    cstDSTDaylight = 2
End Enum

Public Function PI() As Double
    PI = 4# * Atn(1#)
End Function

Public Function FixFileName(DesiredName As String, Optional Path As Boolean = False) As String
    Dim Counter As Long
    Dim FixedFileName As String
    Dim TestChar As String * 1
    Dim Parameter As String
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Long
    Dim myUTTime As Date
    
    For Counter = 1 To Len(DesiredName)
        TestChar = Mid(DesiredName, Counter, 1)
        If TestChar = "<" Then
            'Maybe start of a naming parameter!
            If InStr(Counter, DesiredName, ">") > 0 Then
                'Got a parameter!
                If (Right(FixedFileName, 1) <> "_") And Len(FixedFileName) > 0 And _
                    ((Right(FixedFileName, 1) <> "\") And (Right(FixedFileName, 1) <> "/") And (Right(FixedFileName, 1) <> ":") And Path) Then
                    
                    FixedFileName = FixedFileName & "_"
                End If
                
                Parameter = Mid(DesiredName, Counter + 1, InStr(Counter, DesiredName, ">") - Counter - 1)
                Select Case UCase(Parameter)
                    Case UCase("Date")
                        FixedFileName = FixedFileName & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
                    Case UCase("Time")
                        FixedFileName = FixedFileName & Format(Now, "hhmm")
                    Case UCase("UT")
                        If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
                            LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
                        Else
                            LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
                        End If
                        
                        myUTTime = DateAdd("n", LocalBias, Now)
                        FixedFileName = FixedFileName & Format(myUTTime, "hhmm")
                    Case UCase("DateUT")
                        If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
                            LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
                        Else
                            LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
                        End If
                        
                        myUTTime = DateAdd("n", LocalBias, Now)
                        FixedFileName = FixedFileName & Format(Year(myUTTime), "0000") & Format(Month(myUTTime), "00") & Format(Day(myUTTime), "00")
                    Case UCase("ExposureTime")
                        FixedFileName = FixedFileName & Camera.objCameraControl.ExposureTime & "s"
                    Case UCase("Bin")
                        FixedFileName = FixedFileName & Camera.objCameraControl.BinX & "x" & Camera.objCameraControl.BinY
                    Case UCase("Filter")
                        If (frmOptions.lstFilters.ListCount > 0) Then
                            FixedFileName = FixedFileName & frmOptions.lstFilters.List(Camera.objCameraControl.FilterNumber)
                        Else
                            'ignore this one
                        End If
                    Case UCase("ImageType")
                        FixedFileName = FixedFileName & ImageTypes(Camera.objCameraControl.ImageType - 1)
                    Case UCase("ObjectName")
                        FixedFileName = FixedFileName & FixFileName(Mount.CurrentTargetName) 'calling fixfilename to remove any invalid characters from CurrentTargetName
                    Case UCase("ObjectCoords")
                        FixedFileName = FixedFileName & Misc.FormatRAForFITSHeader(Mount.CurrentRA, True) & "_" & Misc.FormatDecForFITSHeader(Mount.CurrentDec, True)
                    Case UCase("PA")
                        FixedFileName = FixedFileName & Format(Rotator.CurrentAngle, "0.0") & "degN"
                    Case UCase("RotatorAngle")
                        FixedFileName = FixedFileName & Format(Rotator.GetCurrentRotatorAngle, "0.0") & "deg"
                    Case UCase("Temperature")
                        FixedFileName = FixedFileName & Camera.objCameraControl.TemperatureSetPoint & "degC"
                End Select
                
                FixedFileName = FixedFileName & "_"
                
                Counter = InStr(Counter, DesiredName, ">")
            Else
                'No parameter end, ignore
            End If
        ElseIf (TestChar = "\" Or TestChar = "/" Or TestChar = ":") Then
            If Not Path Then
                'ignore character
            Else
                If (Right(FixedFileName, 1) = "_") Then
                    FixedFileName = Left(FixedFileName, Len(FixedFileName) - 1)
                End If
                
                FixedFileName = FixedFileName & TestChar
            End If
            
        ElseIf TestChar = "*" Or TestChar = "?" Or TestChar = Chr(34) Or TestChar = "<" Or TestChar = ">" Or TestChar = "|" Or TestChar = "," Then
            'ignore character
            
        Else
            FixedFileName = FixedFileName & TestChar
        End If
    Next Counter
        
    If Right(FixedFileName, 1) = "_" Then
        FixedFileName = Left(FixedFileName, Len(FixedFileName) - 1)
    End If
    
    FixFileName = FixedFileName
        
End Function

Public Function DoubleModulus(Value As Double, Modulus As Double) As Double
    Dim myVal As Double
    
    myVal = (Value / Modulus)
    myVal = (myVal - Fix(myVal)) * Modulus
    
    If myVal >= 0 Then
        DoubleModulus = myVal
    Else
        DoubleModulus = Modulus + myVal
    End If
End Function

Public Sub RotateVector(VectorX As Double, VectorY As Double, Angle As Double)
    Dim VectorMag As Double
    Dim VectorAngle As Double
        
    VectorMag = Sqr(VectorX * VectorX + VectorY * VectorY)
    If VectorX = 0 Then
        If VectorY >= 0 Then
            VectorAngle = 90
        Else
            VectorAngle = 270
        End If
    ElseIf VectorY = 0 Then
        If VectorX >= 0 Then
            VectorAngle = 0
        Else
            VectorAngle = 180
        End If
    Else
        VectorAngle = Atn(VectorY / VectorX) * 180 / PI
        If VectorX < 0 Then
            VectorAngle = Misc.DoubleModulus(VectorAngle + 180, 360)
        End If
    End If
    
    VectorAngle = DoubleModulus((VectorAngle + Angle), 360)
    
    VectorX = Cos(VectorAngle * PI / 180) * VectorMag
    VectorY = Sin(VectorAngle * PI / 180) * VectorMag
End Sub

Public Function ConvertEquatorialToString(ByVal RA As Double, ByVal Dec As Double, LowPrecision As Boolean) As String
    Dim myString As String
    Dim RoundingValue As Double
    
    If LowPrecision Then
        RoundingValue = 0.5
    Else
        RoundingValue = 0.05
    End If
    
    myString = "RA: " & Format(Fix(RA + (RoundingValue / 3600)), "00") & "h "
    
    RA = (RA - Fix(RA + (RoundingValue / 3600))) * 60
    
    myString = myString & Format(Fix(RA + (RoundingValue / 60)), "00") & "m "
    
    RA = (RA - Fix(RA + (RoundingValue / 60))) * 60
    
    If Not LowPrecision Then
        myString = myString & Format(RA + RoundingValue, "00.0") & "s "
    End If
    
    If (Dec < 0) Then
        myString = myString & "Dec: -"
        Dec = Abs(Dec)
    Else
        myString = myString & "Dec: +"
    End If
    
    RoundingValue = 0.5
    
    myString = myString & Format(Fix(Dec + (RoundingValue / 3600)), "00") & Chr(176)
    
    Dec = (Dec - Fix(Dec + (RoundingValue / 3600))) * 60
    
    myString = myString & Format(Fix(Dec + (RoundingValue / 60)), "00") & "'"
    
    Dec = (Dec - Fix(Dec + (RoundingValue / 60))) * 60
    
    If Not LowPrecision Then
        myString = myString & Format(Dec + RoundingValue, "00") & Chr(34)
    End If
    
    ConvertEquatorialToString = myString
End Function

Public Function ConvertRAToString(ByVal RA As Double, LowPrecision As Boolean) As String
    Dim myString As String
    Dim RoundingValue As Double
    
    If LowPrecision Then
        RoundingValue = 0.5
    Else
        RoundingValue = 0.05
    End If
    
    myString = "RA: " & Format(Fix(RA + (RoundingValue / 3600)), "00") & "h "
    
    RA = (RA - Fix(RA + (RoundingValue / 3600))) * 60
    
    myString = myString & Format(Fix(RA + (RoundingValue / 60)), "00") & "m "
    
    RA = (RA - Fix(RA + (RoundingValue / 60))) * 60
    
    If Not LowPrecision Then
        myString = myString & Format(RA + RoundingValue, "00.0") & "s "
    End If
       
    ConvertRAToString = myString
End Function

Public Function ConvertAltAzToString(ByVal Alt As Double, ByVal Az As Double, LowPrecision As Boolean) As String
    Dim myString As String
    
    myString = "Alt: " & Format(Fix(Alt + (0.5 / 3600)), "00") & Chr(176)
    
    Alt = (Alt - Fix(Alt + (0.5 / 3600))) * 60
    
    myString = myString & Format(Fix(Alt + (0.5 / 60)), "00") & "' "
    
    Alt = (Alt - Fix(Alt + (0.5 / 60))) * 60
    
    If Not LowPrecision Then
        myString = myString & Format(Alt, "00.0") & Chr(34)
    End If
    
    myString = myString & " Az: "
    
    myString = myString & Format(Fix(Az + (0.5 / 3600)), "00") & Chr(176)
    
    Az = (Az - Fix(Az + (0.5 / 3600))) * 60
    
    myString = myString & Format(Fix(Az + (0.5 / 60)), "00") & "'"
    
    Az = (Az - Fix(Az + (0.5 / 60))) * 60
    
    If Not LowPrecision Then
        myString = myString & Format(Az, "00") & Chr(34)
    End If
    
    ConvertAltAzToString = myString
End Function

Public Function FormatRAForFITSHeader(ByVal RA As Double, Optional NoSpace As Boolean = False) As String
    Dim myString As String
    
    myString = Format(Fix(RA + (0.5 / 3600)), "00")
    If Not NoSpace Then
        myString = myString & " "
    Else
        myString = myString & "h"
    End If
    RA = (RA - Fix(RA)) * 60
    myString = myString & Format(Fix(RA + (0.5 / 60)), "00")
    If Not NoSpace Then
        myString = myString & " "
    Else
        myString = myString & "m"
    End If
    RA = (RA - Fix(RA)) * 60
    myString = myString & Format(RA, "00.0")
    If NoSpace Then
        myString = myString & "s"
    End If
    
    FormatRAForFITSHeader = myString
End Function

Public Function FormatDecForFITSHeader(ByVal Dec As Double, Optional NoSpace As Boolean = False) As String
    Dim myString As String
    
    If Dec < 0 Then
        myString = "-"
        Dec = Abs(Dec)
    Else
        myString = "+"
    End If
    myString = myString & Format(Fix(Dec + (0.5 / 3600)), "00")
    If Not NoSpace Then
        myString = myString & " "
    Else
        myString = myString & "d"
    End If
    Dec = (Dec - Fix(Dec)) * 60
    myString = myString & Format(Fix(Dec + (0.5 / 60)), "00")
    If Not NoSpace Then
        myString = myString & " "
    Else
        myString = myString & "m"
    End If
    Dec = (Dec - Fix(Dec)) * 60
    myString = myString & Format(Dec, "00.0")
    If NoSpace Then
        myString = myString & "s"
    End If
    
    FormatDecForFITSHeader = myString
End Function

Public Function ASin(Value As Double) As Double
    If (Value = 1) Then
        ASin = PI / 2
    ElseIf (Value = -1) Then
        ASin = -PI / 2
    Else
        ASin = Atn(Value / Sqr(-Value * Value + 1))
    End If
End Function

Public Function ACos(Value As Double) As Double
    ACos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
End Function

Public Function ConvertDayToJulianDate(Year As Long, Month As Long, Day As Long) As Double
    Dim Y As Long
    Dim m As Long
    Dim jd As Double
    Dim a As Long
    Dim b As Long
    
    If Month > 2 Then
        Y = Year
        m = Month
    Else
        Y = Year - 1
        m = Month + 12
    End If
    
    jd = Int(365.25 * Y) + Int(30.6001 * (m + 1)) + Day + 1720994.5
    
    a = Int(Y / 100)
    b = 2 - a + Int(a / 4)
    
    ConvertDayToJulianDate = jd + b
End Function

Public Sub PrecessCoordinates(RA2000 As Double, Dec2000 As Double, RANow As Double, DecNow As Double)
    Dim P1 As Double
    Dim R1 As Double
    Dim FE As Double
    Dim NY As Double
    Dim T0 As Double
    Dim T1 As Double
    Dim T2 As Double
    Dim T3 As Double
    Dim W As Double
    Dim ZT As Double
    Dim ZD As Double
    Dim TH As Double
    Dim S1 As Double
    Dim S2 As Double
    Dim S3 As Double
    Dim C1 As Double
    Dim C2 As Double
    Dim C3 As Double
    Dim XX As Double
    Dim XY As Double
    Dim XZ As Double
    Dim YX As Double
    Dim YY As Double
    Dim YZ As Double
    Dim ZX As Double
    Dim ZY As Double
    Dim ZZ As Double
    Dim A0 As Double
    Dim D0 As Double
    Dim SA As Double
    Dim SD As Double
    Dim CA As Double
    Dim CD As Double
    Dim X0 As Double
    Dim X1 As Double
    Dim Y0 As Double
    Dim Y1 As Double
    Dim Z0 As Double
    Dim Z1 As Double
    Dim A1 As Double
    Dim D1 As Double
    
    Const H1 = 2306.2181
    Const H2 = 1.39656
    Const H3 = -0.000139
    Const H4 = 0.30188
    Const H5 = -0.000345
    Const H6 = 0.017998
    Const K1 = 1.09468
    Const K2 = 0.000066
    Const K3 = 0.018203
    Const L1 = 2004.3109
    Const L2 = -0.8533
    Const L3 = -0.000217
    Const L4 = -0.42665
    Const L5 = -0.000217
    Const L6 = -0.041833
    
    P1 = Misc.PI
    R1 = P1 / 180#

    '!!!!NEED ACTUAL LONGITUDE HERE!!!!
    FE = Misc.ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d"))) + (Timer / 86400) + (7.98 / 24)
    
    NY = ((FE - 2451545#) / 365.25)   'Convert to "Years since 2000"
    
    T0 = 0
    T1 = NY / 100
    T2 = T1 * T1
    T3 = T1 * T1 * T1
    
    W = (H1 + H2 * T0 + H3 * T0 * T0) * T1
    ZT = W + (H4 + H5 * T0) * T2 + H6 * T3
    ZD = W + (K1 + K2 * T0) * T2 + K3 * T3
    TH = (L1 + L2 * T0 + L3 * T0 * T0) * T1
    TH = TH + (L4 + L5 * T0) * T2 + L6 * T3
    ZT = ZT * R1 / 3600: ZD = ZD * R1 / 3600
    TH = TH * R1 / 3600

    Rem  ZT,ZD,TH = Euler angles
    Rem
    Rem  Rotation matrix
    S1 = Sin(ZT): C1 = Cos(ZT)
    S2 = Sin(ZD): C2 = Cos(ZD)
    S3 = Sin(TH): C3 = Cos(TH)
    XX = C1 * C3 * C2 - S1 * S2
    YX = -S1 * C3 * C2 - C1 * S2: ZX = -S3 * C2
    XY = C1 * C3 * S2 + S1 * C2
    YY = -S1 * C3 * S2 + C1 * C2: ZY = -S3 * S2
    XZ = C1 * S3: YZ = -S1 * S3: ZZ = C3

    A0 = RA2000 * 15 * R1
    D0 = Dec2000 * R1
    Rem
    Rem  Spherical--> rectangular
    SA = Sin(A0): CA = Cos(A0)
    SD = Sin(D0): CD = Cos(D0)
    X0 = CA * CD: Y0 = SA * CD: Z0 = SD
    Rem   3-D transformation
    X1 = X0 * XX + Y0 * YX + Z0 * ZX
    Y1 = X0 * XY + Y0 * YY + Z0 * ZY
    Z1 = X0 * XZ + Y0 * YZ + Z0 * ZZ
    Rem  Rectangular--> spherical
    A1 = Atn(Y1 / X1)
    If X1 < 0 Then A1 = A1 + P1
    If A1 < 0 Then A1 = A1 + 2 * P1
    RANow = A1 / (R1 * 15): Rem Final R.A.
    D1 = Atn(Z1 / Sqr(X1 * X1 + Y1 * Y1))
    DecNow = D1 / R1: Rem  Final Dec.
End Sub

Public Sub ConvertRADecToAltAz(ByVal RA As Double, ByVal Dec As Double, ByVal LST As Double, ByVal Lat As Double, Alt As Double, Az As Double)
    Dim HA As Double
    Dim ErrNo As Long
    
    HA = (LST - RA) * 15
    If (HA < 0) Then HA = HA + 360

    'convert degrees to radians
    HA = HA * PI / 180
    Dec = Dec * PI / 180
    Lat = Lat * PI / 180

    'compute altitude in radians
    Alt = ASin((Sin(Dec) * Sin(Lat)) + (Cos(Dec) * Cos(Lat) * Cos(HA)))
    
    'compute azimuth in radians
    'divide by zero error at poles or if alt = 90 deg
    On Error Resume Next
    Az = ACos((Sin(Dec) - (Sin(Alt) * Sin(Lat))) / (Cos(Alt) * Cos(Lat)))
    ErrNo = Err.Number
    On Error GoTo 0
    If ErrNo = 11 Then
        'Divide by 0, must be at 90 degrees
        Az = 0
        Alt = 90
    Else
        'convert radians to degrees
        Alt = Alt * 180 / PI
        Az = Az * 180 / PI
    End If
    
    'choose hemisphere
    If (Sin(HA) > 0) Then Az = 360 - Az
End Sub

Public Sub ConvertAltAzToRADec(ByVal Alt As Double, ByVal Az As Double, ByVal LST As Double, ByVal Lat As Double, RA As Double, Dec As Double)
    Dim HA As Double
    
    Alt = Alt * PI / 180
    Az = Az * PI / 180
    Lat = Lat * PI / 180
    
    Dec = ASin((Cos(Az) * Cos(Alt) * Cos(Lat)) + (Sin(Alt) * Sin(Lat)))
    
    HA = Atn360(((Cos(Lat) * Sin(Alt)) - (Sin(Lat) * Cos(Az) * Cos(Alt))), -(Sin(Az) * Cos(Alt)))
    'ha = ACos((Sin(Alt) - (Sin(Dec) * Sin(Lat))) / (Cos(Dec) * Cos(Lat)))
    
    HA = HA * 180 / PI
    RA = Misc.DoubleModulus(LST - (HA / 15), 24)
    
    Dec = Dec * 180 / PI
End Sub

Public Function Atn360(X As Double, Y As Double) As Double
    If (X > 0) And (Y > 0) Then
        Atn360 = Atn(Y / X)
    ElseIf (X < 0) And (Y > 0) Then
        Atn360 = PI + Atn(Y / X)
    ElseIf (X < 0) And (Y < 0) Then
        Atn360 = PI + Atn(Y / X)
    ElseIf (X > 0) And (Y < 0) Then
        Atn360 = (2 * PI) + Atn(Y / X)
    ElseIf (X = 0) Then
        If (Y >= 0) Then
            Atn360 = PI / 2
        Else
            Atn360 = -PI / 2
        End If
    ElseIf (Y = 0) Then
        If (X >= 0) Then
            Atn360 = 0
        Else
            Atn360 = PI
        End If
    End If
End Function

Public Function ComputeRiseTime(ByVal RA As Double, ByVal Dec As Double, ByVal Altitude As Double, ByVal Latitude As Double, ByVal Longitude As Double) As Date
    Dim JD_t As Double
    Dim JD_0 As Double
    Dim N As Double
    Dim Ti As Double
    Dim Z3 As Double
    Dim X3 As Double
    Dim LocalHourAngle As Double
    Dim UT As Double
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Long
    Dim LocalTime As Double
    Dim LocalTimeH As Integer
    Dim LocalTimeM As Integer
    Dim LocalTimeS As Integer
    Dim LocalMeanTime As Double
    Dim MyLongitude As Double
    
    MyLongitude = Longitude / 15
    If CInt(Format(Now, "hh")) > 12 Then
        JD_t = ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")) + 1)
    Else
        JD_t = ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")))
    End If
    JD_0 = ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), 1, 0)
    
    N = JD_t - JD_0
    
    Ti = N + (12 + MyLongitude) / 24
    
    If Altitude > (90 - (Latitude - Dec)) Then
        Altitude = (90 - (Latitude - Dec))
    End If
    
    Z3 = 90 - Altitude
        
    X3 = (Cos(Z3 * PI / 180) - (Sin(Dec * PI / 180) * Sin(Latitude * PI / 180))) / (Cos(Dec * PI / 180) * Cos(Latitude * PI / 180))
    
    On Error Resume Next
    LocalHourAngle = (360 - ((180 / PI) * ACos(X3))) / 15
    If Err.Number <> 0 Then
        ' Error probably means I'm trying to get above the maximum altitude (or really close to it)
        ' Just use the maximum hour angle
        LocalHourAngle = 24
    End If
    On Error GoTo 0
    
    LocalMeanTime = (LocalHourAngle + RA - (0.06571 * Ti) - 6.622) * 366.2422 / 365.2422
    LocalMeanTime = DoubleModulus(LocalMeanTime, 24)
    
    UT = DoubleModulus(LocalMeanTime + MyLongitude, 24)
    
    If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
    Else
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
    End If
    
    LocalTime = DoubleModulus(UT - (LocalBias / 60) + (0.5 / 3600), 24)
    LocalTimeH = Int(LocalTime)
    LocalTimeM = Int((LocalTime - Int(LocalTime)) * 60)
    LocalTimeS = Int((((LocalTime - Int(LocalTime)) * 60) - LocalTimeM) * 60)
        
    ComputeRiseTime = TimeSerial(LocalTimeH, LocalTimeM, LocalTimeS)
End Function

Public Function ComputeSetTime(ByVal RA As Double, ByVal Dec As Double, ByVal Altitude As Double, ByVal Latitude As Double, ByVal Longitude As Double) As Date
    Dim JD_t As Double
    Dim JD_0 As Double
    Dim N As Double
    Dim Ti As Double
    Dim Z0 As Double
    Dim X0 As Double
    Dim LocalHourAngle As Double
    Dim UT As Double
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    Dim LocalBias As Long
    Dim LocalTime As Double
    Dim LocalTimeH As Integer
    Dim LocalTimeM As Integer
    Dim LocalTimeS As Integer
    Dim LocalMeanTime As Double
    Dim MyLongitude As Double
    
    MyLongitude = Longitude / 15

    If CInt(Format(Time, "hh")) > 12 Then
        JD_t = ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")) + 1)
    Else
        JD_t = ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), CLng(Format(Now, "m")), CLng(Format(Now, "d")))
    End If
    JD_0 = ConvertDayToJulianDate(CLng(Format(Now, "yyyy")), 1, 0)
    
    N = JD_t - JD_0
    
    Ti = N + ((12 + MyLongitude) / 24)
    
    If Altitude > (90 - (Latitude - Dec)) Then
        Altitude = (90 - (Latitude - Dec))
    End If
    
    Z0 = 90 - Altitude
    
    X0 = (Cos(Z0 * PI / 180) - (Sin(Dec * PI / 180) * Sin(Latitude * PI / 180))) / (Cos(Dec * PI / 180) * Cos(Latitude * PI / 180))
    
    On Error Resume Next
    LocalHourAngle = (180 / PI) * ACos(X0) / 15
    If Err.Number <> 0 Then
        ' Error probably means I'm trying to get above the maximum altitude (or really close to it)
        ' Just use the maximum hour angle
        LocalHourAngle = 24
    End If
    On Error GoTo 0
        
    LocalMeanTime = (LocalHourAngle + RA - (0.06571 * Ti) - 6.622) * 366.2422 / 365.2422
    LocalMeanTime = DoubleModulus(LocalMeanTime, 24)
    
    UT = DoubleModulus(LocalMeanTime + MyLongitude, 24)
    
    If GetTimeZoneInformation(TimeZoneInfo) = cstDSTType.cstDSTDaylight Then
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.DaylightBias
    Else
        LocalBias = TimeZoneInfo.Bias + TimeZoneInfo.StandardBias
    End If
    
    LocalTime = DoubleModulus(UT - (LocalBias / 60) + (0.5 / 3600), 24)
    LocalTimeH = Int(LocalTime)
    LocalTimeM = Int((LocalTime - Int(LocalTime)) * 60)
    LocalTimeS = Int((((LocalTime - Int(LocalTime)) * 60) - LocalTimeM) * 60)
    
    ComputeSetTime = TimeSerial(LocalTimeH, LocalTimeM, LocalTimeS)
End Function

Public Sub CreatePath(myPath As String)
    Dim SubPath As String
    
    SubPath = Left(myPath, InStrRev(myPath, "\", Len(myPath) - 1))
    
    If SubPath = "" Then
        Err.Raise &H80000FFF, "CCD Commander", "Cannot create directory: " & myPath
    End If
    
    On Error Resume Next
    Call ChDir(SubPath)
    If Err.Number <> 0 Then
        On Error GoTo 0
        'Can't create this one - just recurse until I find one that exists
        Call CreatePath(SubPath)
        'Now I should be able to create this path
    End If
    
    On Error GoTo 0
    'Have all but the last directory
    Call MkDir(myPath)
    Call ChDir(myPath)
End Sub
