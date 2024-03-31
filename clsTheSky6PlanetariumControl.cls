VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheSky6Planetarium"
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

Public Function GetFOVIPositionAngle() As Double
    Dim myFOVIs As New TheSky6Library.MyFOVs
    Dim Counter As Integer
    
    For Counter = 0 To myFOVIs.Count - 1
        If myFOVIs.Property(myFOVIs.Name(Counter), 0, skMyFOVProp_Visible) = 1 Then
            Exit For
        End If
    Next Counter
    
    If Counter >= myFOVIs.Count Then
        MsgBox "You must have at least one FOVI enabled in TheSky.", vbExclamation
        GetFOVIPositionAngle = 0
    Else
        GetFOVIPositionAngle = myFOVIs.Property(myFOVIs.Name(Counter), 0, skMyFOVProp_PositionAngleDegrees)
    End If
End Function

Public Sub SetFOVIPositionAngle(Value As Double)
    Dim myFOVIs As New TheSky6Library.MyFOVs
    Dim Counter As Integer
    
    For Counter = 0 To myFOVIs.Count - 1
        If myFOVIs.Property(myFOVIs.Name(Counter), 0, skMyFOVProp_Visible) = 1 Then
            Exit For
        End If
    Next Counter
    
    If Counter >= myFOVIs.Count Then
        'just bail out
        Exit Sub
    Else
        myFOVIs.Property(myFOVIs.Name(Counter), 0, skMyFOVProp_PositionAngleDegrees) = Value
    End If
End Sub

Public Sub GetObjectRADec(ObjectName As String, RA As Double, Dec As Double, Optional J2000 As Boolean = False)
    Dim TheSkyStarChart As New TheSky6Library.StarChart
    Dim ObjectInfo As TheSky6Library.ObjectInformation
    Dim Counter As Integer
    Dim MyObjectName As String
    Dim MPLObjectName As String
    
    On Error Resume Next
    Call TheSkyStarChart.UpdateSolarSystem
    Set ObjectInfo = TheSkyStarChart.Find(ObjectName)
    If Err.Number <> 0 Then
        RA = 0
        Dec = 0
    Else
        If ObjectInfo.Count > 1 Then
            'Find returned more than one object.  Try to determine the right one
            MyObjectName = UCase(ObjectName)
            
            'remove any leading spaces
            Do While (Left(MyObjectName, 1) = " ")
                MyObjectName = Mid(MyObjectName, 2)
            Loop
            
            'remove any trailing spaces
            Do While (Right(MyObjectName, 1) = " ")
                MyObjectName = Left(MyObjectName, Len(MyObjectName) - 1)
            Loop
            
            For Counter = 0 To ObjectInfo.Count - 1
                ObjectInfo.Index = Counter
                If UCase(ObjectInfo.Property(skObjInfoProp_NAME1)) = MyObjectName Then
                    'these have the same name, use this one
                    Exit For
                End If
                
                If InStr(UCase(ObjectName), "MPL") Then
                    'MPL object names don't necessarily match up correctly with the name returned by TheSky
                    'Need to use substrings to check...
                    MPLObjectName = UCase(Mid(ObjectName, InStr(UCase(ObjectName), "MPL") + 3))
                    
                    'remove any leading spaces
                    Do While (Left(MPLObjectName, 1) = " ")
                        MPLObjectName = Mid(MPLObjectName, 2)
                    Loop
                    
                    If InStr(UCase(ObjectInfo.Property(skObjInfoProp_NAME1)), MPLObjectName) Then
                        'this looks like a match, roll with it
                        Exit For
                    End If
                End If
                
                If Counter = ObjectInfo.Count - 1 Then
                    'couldn't find any match - use the first object returned
                    ObjectInfo.Index = 0
                End If
            Next Counter
        End If
        
        If J2000 = False Then
            RA = ObjectInfo.Property(skObjInfoProp_RA_NOW)
            Dec = ObjectInfo.Property(skObjInfoProp_DEC_NOW)
        Else
            RA = ObjectInfo.Property(skObjInfoProp_RA_2000)
            Dec = ObjectInfo.Property(skObjInfoProp_DEC_2000)
        End If
    End If
    On Error GoTo 0
End Sub

Public Function GetSelectedObject(RA As Double, Dec As Double, Optional J2000 As Boolean = False) As String
    Dim ObjectInfo As New TheSky6Library.ObjectInformation
    Dim Counter As Integer
    
    If (ObjectInfo.Count > 1) Then
        Load frmObjectSelect
        Do
            frmObjectSelect.lstObjects.Clear
            For Counter = 0 To ObjectInfo.Count - 1
                ObjectInfo.Index = Counter
                frmObjectSelect.lstObjects.AddItem ObjectInfo.Property(skObjInfoProp_NAME1)
            Next Counter
            
            frmObjectSelect.lstObjects.Selected(0) = True
            
            frmObjectSelect.Show vbModal, frmMain
        Loop While frmObjectSelect.Tag = "1"
        
        ObjectInfo.Index = frmObjectSelect.lstObjects.ListIndex
        
        Unload frmObjectSelect
    Else
        ObjectInfo.Index = 0
    End If
    
    If J2000 = False Then
        RA = ObjectInfo.Property(skObjInfoProp_RA_NOW)
        Dec = ObjectInfo.Property(skObjInfoProp_DEC_NOW)
    Else
        RA = ObjectInfo.Property(skObjInfoProp_RA_2000)
        Dec = ObjectInfo.Property(skObjInfoProp_DEC_2000)
    End If
    
    GetSelectedObject = ObjectInfo.Property(skObjInfoProp_NAME1)
End Function



