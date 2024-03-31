VERSION 5.00
Begin VB.Form frmFileNameBuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Name Builder"
   ClientHeight    =   2490
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8160
   HelpContextID   =   2100
   Icon            =   "frmFileNameBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Clear Filename"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   900
      Width           =   1875
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to Filename"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   1875
   End
   Begin VB.ListBox lstParameters 
      Height          =   1230
      Left            =   1140
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtFileNamePrefix 
      Height          =   285
      Left            =   900
      TabIndex        =   2
      Top             =   1680
      Width           =   7095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6780
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6780
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Parameters:"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblExample 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   900
      TabIndex        =   5
      Top             =   2100
      Width           =   7065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Example:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   2100
      Width           =   645
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Filename:"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   1740
      Width           =   675
   End
End
Attribute VB_Name = "frmFileNameBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Sub CancelButton_Click()
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub cmdAdd_Click()
    If Me.lstParameters.ListIndex = -1 Then
        MsgBox "You must select a parameter first!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    Me.txtFileNamePrefix.Text = Me.txtFileNamePrefix.Text & GetParameterString(Me.lstParameters.List(Me.lstParameters.ListIndex))
End Sub

Private Sub cmdRemove_Click()
    Me.txtFileNamePrefix.Text = ""
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
        
    Call Me.lstParameters.AddItem("Date")
    Call Me.lstParameters.AddItem("Date (UT)")
    Call Me.lstParameters.AddItem("Time (UT)")
    Call Me.lstParameters.AddItem("Time")
    Call Me.lstParameters.AddItem("Exposure Time")
    Call Me.lstParameters.AddItem("Binning")
    Call Me.lstParameters.AddItem("Filter Name")
    Call Me.lstParameters.AddItem("Image Type")
    Call Me.lstParameters.AddItem("Object Name")
    Call Me.lstParameters.AddItem("Object Coordinates")
    Call Me.lstParameters.AddItem("Position Angle")
    Call Me.lstParameters.AddItem("Rotator Angle")
    Call Me.lstParameters.AddItem("Temperature Set Point")
End Sub

Private Function GetParameterString(ItemString As String) As String
    Select Case ItemString
        Case "Date"
            GetParameterString = "<Date>"
        Case "Time"
            GetParameterString = "<Time>"
        Case "Date (UT)"
            GetParameterString = "<DateUT>"
        Case "Time (UT)"
            GetParameterString = "<UT>"
        Case "Exposure Time"
            GetParameterString = "<ExposureTime>"
        Case "Binning"
            GetParameterString = "<Bin>"
        Case "Filter Name"
            GetParameterString = "<Filter>"
        Case "Image Type"
            GetParameterString = "<ImageType>"
        Case "Object Name"
            GetParameterString = "<ObjectName>"
        Case "Object Coordinates"
            GetParameterString = "<ObjectCoords>"
        Case "Position Angle"
            GetParameterString = "<PA>"
        Case "Rotator Angle"
            GetParameterString = "<RotatorAngle>"
        Case "Temperature Set Point"
            GetParameterString = "<Temperature>"
    End Select
End Function

Private Sub lstParameters_DblClick()
    Call cmdAdd_Click
End Sub

Private Sub OKButton_Click()
    Me.Tag = Me.txtFileNamePrefix.Text
    Me.Hide
End Sub

Private Sub txtFileNamePrefix_Change()
    Me.lblExample.Caption = FixFileNameExample(Me.txtFileNamePrefix.Text)
End Sub

Private Function FixFileNameExample(DesiredName As String) As String
    Dim Counter As Long
    Dim FixedFileName As String
    Dim TestChar As String * 1
    Dim Parameter As String
    
    For Counter = 1 To Len(DesiredName)
        TestChar = Mid(DesiredName, Counter, 1)
        If TestChar = "<" Then
            'Maybe start of a naming parameter!
            If InStr(Counter, DesiredName, ">") > 0 Then
                'Got a parameter!
                If (Right(FixedFileName, 1) <> "_") And Len(FixedFileName) > 0 Then
                    FixedFileName = FixedFileName & "_"
                End If
                
                Parameter = Mid(DesiredName, Counter + 1, InStr(Counter, DesiredName, ">") - Counter - 1)
                Select Case UCase(Parameter)
                    Case UCase("Date")
                        FixedFileName = FixedFileName & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
                    Case UCase("DateUT")
                        FixedFileName = FixedFileName & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
                    Case UCase("Time")
                        FixedFileName = FixedFileName & Format(Now, "hhmm")
                    Case UCase("UT")
                        FixedFileName = FixedFileName & Format(Now, "hhmm")
                    Case UCase("ExposureTime")
                        FixedFileName = FixedFileName & "1200s"
                    Case UCase("Bin")
                        FixedFileName = FixedFileName & "1x1"
                    Case UCase("Filter")
                        If (frmOptions.lstFilters.ListCount > 0) Then
                            FixedFileName = FixedFileName & frmOptions.lstFilters.List(0)
                        Else
                            'ignore this one
                        End If
                    Case UCase("ImageType")
                        FixedFileName = FixedFileName & "Light"
                    Case UCase("ObjectName")
                        FixedFileName = FixedFileName & "M33"
                    Case UCase("ObjectCoords")
                        FixedFileName = FixedFileName & Misc.FormatRAForFITSHeader(1.57195, True) & "_" & Misc.FormatDecForFITSHeader(30.7, True)
                    Case UCase("PA")
                        FixedFileName = FixedFileName & "28degN"
                    Case UCase("RotatorAngle")
                        FixedFileName = FixedFileName & "126deg"
                    Case UCase("Temperature")
                        FixedFileName = FixedFileName & "-45degC"
                End Select
                
                FixedFileName = FixedFileName & "_"
                
                Counter = InStr(Counter, DesiredName, ">")
            Else
                'No parameter end, ignore
            End If
        ElseIf TestChar = "\" Or TestChar = "/" Or TestChar = ":" Or TestChar = "*" Or TestChar = "?" Or TestChar = Chr(34) _
            Or TestChar = "<" Or TestChar = ">" Or TestChar = "|" Or TestChar = "," Then
            
            'ignore character
        Else
            FixedFileName = FixedFileName & TestChar
        End If
    Next Counter
        
    If Right(FixedFileName, 1) = "_" Then
        FixedFileName = Left(FixedFileName, Len(FixedFileName) - 1)
    End If
    
    FixFileNameExample = FixedFileName
        
End Function


