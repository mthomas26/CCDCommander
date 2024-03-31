VERSION 5.00
Begin VB.Form frmWaitForAlt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wait for Altitude Action"
   ClientHeight    =   5610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4050
   HelpContextID   =   1100
   Icon            =   "frmWaitForAlt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdComputeTime 
      Caption         =   "Compute Approximate Wait Time"
      Height          =   855
      Left            =   2760
      TabIndex        =   32
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Frame fraObjectType 
      Caption         =   "Object Type"
      Height          =   855
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   2535
      Begin VB.OptionButton optSun 
         Caption         =   "Sun"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton optMoon 
         Caption         =   "Moon"
         Height          =   255
         Left            =   1260
         TabIndex        =   30
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   540
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.OptionButton optDirection 
      Caption         =   "Setting"
      Height          =   315
      Index           =   1
      Left            =   1388
      TabIndex        =   11
      Top             =   4680
      Width           =   1275
   End
   Begin VB.OptionButton optDirection 
      Caption         =   "Rising"
      Height          =   315
      Index           =   0
      Left            =   1388
      TabIndex        =   10
      Top             =   4380
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtAlt 
      Height          =   312
      Left            =   2280
      TabIndex        =   9
      Text            =   "45"
      Top             =   3900
      Width           =   312
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraObjectCoordinates 
      Caption         =   "Object Coordinates"
      Height          =   1752
      Left            =   60
      TabIndex        =   16
      Top             =   960
      Width           =   2532
      Begin VB.TextBox txtRAH 
         Height          =   312
         Left            =   600
         TabIndex        =   0
         Text            =   "23"
         Top             =   240
         Width           =   312
      End
      Begin VB.TextBox txtRAM 
         Height          =   312
         Left            =   1200
         TabIndex        =   1
         Text            =   "10"
         Top             =   240
         Width           =   312
      End
      Begin VB.TextBox txtRAS 
         Height          =   312
         Left            =   1860
         TabIndex        =   2
         Text            =   "58"
         Top             =   240
         Width           =   312
      End
      Begin VB.TextBox txtDecD 
         Height          =   312
         Left            =   600
         TabIndex        =   3
         Text            =   "5"
         Top             =   720
         Width           =   312
      End
      Begin VB.TextBox txtDecM 
         Height          =   312
         Left            =   1200
         TabIndex        =   4
         Text            =   "23"
         Top             =   720
         Width           =   312
      End
      Begin VB.TextBox txtDecS 
         Height          =   312
         Left            =   1860
         TabIndex        =   5
         Text            =   "15"
         Top             =   720
         Width           =   312
      End
      Begin VB.CommandButton cmdGetFromSky 
         Caption         =   "Get RA/Dec from TheSky"
         Height          =   432
         Left            =   180
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RA:"
         Height          =   192
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   264
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "h"
         Height          =   192
         Left            =   960
         TabIndex        =   23
         Top             =   300
         Width           =   84
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   192
         Left            =   1560
         TabIndex        =   22
         Top             =   300
         Width           =   132
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   192
         Left            =   2220
         TabIndex        =   21
         Top             =   300
         Width           =   84
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dec:"
         Height          =   192
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   336
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "d"
         Height          =   192
         Left            =   960
         TabIndex        =   19
         Top             =   780
         Width           =   96
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   192
         Left            =   1560
         TabIndex        =   18
         Top             =   780
         Width           =   132
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   192
         Left            =   2220
         TabIndex        =   17
         Top             =   780
         Width           =   84
      End
   End
   Begin VB.Frame fraObjectName 
      Height          =   1215
      Left            =   60
      TabIndex        =   26
      Top             =   2595
      Width           =   3915
      Begin VB.TextBox txtObjectName 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdFindObject 
         Caption         =   "Find Object in TheSky"
         Height          =   432
         Left            =   900
         TabIndex        =   8
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Object Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Label lblMeridian 
      AutoSize        =   -1  'True
      Caption         =   "If the object crosses the meridian, the waiting will stop, regardless of the alititude set above."
      Height          =   390
      Left            =   180
      TabIndex        =   25
      Top             =   5100
      Width           =   3630
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "degrees"
      Height          =   195
      Left            =   2700
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Wait until object is at altitude:"
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   3960
      Width           =   2025
   End
End
Attribute VB_Name = "frmWaitForAlt"
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

Private Const RegistryName = "WaitForAlt"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub cmdComputeTime_Click()
    Dim SkipTime As Date
    Dim tempTime As Date
    Dim RA As Double
    Dim Dec As Double
    
    If Me.optSun.Value Then
        If Me.optDirection(0).Value Then
            'Rising
            Call Mount.ComputeTwilightStartTime(CDbl(Me.txtAlt.Text))
            
            If CInt(Format(Time, "hh")) > 12 And CInt(Format(Mount.TwilightStartTime, "hh")) < 12 Then
                tempTime = Format(Date + 1, "Short Date") & " " & Mount.TwilightStartTime
            Else
                tempTime = Format(Date, "Short Date") & " " & Mount.TwilightStartTime
            End If
        Else
            Call Mount.ComputeSunSetTime(CDbl(Me.txtAlt.Text))
            tempTime = Format(Date, "Short Date") & " " & Mount.SunSetTime
        End If
    ElseIf Me.optMoon.Value Then
        If Me.optDirection(0).Value Then
            'Rising
            Call AstroFunctions.MoonRise(CDbl(Me.txtAlt.Text))
            tempTime = AstroFunctions.MoonRiseTime
        Else
            Call AstroFunctions.Moonset(CDbl(Me.txtAlt.Text))
            tempTime = AstroFunctions.MoonSetTime
        End If
    Else
        RA = CDbl(Me.txtRAH.Text) + (CDbl(Me.txtRAM.Text) / 60) + (CDbl(Me.txtRAS.Text) / 3600)
        Dec = CDbl(Me.txtDecD.Text) + (CDbl(Me.txtDecM.Text) / 60) + (CDbl(Me.txtDecS.Text) / 3600)
        
        If Me.optDirection(0).Value Then
            'Rising
            tempTime = Misc.ComputeRiseTime(RA, Dec, CDbl(Me.txtAlt.Text), Mount.GetLatitude, Mount.GetLongitude)
        Else
            tempTime = Misc.ComputeSetTime(RA, Dec, CDbl(Me.txtAlt.Text), Mount.GetLatitude, Mount.GetLongitude)
        End If
    End If
    
    'Add 30s to time - this will round the time to the next minute
    tempTime = DateAdd("s", 30, tempTime)
    
    If Hour(tempTime) < 12 And CInt(Format(Time, "hh")) >= 12 Then
        SkipTime = (Format(DateAdd("d", 1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
    ElseIf Hour(tempTime) >= 12 And CInt(Format(Time, "hh")) < 12 Then
        SkipTime = (Format(DateAdd("d", -1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
    Else
        SkipTime = Format(Date, "Short Date") & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
    End If
    
    MsgBox "Wait will finish at about " & Format(SkipTime, "Short Time") & ".", vbInformation
End Sub

Private Sub cmdFindObject_Click()
    Dim RA As Double
    Dim Dec As Double
    
    Call Planetarium.GetObjectRADec(Me.txtObjectName.Text, RA, Dec)
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
    Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
    
    If Dec < 0 And Dec > -1 Then
        Me.txtDecD.Text = "-0"
    Else
        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
    End If
    Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
    Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
End Sub

Private Sub cmdGetFromSky_Click()
    Dim RA As Double
    Dim Dec As Double
    
    Me.txtObjectName.Text = Planetarium.GetSelectedObject(RA, Dec, False)
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
    Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
    
    If Dec < 0 And Dec > -1 Then
        Me.txtDecD.Text = "-" & Fix(Dec + (0.5 / 3600))
    Else
        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
    End If
    Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
    Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    If frmOptions.lstPlanetarium.ListIndex = 0 Then
        Me.cmdGetFromSky.Enabled = False
        Me.cmdFindObject.Enabled = False
    Else
        Me.cmdGetFromSky.Enabled = True
        Me.cmdFindObject.Enabled = True
    End If
    
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Tag = "False"
        Me.Hide
        Call GetSettings
    End If
End Sub

Private Sub GetSettings()
    Me.txtRAH.Text = GetMySetting(RegistryName, "RAH", "0")
    Me.txtRAM.Text = GetMySetting(RegistryName, "RAM", "0")
    Me.txtRAS.Text = GetMySetting(RegistryName, "RAS", "0")
    Me.txtDecD.Text = GetMySetting(RegistryName, "DecD", "0")
    Me.txtDecM.Text = GetMySetting(RegistryName, "DecM", "0")
    Me.txtDecS.Text = GetMySetting(RegistryName, "DecS", "0")
    Me.txtObjectName.Text = GetMySetting(RegistryName, "ObjectName", "")
    
    If Me.txtObjectName.Text = "Sun" Then
        Me.optSun.Value = True
    ElseIf Me.txtObjectName.Text = "Moon" Then
        Me.optMoon.Value = True
    Else
        Me.optOther.Value = True
    End If
    
    Me.txtAlt.Text = GetMySetting(RegistryName, "Alt", "45")
    If CBool(GetMySetting(RegistryName, "Rising", "True")) Then
        Me.optDirection(0).Value = True
    Else
        Me.optDirection(1).Value = True
    End If
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtRAH_Validate(Cancel)
    If Not Cancel Then Call txtRAM_Validate(Cancel)
    If Not Cancel Then Call txtRAS_Validate(Cancel)
    If Not Cancel Then Call txtDecD_Validate(Cancel)
    If Not Cancel Then Call txtDecM_Validate(Cancel)
    If Not Cancel Then Call txtDecS_Validate(Cancel)
    If Not Cancel Then Call txtAlt_Validate(Cancel)
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "RAH", Me.txtRAH.Text)
    Call SaveMySetting(RegistryName, "RAM", Me.txtRAM.Text)
    Call SaveMySetting(RegistryName, "RAS", Me.txtRAS.Text)
    Call SaveMySetting(RegistryName, "DecD", Me.txtDecD.Text)
    Call SaveMySetting(RegistryName, "DecM", Me.txtDecM.Text)
    Call SaveMySetting(RegistryName, "DecS", Me.txtDecS.Text)
    Call SaveMySetting(RegistryName, "ObjectName", Me.txtObjectName.Text)
    Call SaveMySetting(RegistryName, "Alt", Me.txtAlt.Text)
    Call SaveMySetting(RegistryName, "Rising", Me.optDirection(0).Value)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub optDirection_Click(Index As Integer)
    If Index = 0 Then
        Me.lblMeridian.Visible = True
    Else
        Me.lblMeridian.Visible = False
    End If
End Sub

Private Sub optMoon_Click()
    Me.fraObjectCoordinates.Enabled = False
    Me.txtRAH.Text = "0"
    Me.txtRAM.Text = "0"
    Me.txtRAS.Text = "0"
    Me.txtDecD.Text = "0"
    Me.txtDecM.Text = "0"
    Me.txtDecS.Text = "0"
    Me.fraObjectName.Enabled = False
    Me.txtObjectName.Text = "Moon"

End Sub

Private Sub optOther_Click()
    Me.fraObjectCoordinates.Enabled = True
    Me.fraObjectName.Enabled = True

End Sub

Private Sub optSun_Click()
    Me.fraObjectCoordinates.Enabled = False
    Me.txtRAH.Text = "0"
    Me.txtRAM.Text = "0"
    Me.txtRAS.Text = "0"
    Me.txtDecD.Text = "0"
    Me.txtDecM.Text = "0"
    Me.txtDecS.Text = "0"
    Me.fraObjectName.Enabled = False
    Me.txtObjectName.Text = "Sun"

End Sub

Private Sub txtAlt_GotFocus()
    Me.txtAlt.SelStart = 0
    Me.txtAlt.SelLength = Len(Me.txtAlt.Text)
End Sub

Private Sub txtAlt_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAlt.Text)
    If Err.Number <> 0 Or Test <= -90 Or Test >= 90 Or Test <> Me.txtAlt.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtRAH_GotFocus()
    Me.txtRAH.SelStart = 0
    Me.txtRAH.SelLength = Len(Me.txtRAH.Text)
End Sub

Private Sub txtRAH_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtRAH.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 24 Or Test <> Me.txtRAH.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtRAM_GotFocus()
    Me.txtRAM.SelStart = 0
    Me.txtRAM.SelLength = Len(Me.txtRAM.Text)
End Sub

Private Sub txtRAM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtRAM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtRAM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtRAS_GotFocus()
    Me.txtRAS.SelStart = 0
    Me.txtRAS.SelLength = Len(Me.txtRAS.Text)
End Sub

Private Sub txtRAS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtRAS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtRAS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDecD_GotFocus()
    Me.txtDecD.SelStart = 0
    Me.txtDecD.SelLength = Len(Me.txtDecD.Text)
End Sub

Private Sub txtDecD_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDecD.Text)
    If Err.Number <> 0 Or Test < -90 Or Test > 90 Or Test <> Me.txtDecD.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDecM_GotFocus()
    Me.txtDecM.SelStart = 0
    Me.txtDecM.SelLength = Len(Me.txtDecM.Text)
End Sub

Private Sub txtDecM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDecM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtDecM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDecS_GotFocus()
    Me.txtDecS.SelStart = 0
    Me.txtDecS.SelLength = Len(Me.txtDecS.Text)
End Sub

Private Sub txtDecS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDecS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtDecS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Public Sub PutFormDataIntoClass(clsAction As WaitForAltAction)
    clsAction.RA = CDbl(Me.txtRAH.Text) + (CDbl(Me.txtRAM.Text) / 60) + (CDbl(Me.txtRAS.Text) / 3600)
    clsAction.Dec = Abs(CDbl(Me.txtDecD.Text)) + (CDbl(Me.txtDecM.Text) / 60) + (CDbl(Me.txtDecS.Text) / 3600)
    If InStr(Me.txtDecD.Text, "-") <> 0 Then
        clsAction.Dec = -clsAction.Dec
    End If
    clsAction.Name = Me.txtObjectName.Text
    clsAction.Alt = CDbl(Me.txtAlt.Text)
    clsAction.Rising = Me.optDirection(0).Value
End Sub

Public Sub GetFormDataFromClass(clsAction As WaitForAltAction)
    Dim RA As Double
    Dim Dec As Double
    
    RA = clsAction.RA
    Dec = clsAction.Dec
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    RA = (RA - Fix(RA + (0.5 / 3600))) * 60
    Me.txtRAM.Text = Fix(RA + (0.5 / 60))
    RA = (RA - Fix(RA + (0.5 / 60))) * 60
    Me.txtRAS.Text = Fix(RA + 0.5)

    If Dec < 0 And Dec > -1 Then
        Dec = -Dec
        Me.txtDecD.Text = "-0"
    ElseIf Dec < 0 Then
        Dec = -Dec
        Me.txtDecD.Text = -Fix(Dec + (0.5 / 3600))
    Else
        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
    End If
    Dec = (Dec - Fix(Dec + (0.5 / 3600))) * 60
    Me.txtDecM.Text = Fix(Dec + (0.5 / 60))
    Dec = (Dec - Fix(Dec + (0.5 / 60))) * 60
    Me.txtDecS.Text = Fix(Dec + 0.5)
    
    Me.txtObjectName.Text = clsAction.Name
    
    Me.txtAlt.Text = clsAction.Alt
    
    If clsAction.Rising Then
        Me.optDirection(0).Value = True
    Else
        Me.optDirection(1).Value = True
    End If

    If Me.txtObjectName.Text = "Sun" Then
        Me.optSun.Value = True
    ElseIf Me.txtObjectName.Text = "Moon" Then
        Me.optMoon.Value = True
    Else
        Me.optOther.Value = True
    End If
    
    Call SaveMySetting(RegistryName, "RAH", Me.txtRAH.Text)
    Call SaveMySetting(RegistryName, "RAM", Me.txtRAM.Text)
    Call SaveMySetting(RegistryName, "RAS", Me.txtRAS.Text)
    Call SaveMySetting(RegistryName, "DecD", Me.txtDecD.Text)
    Call SaveMySetting(RegistryName, "DecM", Me.txtDecM.Text)
    Call SaveMySetting(RegistryName, "DecS", Me.txtDecS.Text)
    Call SaveMySetting(RegistryName, "Alt", Me.txtAlt.Text)
    Call SaveMySetting(RegistryName, "ObjectName", Me.txtObjectName.Text)
    Call SaveMySetting(RegistryName, "Rising", Me.optDirection(0).Value)
End Sub

