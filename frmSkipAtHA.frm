VERSION 5.00
Begin VB.Form frmSkipAtHA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Skip At Hour Angle Action"
   ClientHeight    =   4335
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4050
   HelpContextID   =   2200
   Icon            =   "frmSkipAtHA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSoftSkip 
      Caption         =   "Soft Skip"
      Height          =   195
      Left            =   1440
      TabIndex        =   19
      Top             =   3900
      Width           =   1000
   End
   Begin VB.CommandButton cmdComputeTime 
      Caption         =   "Compute Approximate Skip Time"
      Height          =   855
      Left            =   2760
      TabIndex        =   17
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtHA 
      Alignment       =   1  'Right Justify
      Height          =   312
      Left            =   2580
      TabIndex        =   6
      Text            =   "23.1234"
      Top             =   3420
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraObjectCoordinates 
      Caption         =   "Object Coordinates"
      Height          =   1752
      Left            =   60
      TabIndex        =   10
      Top             =   120
      Width           =   2532
      Begin VB.TextBox txtRAH 
         Height          =   312
         Left            =   600
         TabIndex        =   0
         Text            =   "23"
         Top             =   480
         Width           =   312
      End
      Begin VB.TextBox txtRAM 
         Height          =   312
         Left            =   1200
         TabIndex        =   1
         Text            =   "10"
         Top             =   480
         Width           =   312
      End
      Begin VB.TextBox txtRAS 
         Height          =   312
         Left            =   1860
         TabIndex        =   2
         Text            =   "58"
         Top             =   480
         Width           =   312
      End
      Begin VB.CommandButton cmdGetFromSky 
         Caption         =   "Get RA from TheSky"
         Height          =   432
         Left            =   180
         TabIndex        =   3
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RA:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "h"
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Top             =   540
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   540
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   195
         Left            =   2220
         TabIndex        =   11
         Top             =   540
         Width           =   90
      End
   End
   Begin VB.Frame fraObjectName 
      Height          =   1215
      Left            =   60
      TabIndex        =   15
      Top             =   2040
      Width           =   3915
      Begin VB.TextBox txtObjectName 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdFindObject 
         Caption         =   "Find Object in TheSky"
         Height          =   432
         Left            =   900
         TabIndex        =   5
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Object Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "hours"
      Height          =   195
      Left            =   3360
      TabIndex        =   18
      Top             =   3480
      Width           =   390
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Skip when object is at Hour Angle:"
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   3480
      Width           =   2445
   End
End
Attribute VB_Name = "frmSkipAtHA"
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

Private Const RegistryName = "SkipAheadAtHA"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub cmdComputeTime_Click()
    Dim SkipTime As Date
    Dim tempTime As Date
    Dim TargetLST As Double
    Dim LSTDiff As Double
    Dim RA As Double
    Dim HA As Double
    
    RA = CDbl(Me.txtRAH.Text) + (CDbl(Me.txtRAM.Text) / 60) + (CDbl(Me.txtRAS.Text) / 3600)
    HA = CDbl(Me.txtHA.Text)
    
    TargetLST = HA + RA
    
    LSTDiff = TargetLST - Mount.GetSiderealTime()
    
    tempTime = DateAdd("s", LSTDiff * 3600, Now)
    
    'Add 30s to time - this will round the time to the next minute
    tempTime = DateAdd("s", 30, tempTime)
    
    If Hour(tempTime) < 12 And CInt(Format(Time, "hh")) >= 12 Then
        SkipTime = (Format(DateAdd("d", 1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
    ElseIf Hour(tempTime) >= 12 And CInt(Format(Time, "hh")) < 12 Then
        SkipTime = (Format(DateAdd("d", -1, Date), "Short Date")) & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
    Else
        SkipTime = Format(Date, "Short Date") & " " & Hour(tempTime) & ":" & Minute(tempTime) & ":00"
    End If
    
    MsgBox "Skip will activate at about " & Format(SkipTime, "Short Time") & ".", vbInformation
End Sub

Private Sub cmdFindObject_Click()
    Dim RA As Double
    Dim Dec As Double
    
    Call Planetarium.GetObjectRADec(Me.txtObjectName.Text, RA, Dec)
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
    Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
    
'    If Dec < 0 And Dec > -1 Then
'        Me.txtDecD.Text = "-0"
'    Else
'        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
'    End If
'    Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
'    Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
End Sub

Private Sub cmdGetFromSky_Click()
    Dim RA As Double
    Dim Dec As Double
    
    Me.txtObjectName.Text = Planetarium.GetSelectedObject(RA, Dec, False)
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
    Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
    
'    If Dec < 0 And Dec > -1 Then
'        Me.txtDecD.Text = "-" & Fix(Dec + (0.5 / 3600))
'    Else
'        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
'    End If
'    Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
'    Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
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
'    Me.txtDecD.Text = GetMySetting(RegistryName, "DecD", "0")
'    Me.txtDecM.Text = GetMySetting(RegistryName, "DecM", "0")
'    Me.txtDecS.Text = GetMySetting(RegistryName, "DecS", "0")
    Me.txtObjectName.Text = GetMySetting(RegistryName, "ObjectName", "")
    
    Me.txtHA.Text = GetMySetting(RegistryName, "HA", "0.0000")
    
    Me.chkSoftSkip.Value = CInt(GetMySetting(RegistryName, "SoftSkip", "0"))
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtRAH_Validate(Cancel)
    If Not Cancel Then Call txtRAM_Validate(Cancel)
    If Not Cancel Then Call txtRAS_Validate(Cancel)
'    If Not Cancel Then Call txtDecD_Validate(Cancel)
'    If Not Cancel Then Call txtDecM_Validate(Cancel)
'    If Not Cancel Then Call txtDecS_Validate(Cancel)
    If Not Cancel Then Call txtHA_Validate(Cancel)
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "RAH", Me.txtRAH.Text)
    Call SaveMySetting(RegistryName, "RAM", Me.txtRAM.Text)
    Call SaveMySetting(RegistryName, "RAS", Me.txtRAS.Text)
'    Call SaveMySetting(RegistryName, "DecD", Me.txtDecD.Text)
'    Call SaveMySetting(RegistryName, "DecM", Me.txtDecM.Text)
'    Call SaveMySetting(RegistryName, "DecS", Me.txtDecS.Text)
    Call SaveMySetting(RegistryName, "ObjectName", Me.txtObjectName.Text)
    Call SaveMySetting(RegistryName, "HA", Me.txtHA.Text)
    Call SaveMySetting(RegistryName, "SoftSkip", Me.chkSoftSkip.Value)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub txtHA_GotFocus()
    Me.txtHA.SelStart = 0
    Me.txtHA.SelLength = Len(Me.txtHA.Text)
End Sub

Private Sub txtHA_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtHA.Text)
    If Err.Number <> 0 Or Test <= -12 Or Test >= 12 Then
        Beep
        Cancel = True
    Else
        Me.txtHA.Text = Format(Test, "0.0000")
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

'Private Sub txtDecD_GotFocus()
'    Me.txtDecD.SelStart = 0
'    Me.txtDecD.SelLength = Len(Me.txtDecD.Text)
'End Sub
'
'Private Sub txtDecD_Validate(Cancel As Boolean)
'    Dim Test As Integer
'    On Error Resume Next
'    Test = CInt(Me.txtDecD.Text)
'    If Err.Number <> 0 Or Test < -90 Or Test > 90 Or Test <> Me.txtDecD.Text Then
'        Beep
'        Cancel = True
'    End If
'    On Error GoTo 0
'End Sub
'
'Private Sub txtDecM_GotFocus()
'    Me.txtDecM.SelStart = 0
'    Me.txtDecM.SelLength = Len(Me.txtDecM.Text)
'End Sub
'
'Private Sub txtDecM_Validate(Cancel As Boolean)
'    Dim Test As Integer
'    On Error Resume Next
'    Test = CInt(Me.txtDecM.Text)
'    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtDecM.Text Then
'        Beep
'        Cancel = True
'    End If
'    On Error GoTo 0
'End Sub
'
'Private Sub txtDecS_GotFocus()
'    Me.txtDecS.SelStart = 0
'    Me.txtDecS.SelLength = Len(Me.txtDecS.Text)
'End Sub
'
'Private Sub txtDecS_Validate(Cancel As Boolean)
'    Dim Test As Integer
'    On Error Resume Next
'    Test = CInt(Me.txtDecS.Text)
'    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtDecS.Text Then
'        Beep
'        Cancel = True
'    End If
'    On Error GoTo 0
'End Sub

Public Sub PutFormDataIntoClass(clsAction As SkipAheadAtHAAction)
    clsAction.RA = CDbl(Me.txtRAH.Text) + (CDbl(Me.txtRAM.Text) / 60) + (CDbl(Me.txtRAS.Text) / 3600)
'    clsAction.Dec = Abs(CDbl(Me.txtDecD.Text)) + (CDbl(Me.txtDecM.Text) / 60) + (CDbl(Me.txtDecS.Text) / 3600)
'    If InStr(Me.txtDecD.Text, "-") <> 0 Then
'        clsAction.Dec = -clsAction.Dec
'    End If
    clsAction.Name = Me.txtObjectName.Text
    clsAction.HA = CDbl(Me.txtHA.Text)
    If Me.chkSoftSkip.Value = vbChecked Then
        clsAction.SoftSkip = True
    Else
        clsAction.SoftSkip = False
    End If
End Sub

Public Sub GetFormDataFromClass(clsAction As SkipAheadAtHAAction)
    Dim RA As Double
    Dim Dec As Double
    
    RA = clsAction.RA
'    Dec = clsAction.Dec
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    RA = (RA - Fix(RA + (0.5 / 3600))) * 60
    Me.txtRAM.Text = Fix(RA + (0.5 / 60))
    RA = (RA - Fix(RA + (0.5 / 60))) * 60
    Me.txtRAS.Text = Fix(RA + 0.5)

'    If Dec < 0 And Dec > -1 Then
'        Dec = -Dec
'        Me.txtDecD.Text = "-0"
'    ElseIf Dec < 0 Then
'        Dec = -Dec
'        Me.txtDecD.Text = -Fix(Dec)
'    Else
'        Me.txtDecD.Text = Fix(Dec)
'    End If
'    Dec = (Dec - Fix(Dec)) * 60
'    Me.txtDecM.Text = Fix(Dec)
'    Dec = (Dec - Fix(Dec)) * 60 + 0.5
'    Me.txtDecS.Text = Fix(Dec)
    
    Me.txtObjectName.Text = clsAction.Name
    
    Me.txtHA.Text = Format(clsAction.HA, "0.0000")
    
    If clsAction.SoftSkip Then
        Me.chkSoftSkip.Value = vbChecked
    Else
        Me.chkSoftSkip.Value = vbUnchecked
    End If
    
    Call SaveMySetting(RegistryName, "RAH", Me.txtRAH.Text)
    Call SaveMySetting(RegistryName, "RAM", Me.txtRAM.Text)
    Call SaveMySetting(RegistryName, "RAS", Me.txtRAS.Text)
'    Call SaveMySetting(RegistryName, "DecD", Me.txtDecD.Text)
'    Call SaveMySetting(RegistryName, "DecM", Me.txtDecM.Text)
'    Call SaveMySetting(RegistryName, "DecS", Me.txtDecS.Text)
    Call SaveMySetting(RegistryName, "ObjectName", Me.txtObjectName.Text)
    Call SaveMySetting(RegistryName, "HA", Me.txtHA.Text)
    Call SaveMySetting(RegistryName, "SoftSkip", Me.chkSoftSkip.Value)
    
End Sub

