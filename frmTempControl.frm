VERSION 5.00
Begin VB.Form frmTempControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intelligent Temperature Control Action"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   HelpContextID   =   1900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRampWarmUp 
      Caption         =   "Ramp Warm-up"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   420
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkUseIntelligentCooling 
      Caption         =   "Use Intelligent Cooling"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.CheckBox chkFanOn 
      Caption         =   "Fan On"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.OptionButton optEnaDisCooler 
      Caption         =   "Disable Cooler"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   2115
   End
   Begin VB.OptionButton optEnaDisCooler 
      Caption         =   "Enable Cooler"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraSimpleCooling 
      Caption         =   "Simple Cooling"
      Height          =   915
      Left            =   60
      TabIndex        =   25
      Top             =   1140
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtSimpleTemp 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Text            =   "-20.0"
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Temperature Set Point:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "*C"
         Height          =   195
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   165
      End
   End
   Begin VB.Frame fraCooling 
      Caption         =   "Intelligent Cooling"
      Height          =   2415
      Left            =   60
      TabIndex        =   14
      Top             =   1140
      Width           =   5835
      Begin VB.TextBox txtCoolerDeviation 
         Height          =   315
         Left            =   4875
         TabIndex        =   11
         Text            =   "4"
         Top             =   1980
         Width           =   555
      End
      Begin VB.TextBox txtDeviation 
         Height          =   315
         Left            =   4860
         TabIndex        =   10
         Text            =   "0.5"
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtMaxTime 
         Height          =   315
         Left            =   4860
         TabIndex        =   9
         Text            =   "5"
         Top             =   900
         Width           =   555
      End
      Begin VB.TextBox txtMaxPower 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Text            =   "80"
         Top             =   420
         Width           =   555
      End
      Begin VB.ListBox lstTemps 
         Height          =   1815
         ItemData        =   "frmTempControl.frx":0000
         Left            =   1200
         List            =   "frmTempControl.frx":0002
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdAddTemp 
         Caption         =   "Add Temperature"
         Height          =   555
         Left            =   60
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteTemp 
         Caption         =   "Delete Temperature"
         Height          =   555
         Left            =   60
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "cooler deviation is <"
         Height          =   195
         Left            =   3405
         TabIndex        =   24
         Top             =   2025
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   5475
         TabIndex        =   23
         Top             =   2040
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Temperature is stable when:"
         Height          =   195
         Left            =   2400
         TabIndex        =   22
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "*C"
         Height          =   195
         Left            =   5460
         TabIndex        =   21
         Top             =   1680
         Width           =   165
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "temperature deviation is <"
         Height          =   195
         Left            =   2985
         TabIndex        =   20
         Top             =   1665
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "min"
         Height          =   195
         Left            =   5460
         TabIndex        =   19
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   5460
         TabIndex        =   18
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum time for temperature to stabilize:"
         Height          =   375
         Left            =   2820
         TabIndex        =   17
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum cooler power:"
         Height          =   195
         Left            =   3150
         TabIndex        =   16
         Top             =   480
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Desired Temperatures:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmTempControl"
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

Private Const RegistryName = "IntelligentTempAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub chkUseIntelligentCooling_Click()
    If Me.chkUseIntelligentCooling.Value = vbChecked Then
        Me.fraCooling.Visible = True
        Me.fraSimpleCooling.Visible = False
    Else
        Me.fraCooling.Visible = False
        Me.fraSimpleCooling.Visible = True
    End If
End Sub

Private Sub cmdAddTemp_Click()
    Dim NewTemp As String
    Dim NewTempVal As Double
    
    NewTemp = InputBox("Enter the temperature to add.", "Adding a temperature...")

    If NewTemp <> "" Then
        On Error Resume Next
        NewTempVal = CDbl(NewTemp)
        If Err.Number <> 0 Or NewTempVal <> NewTemp Then
            Beep
            Exit Sub
        End If
        On Error GoTo 0
        
        Me.lstTemps.AddItem (Format(NewTempVal, "0.0"))
        
        Call SortTemps
    End If
End Sub

Private Sub cmdDeleteTemp_Click()
    If Me.lstTemps.ListIndex = -1 Then
        MsgBox "You must select a temperature before you can delete it."
        Exit Sub
    End If
    
    Call Me.lstTemps.RemoveItem(Me.lstTemps.ListIndex)
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    If frmOptions.lstCameraControl.ListIndex = CameraControl.CCDSoft Or frmOptions.lstCameraControl.ListIndex = CameraControl.CCDSoftAO Then
        Me.chkFanOn.Enabled = True
    ElseIf frmOptions.lstCameraControl.ListIndex = CameraControl.MaxIm Then
        Me.chkFanOn.Enabled = True
    End If
    
    Call GetSettings
End Sub

Private Sub GetSettings()
    Dim TempList As String
    Dim NewTemp As Double
    Dim NewTempStr As String
    
    If CBool(GetMySetting(RegistryName, "CoolerOn", "True")) Then
        Me.optEnaDisCooler(0).Value = True
        Me.optEnaDisCooler(1).Value = False
    Else
        Me.optEnaDisCooler(1).Value = True
        Me.optEnaDisCooler(0).Value = False
    End If
    
    TempList = GetMySetting(RegistryName, "TempList", "-10,-15,-20,-25,-30,")
    Me.lstTemps.Clear
    Do While Len(TempList) > 0
        NewTemp = CDbl(Left(TempList, InStr(TempList, ",") - 1))
        NewTempStr = Format(NewTemp, "0.0")
        
        Me.lstTemps.AddItem NewTempStr
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop

    Me.txtMaxPower.Text = CDbl(GetMySetting(RegistryName, "MaxPower", "80"))
    Me.txtMaxTime.Text = CDbl(GetMySetting(RegistryName, "MaxTime", "5"))

    Me.txtDeviation.Text = CDbl(GetMySetting(RegistryName, "Deviation", Format(1 / 2, "0.0")))
    Me.txtCoolerDeviation.Text = CDbl(GetMySetting(RegistryName, "CoolerDeviation", "4"))
    
    Me.chkFanOn.Value = CInt(GetMySetting(RegistryName, "FanOn", "1"))
    
    Me.chkUseIntelligentCooling.Value = CInt(GetMySetting(RegistryName, "IntelligentAction", "1"))
    
    Me.chkRampWarmUp.Value = CInt(GetMySetting(RegistryName, "RampWarmUp", "1"))
    
    If Me.chkUseIntelligentCooling.Value = 0 Then
        If Me.lstTemps.ListCount > 0 Then
            Me.txtSimpleTemp.Text = Me.lstTemps.List(0)
        Else
            Me.txtSimpleTemp.Text = Format(-20, "0.0")
        End If
    Else
        Me.txtSimpleTemp.Text = Format(-20, "0.0")
    End If
End Sub

Private Sub SortTemps()
    Dim TempsList() As Double
    Dim Counter As Integer
    Dim Counter2 As Integer
    Dim NumItems As Integer
    Dim Temporary As Double
    
    NumItems = Me.lstTemps.ListCount - 1
    ReDim TempsList(NumItems)
    
    For Counter = 0 To NumItems
        TempsList(Counter) = Me.lstTemps.List(Counter)
    Next Counter
    
    For Counter = NumItems To 0 Step -1
        For Counter2 = 1 To Counter
            If TempsList(Counter2 - 1) < TempsList(Counter2) Then
                Temporary = TempsList(Counter2 - 1)
                TempsList(Counter2 - 1) = TempsList(Counter2)
                TempsList(Counter2) = Temporary
            End If
        Next Counter2
    Next Counter
            
    Me.lstTemps.Clear
    For Counter = 0 To NumItems
        Me.lstTemps.AddItem Format(TempsList(Counter), "0.0")
    Next Counter
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    
    Cancel = False
    If Not Cancel Then Call txtMaxPower_Validate(Cancel)
    If Not Cancel Then Call txtMaxTime_Validate(Cancel)
    If Not Cancel Then Call txtDeviation_Validate(Cancel)
    If Not Cancel Then Call txtCoolerDeviation_Validate(Cancel)
    If Not Cancel Then Call txtSimpleTemp_Validate(Cancel)
    
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveSettings
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub SaveSettings()
    Dim Counter As Integer
    Dim TempList As String
    
    If Me.optEnaDisCooler(0).Value Then
        Call SaveMySetting(RegistryName, "CoolerOn", "True")
    Else
        Call SaveMySetting(RegistryName, "CoolerOn", "False")
    End If
    
    If Me.chkUseIntelligentCooling.Value = 1 Then
        TempList = ""
        For Counter = 0 To Me.lstTemps.ListCount - 1
            TempList = TempList & Me.lstTemps.List(Counter) & ","
        Next Counter
    Else
        TempList = Me.txtSimpleTemp.Text & ","
    End If
    
    Call SaveMySetting(RegistryName, "TempList", TempList)
    
    Call SaveMySetting(RegistryName, "MaxPower", Me.txtMaxPower.Text)
    Call SaveMySetting(RegistryName, "MaxTime", Me.txtMaxTime.Text)
    Call SaveMySetting(RegistryName, "Deviation", Me.txtDeviation.Text)
    Call SaveMySetting(RegistryName, "CoolerDeviation", Me.txtCoolerDeviation.Text)
    
    Call SaveMySetting(RegistryName, "FanOn", Me.chkFanOn.Value)
    
    Call SaveMySetting(RegistryName, "IntelligentAction", Me.chkUseIntelligentCooling.Value)
    
    Call SaveMySetting(RegistryName, "RampWarmUp", Me.chkRampWarmUp.Value)

End Sub

Private Sub optEnaDisCooler_Click(Index As Integer)
    If Me.optEnaDisCooler(0).Value = True Then
        Me.fraCooling.Enabled = True
        Me.chkUseIntelligentCooling.Enabled = True
        Me.fraSimpleCooling.Enabled = True
        Me.chkRampWarmUp.Enabled = False
    Else
        Me.fraCooling.Enabled = False
        Me.chkUseIntelligentCooling.Enabled = False
        Me.fraSimpleCooling.Enabled = False
        Me.chkRampWarmUp.Enabled = True
    End If
End Sub

Private Sub txtMaxPower_GotFocus()
    Me.txtMaxPower.SelStart = 0
    Me.txtMaxPower.SelLength = Len(Me.txtMaxPower.Text)
End Sub

Private Sub txtMaxPower_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMaxPower.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxPower.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxTime_GotFocus()
    Me.txtMaxTime.SelStart = 0
    Me.txtMaxTime.SelLength = Len(Me.txtMaxTime.Text)
End Sub

Private Sub txtMaxTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMaxTime.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxTime.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtCoolerDeviation_GotFocus()
    Me.txtCoolerDeviation.SelStart = 0
    Me.txtCoolerDeviation.SelLength = Len(Me.txtCoolerDeviation.Text)
End Sub

Private Sub txtCoolerDeviation_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtCoolerDeviation.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtCoolerDeviation.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDeviation_GotFocus()
    Me.txtDeviation.SelStart = 0
    Me.txtDeviation.SelLength = Len(Me.txtDeviation.Text)
End Sub

Private Sub txtDeviation_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDeviation.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtDeviation.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Public Sub PutFormDataIntoClass(clsAction As IntelligentTempAction)
    Dim Counter As Long
    
    If Me.optEnaDisCooler(0).Value Then
        clsAction.CoolerOn = True
    Else
        clsAction.CoolerOn = False
    End If
    
    clsAction.MaxCoolerPower = CDbl(Me.txtMaxPower.Text)
    clsAction.MaxTime = CDbl(Me.txtMaxTime.Text)
    clsAction.Deviation = CDbl(Me.txtDeviation.Text)
    clsAction.CoolerDeviation = CDbl(Me.txtCoolerDeviation.Text)
    
    clsAction.FanOn = Me.chkFanOn.Value
    
    clsAction.IntelligentAction = Me.chkUseIntelligentCooling.Value
    
    If Me.chkUseIntelligentCooling.Value = 1 Then
        clsAction.NumTemps = Me.lstTemps.ListCount - 1
        
        For Counter = 0 To clsAction.NumTemps
            clsAction.DesiredTemperatures(Counter) = CDbl(Me.lstTemps.List(Counter))
        Next Counter
    Else
        clsAction.NumTemps = 1
        clsAction.DesiredTemperatures(0) = CDbl(Me.txtSimpleTemp.Text)
    End If
    
    clsAction.RampWarmUp = Me.chkRampWarmUp.Value
End Sub

Public Sub GetFormDataFromClass(clsAction As IntelligentTempAction)
    Dim Counter As Long
    
    If clsAction.CoolerOn Then
        Me.optEnaDisCooler(0).Value = True
        Me.optEnaDisCooler(1).Value = False
    Else
        Me.optEnaDisCooler(0).Value = False
        Me.optEnaDisCooler(1).Value = True
    End If
    
    Me.txtMaxPower.Text = clsAction.MaxCoolerPower
    Me.txtMaxTime.Text = clsAction.MaxTime
    Me.txtDeviation.Text = Format(clsAction.Deviation, "0.0")
    Me.txtCoolerDeviation.Text = Format(clsAction.CoolerDeviation, "0")
    
    Me.chkFanOn.Value = clsAction.FanOn
    
    Me.chkUseIntelligentCooling.Value = clsAction.IntelligentAction
    
    Me.lstTemps.Clear
    For Counter = 0 To clsAction.NumTemps
        Me.lstTemps.AddItem Format(clsAction.DesiredTemperatures(Counter), "0.0")
    Next Counter
    
    Me.txtSimpleTemp.Text = Format(clsAction.DesiredTemperatures(0), "0.0")

    Me.chkRampWarmUp.Value = clsAction.RampWarmUp

    Call SaveSettings
End Sub

Private Sub txtSimpleTemp_GotFocus()
    Me.txtSimpleTemp.SelStart = 0
    Me.txtSimpleTemp.SelLength = Len(Me.txtSimpleTemp.Text)
End Sub

Private Sub txtSimpleTemp_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtSimpleTemp.Text)
    If Err.Number <> 0 Or Test <> Me.txtSimpleTemp.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub
