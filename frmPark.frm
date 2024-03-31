VERSION 5.00
Begin VB.Form frmPark 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Park Action"
   ClientHeight    =   3045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4620
   HelpContextID   =   900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optParkMethod 
      Caption         =   "Tracking Off"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   21
      Top             =   720
      Width           =   2115
   End
   Begin VB.OptionButton optParkMethod 
      Caption         =   "Home && Tracking Off"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   20
      Top             =   420
      Width           =   2115
   End
   Begin VB.CheckBox chkParkRotator 
      Caption         =   "Park Rotator"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2700
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.OptionButton optParkMethod 
      Caption         =   "Simulated Park"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   2115
   End
   Begin VB.OptionButton optParkMethod 
      Caption         =   "Real Park"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2115
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame fraSimParkCoordinates 
      Caption         =   "Simulated Park Coordinates"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   120
      TabIndex        =   4
      Top             =   1380
      Width           =   2955
      Begin VB.TextBox txtAltD 
         Height          =   312
         Left            =   660
         TabIndex        =   10
         Text            =   "23"
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox txtAltM 
         Height          =   312
         Left            =   1380
         TabIndex        =   9
         Text            =   "10"
         Top             =   360
         Width           =   312
      End
      Begin VB.TextBox txtAltS 
         Height          =   312
         Left            =   2040
         TabIndex        =   8
         Text            =   "58"
         Top             =   360
         Width           =   312
      End
      Begin VB.TextBox txtAzimD 
         Height          =   312
         Left            =   660
         TabIndex        =   7
         Text            =   "5"
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox txtAzimM 
         Height          =   312
         Left            =   1380
         TabIndex        =   6
         Text            =   "23"
         Top             =   720
         Width           =   312
      End
      Begin VB.TextBox txtAzimS 
         Height          =   312
         Left            =   2040
         TabIndex        =   5
         Text            =   "15"
         Top             =   720
         Width           =   312
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Alt:"
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "d"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Left            =   1740
         TabIndex        =   16
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   420
         Width           =   90
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "d"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   780
         Width           =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Left            =   1740
         TabIndex        =   13
         Top             =   780
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   780
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Azim:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   780
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmPark"
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

Private Const RegistryName = "ParkAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    If frmOptions.lstRotator.ListIndex > 0 Then
        Me.chkParkRotator.Enabled = True
    Else
        Me.chkParkRotator.Enabled = False
    End If
    
    If True Then 'frmOptions.lstMountControl.ListIndex = 1 Then
        Me.optParkMethod(2).Enabled = True
        Me.optParkMethod(3).Enabled = True
    Else
        Me.optParkMethod(2).Enabled = False
        Me.optParkMethod(3).Enabled = False
    End If
    
    Call GetSettings
End Sub

Private Sub GetSettings()
    Dim TempList As String
    Dim NewTemp As Double
    Dim NewTempStr As String
    
    If CBool(GetMySetting(RegistryName, "SimulatedPark", "True")) Then
        Me.optParkMethod(0).Value = False
        Me.optParkMethod(1).Value = True
    Else
        Me.optParkMethod(0).Value = False
        Me.optParkMethod(1).Value = False
        Me.optParkMethod(2).Value = False
        Me.optParkMethod(3).Value = False
        Me.optParkMethod(CInt(GetMySetting(RegistryName, "ParkType", "0"))).Value = True
    End If
    
    
    Me.txtAltD.Text = GetMySetting(RegistryName, "AltD", "0")
    Me.txtAltM.Text = GetMySetting(RegistryName, "AltM", "0")
    Me.txtAltS.Text = GetMySetting(RegistryName, "AltS", "0")
    Me.txtAzimD.Text = GetMySetting(RegistryName, "AzimD", "0")
    Me.txtAzimM.Text = GetMySetting(RegistryName, "AzimM", "0")
    Me.txtAzimS.Text = GetMySetting(RegistryName, "AzimS", "0")
    
    Me.chkParkRotator.Value = GetMySetting(RegistryName, "ParkRotator", "1")
End Sub

Private Sub OKButton_Click()
    Dim Counter As Integer
    Dim TempList As String
    Dim Cancel As Boolean
    
    Cancel = False
    If Not Cancel Then Call txtAltD_Validate(Cancel)
    If Not Cancel Then Call txtAltM_Validate(Cancel)
    If Not Cancel Then Call txtAltS_Validate(Cancel)
    If Not Cancel Then Call txtAzimD_Validate(Cancel)
    If Not Cancel Then Call txtAzimM_Validate(Cancel)
    If Not Cancel Then Call txtAzimS_Validate(Cancel)
    
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "SimulatedPark", "False")
    
    If Me.optParkMethod(0).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 0)
    ElseIf Me.optParkMethod(1).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 1)
    ElseIf Me.optParkMethod(2).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 2)
    ElseIf Me.optParkMethod(3).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 3)
    End If
    
    Call SaveMySetting(RegistryName, "AltD", Me.txtAltD.Text)
    Call SaveMySetting(RegistryName, "AltM", Me.txtAltM.Text)
    Call SaveMySetting(RegistryName, "AltS", Me.txtAltS.Text)
    Call SaveMySetting(RegistryName, "AzimD", Me.txtAzimD.Text)
    Call SaveMySetting(RegistryName, "AzimM", Me.txtAzimM.Text)
    Call SaveMySetting(RegistryName, "AzimS", Me.txtAzimS.Text)
    
    Call SaveMySetting(RegistryName, "ParkRotator", Me.chkParkRotator.Value)

    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As ParkMountAction)
    
    If Me.optParkMethod(0).Value Then
        clsAction.DoSimulatedPark = False
        clsAction.DoHomePark = False
        clsAction.DoTrackingOff = False
    ElseIf Me.optParkMethod(1).Value Then
        clsAction.DoSimulatedPark = True
        clsAction.DoHomePark = False
        clsAction.DoTrackingOff = False
    ElseIf Me.optParkMethod(2).Value Then
        clsAction.DoSimulatedPark = False
        clsAction.DoHomePark = True
        clsAction.DoTrackingOff = False
    ElseIf Me.optParkMethod(3).Value Then
        clsAction.DoSimulatedPark = False
        clsAction.DoHomePark = False
        clsAction.DoTrackingOff = True
    End If
    
    clsAction.AltD = CDbl(Me.txtAltD.Text)
    clsAction.AltM = CDbl(Me.txtAltM.Text)
    clsAction.AltS = CDbl(Me.txtAltS.Text)
    clsAction.AzimD = CDbl(Me.txtAzimD.Text)
    clsAction.AzimM = CDbl(Me.txtAzimM.Text)
    clsAction.AzimS = CDbl(Me.txtAzimS.Text)
    
    If (Me.chkParkRotator.Value = vbChecked) Then
        clsAction.ParkRotator = True
    Else
        clsAction.ParkRotator = False
    End If
End Sub

Public Sub GetFormDataFromClass(clsAction As ParkMountAction)
    Dim Counter As Long
    
    If clsAction.DoSimulatedPark Then
        Me.optParkMethod(0).Value = False
        Me.optParkMethod(1).Value = True
        Me.optParkMethod(2).Value = False
        Me.optParkMethod(3).Value = False
    ElseIf clsAction.DoHomePark Then
        Me.optParkMethod(0).Value = False
        Me.optParkMethod(1).Value = False
        Me.optParkMethod(2).Value = True
        Me.optParkMethod(3).Value = False
    ElseIf clsAction.DoTrackingOff Then
        Me.optParkMethod(0).Value = False
        Me.optParkMethod(1).Value = False
        Me.optParkMethod(2).Value = False
        Me.optParkMethod(3).Value = True
    Else
        Me.optParkMethod(0).Value = True
        Me.optParkMethod(1).Value = False
        Me.optParkMethod(2).Value = False
        Me.optParkMethod(3).Value = False
    End If
    
    Me.txtAltD.Text = clsAction.AltD
    Me.txtAltM.Text = clsAction.AltM
    Me.txtAltS.Text = clsAction.AltS
    Me.txtAzimD.Text = clsAction.AzimD
    Me.txtAzimM.Text = clsAction.AzimM
    Me.txtAzimS.Text = clsAction.AzimS
    
    If clsAction.ParkRotator Then
        Me.chkParkRotator.Value = vbChecked
    Else
        Me.chkParkRotator.Value = vbUnchecked
    End If

    If Me.optParkMethod(0).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 0)
    ElseIf Me.optParkMethod(1).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 1)
    ElseIf Me.optParkMethod(2).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 2)
    ElseIf Me.optParkMethod(3).Value Then
        Call SaveMySetting(RegistryName, "ParkType", 3)
    End If
    
    Call SaveMySetting(RegistryName, "AltD", Me.txtAltD.Text)
    Call SaveMySetting(RegistryName, "AltM", Me.txtAltM.Text)
    Call SaveMySetting(RegistryName, "AltS", Me.txtAltS.Text)
    Call SaveMySetting(RegistryName, "AzimD", Me.txtAzimD.Text)
    Call SaveMySetting(RegistryName, "AzimM", Me.txtAzimM.Text)
    Call SaveMySetting(RegistryName, "AzimS", Me.txtAzimS.Text)
    
    Call SaveMySetting(RegistryName, "ParkRotator", Me.chkParkRotator.Value)
End Sub

Private Sub optParkMethod_Click(Index As Integer)
    If Me.optParkMethod(1).Value Then
        Me.fraSimParkCoordinates.Enabled = True
    Else
        Me.fraSimParkCoordinates.Enabled = False
    End If
End Sub

Private Sub txtAltD_GotFocus()
    Me.txtAltD.SelStart = 0
    Me.txtAltD.SelLength = Len(Me.txtAltD.Text)
End Sub

Private Sub txtAltD_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAltD.Text)
    If Err.Number <> 0 Or Test < -90 Or Test >= 90 Or Test <> Me.txtAltD.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAltM_GotFocus()
    Me.txtAltM.SelStart = 0
    Me.txtAltM.SelLength = Len(Me.txtAltM.Text)
End Sub

Private Sub txtAltM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAltM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAltM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAltS_GotFocus()
    Me.txtAltS.SelStart = 0
    Me.txtAltS.SelLength = Len(Me.txtAltS.Text)
End Sub

Private Sub txtAltS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAltS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAltS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAzimD_GotFocus()
    Me.txtAzimD.SelStart = 0
    Me.txtAzimD.SelLength = Len(Me.txtAzimD.Text)
End Sub

Private Sub txtAzimD_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimD.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 360 Or Test <> Me.txtAzimD.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAzimM_GotFocus()
    Me.txtAzimM.SelStart = 0
    Me.txtAzimM.SelLength = Len(Me.txtAzimM.Text)
End Sub

Private Sub txtAzimM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAzimM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAzimS_GotFocus()
    Me.txtAzimS.SelStart = 0
    Me.txtAzimS.SelLength = Len(Me.txtAzimS.Text)
End Sub

Private Sub txtAzimS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAzimS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub
