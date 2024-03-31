VERSION 5.00
Begin VB.Form frmRotator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instrument Rotator Action"
   ClientHeight    =   1230
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5190
   HelpContextID   =   1700
   Icon            =   "frmRotator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRotationAngle 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "320.45"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdGetFromSky 
      Caption         =   "Get Position Angle from TheSky6 FOVI"
      Height          =   615
      Left            =   900
      TabIndex        =   1
      Top             =   540
      Width           =   2115
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "degrees from North"
      Height          =   195
      Left            =   2340
      TabIndex        =   5
      Top             =   180
      Width           =   1350
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "New Rotation Angle:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   180
      Width           =   1470
   End
End
Attribute VB_Name = "frmRotator"
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

Private Const RegistryName = "RotatorAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub cmdGetFromSky_Click()
    Me.txtRotationAngle = Format(Planetarium.GetFOVIPositionAngle, "0.00")
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    If frmOptions.lstPlanetarium.ListIndex = 0 Then
        Me.cmdGetFromSky.Enabled = False
    Else
        Me.cmdGetFromSky.Enabled = True
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
    Me.txtRotationAngle.Text = GetMySetting(RegistryName, "RotationAngle", "0")
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtRotationAngle_Validate(Cancel)
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "RotationAngle", Me.txtRotationAngle.Text)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub txtRotationAngle_GotFocus()
    Me.txtRotationAngle.SelStart = 0
    Me.txtRotationAngle.SelLength = Len(Me.txtRotationAngle.Text)
End Sub

Private Sub txtRotationAngle_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtRotationAngle.Text)
    If Err.Number <> 0 Or Test < 0 Or Test > 360 Or Test <> Me.txtRotationAngle.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Public Sub PutFormDataIntoClass(clsAction As RotatorAction)
    clsAction.RotationAngle = CDbl(Me.txtRotationAngle.Text)
End Sub

Public Sub GetFormDataFromClass(clsAction As RotatorAction)
    Me.txtRotationAngle.Text = clsAction.RotationAngle
    Call SaveMySetting(RegistryName, "RotationAngle", Me.txtRotationAngle.Text)
End Sub

