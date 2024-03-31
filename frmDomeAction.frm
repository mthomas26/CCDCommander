VERSION 5.00
Begin VB.Form frmDomeAction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dome Control Action"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3735
   HelpContextID   =   1600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAzimuth 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Text            =   "359"
      Top             =   1560
      Width           =   435
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Slew to Azimuth:"
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Park"
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1395
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Home"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1395
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Uncouple from mount"
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1995
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Couple to mount"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Open Slit"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1395
   End
   Begin VB.OptionButton optDomeAction 
      Caption         =   "Close Slit"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "degrees"
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   1620
      Width           =   570
   End
End
Attribute VB_Name = "frmDomeAction"
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

Private Const RegistryName = "DomeAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
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
    Me.optDomeAction(CInt(GetMySetting(RegistryName, "DomeAction", "0"))).Value = True
    
    Me.txtAzimuth.Text = GetMySetting(RegistryName, "DomeAzimuth", "0")
End Sub

Private Sub OKButton_Click()
    If Me.optDomeAction(0).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "0")
    ElseIf Me.optDomeAction(1).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "1")
    ElseIf Me.optDomeAction(2).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "2")
    ElseIf Me.optDomeAction(3).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "3")
    ElseIf Me.optDomeAction(4).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "4")
    ElseIf Me.optDomeAction(5).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "5")
    ElseIf Me.optDomeAction(6).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "6")
    End If
    
    Call SaveMySetting(RegistryName, "DomeAzimuth", Me.txtAzimuth.Text)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As DomeAction)
    If Me.optDomeAction(0).Value Then
        clsAction.ThisDomeAction = actCloseDome
    ElseIf Me.optDomeAction(1).Value Then
        clsAction.ThisDomeAction = actOpenDome
    ElseIf Me.optDomeAction(2).Value Then
        clsAction.ThisDomeAction = actCoupleDome
    ElseIf Me.optDomeAction(3).Value Then
        clsAction.ThisDomeAction = actUnCoupleDome
    ElseIf Me.optDomeAction(4).Value Then
        clsAction.ThisDomeAction = actHomeDome
    ElseIf Me.optDomeAction(5).Value Then
        clsAction.ThisDomeAction = actParkDome
    ElseIf Me.optDomeAction(6).Value Then
        clsAction.ThisDomeAction = actSlewDomeAzimuth
    End If
    
    clsAction.Azimuth = CInt(Me.txtAzimuth.Text)
End Sub

Public Sub GetFormDataFromClass(clsAction As DomeAction)
    Me.optDomeAction(clsAction.ThisDomeAction).Value = True

    If Me.optDomeAction(0).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "0")
    ElseIf Me.optDomeAction(1).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "1")
    ElseIf Me.optDomeAction(2).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "2")
    ElseIf Me.optDomeAction(3).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "3")
    ElseIf Me.optDomeAction(4).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "4")
    ElseIf Me.optDomeAction(5).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "5")
    ElseIf Me.optDomeAction(6).Value Then
        Call SaveMySetting(RegistryName, "DomeAction", "6")
    End If
    
    Call SaveMySetting(RegistryName, "DomeAzimuth", clsAction.Azimuth)
    
End Sub



Private Sub optDomeAction_Click(Index As Integer)
    If (Index = 6) Then
        Me.txtAzimuth.Enabled = True
    Else
        Me.txtAzimuth.Enabled = False
    End If
End Sub

Private Sub txtAzimuth_GotFocus()
    Me.txtAzimuth.SelStart = 0
    Me.txtAzimuth.SelLength = Len(Me.txtAzimuth.Text)
End Sub

Private Sub txtAzimuth_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimuth.Text)
    If Err.Number <> 0 Or Test < 0 Or Test > 359 Or Test <> Me.txtAzimuth.Text Then
        Beep
        Cancel = True
    Else
        Me.txtAzimuth.Text = Format(Me.txtAzimuth.Text, "0")
    End If
    On Error GoTo 0
End Sub


