VERSION 5.00
Begin VB.Form frmCloudMonitorAction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weather Monitor Action"
   ClientHeight    =   1260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3270
   HelpContextID   =   2000
   Icon            =   "frmCloudMonitorAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optStatus 
      Caption         =   "Disabled"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1515
   End
   Begin VB.OptionButton optStatus 
      Caption         =   "Enabled"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCloudMonitorAction"
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

Private Const RegistryName = "CloudMonitorAction"

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
    If CBool(GetMySetting(RegistryName, "Status", True)) = True Then
        Me.optStatus(0).Value = True
    Else
        Me.optStatus(1).Value = True
    End If
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    If Me.optStatus(0).Value = True Then
        Call SaveMySetting(RegistryName, "Status", True)
    Else
        Call SaveMySetting(RegistryName, "Status", False)
    End If
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As CloudMonitorAction)
    If Me.optStatus(0).Value = True Then
        clsAction.Enabled = True
    Else
        clsAction.Enabled = False
    End If
End Sub

Public Sub GetFormDataFromClass(clsAction As CloudMonitorAction)
    If clsAction.Enabled = True Then
        Me.optStatus(0).Value = True
        Call SaveMySetting(RegistryName, "Status", True)
    Else
        Me.optStatus(1).Value = True
        Call SaveMySetting(RegistryName, "Status", False)
    End If
End Sub
