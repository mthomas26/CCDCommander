VERSION 5.00
Begin VB.Form frmSkipAtTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Skip Ahead at Time Action"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3270
   HelpContextID   =   1800
   Icon            =   "frmSkipAtTime.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSoftSkip 
      Caption         =   "Soft Skip"
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   1300
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time to Skip Ahead"
      Height          =   1155
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   1692
      Begin VB.TextBox txtMinute 
         Height          =   252
         Left            =   900
         TabIndex        =   1
         Text            =   "32"
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtHour 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   420
         TabIndex        =   0
         Text            =   "22"
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(24 hour format)"
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   780
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   780
         TabIndex        =   5
         Top             =   360
         Width           =   96
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSkipAtTime"
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

Private Const RegistryName = "SkipAheadAtTime"

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
    Me.txtHour.Text = GetMySetting(RegistryName, "Hour", "00")
    Me.txtMinute.Text = GetMySetting(RegistryName, "Minute", "00")
    Me.chkSoftSkip.Value = CInt(GetMySetting(RegistryName, "SoftSkip", "0"))
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtHour_Validate(Cancel)
    If Not Cancel Then Call txtMinute_Validate(Cancel)
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "Hour", Me.txtHour.Text)
    Call SaveMySetting(RegistryName, "Minute", Me.txtMinute.Text)
    Call SaveMySetting(RegistryName, "SoftSkip", Me.chkSoftSkip.Value)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As SkipAheadAtTimeAction)
    clsAction.Hour = Me.txtHour.Text
    clsAction.Minute = Me.txtMinute.Text
    If Me.chkSoftSkip.Value = vbChecked Then
        clsAction.SoftSkip = True
    Else
        clsAction.SoftSkip = False
    End If
End Sub

Public Sub GetFormDataFromClass(clsAction As SkipAheadAtTimeAction)
    If clsAction.Hour < 10 Then
        Me.txtHour.Text = "0" & clsAction.Hour
    Else
        Me.txtHour.Text = clsAction.Hour
    End If
    
    If clsAction.Minute < 10 Then
        Me.txtMinute.Text = "0" & clsAction.Minute
    Else
        Me.txtMinute.Text = clsAction.Minute
    End If
    
    If clsAction.SoftSkip Then
        Me.chkSoftSkip.Value = vbChecked
    Else
        Me.chkSoftSkip.Value = vbUnchecked
    End If
    
    Call SaveMySetting(RegistryName, "Hour", Me.txtHour.Text)
    Call SaveMySetting(RegistryName, "Minute", Me.txtMinute.Text)
    Call SaveMySetting(RegistryName, "SoftSkip", Me.chkSoftSkip.Value)
End Sub

Private Sub txtHour_GotFocus()
    Me.txtHour.SelStart = 0
    Me.txtHour.SelLength = Len(Me.txtHour.Text)
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtHour.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 24 Or Test <> Me.txtHour.Text Then
        Beep
        Cancel = True
    Else
        If CInt(Me.txtHour.Text) < 10 Then
            Me.txtHour.Text = "0" & CInt(Me.txtHour.Text)
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub txtMinute_GotFocus()
    Me.txtMinute.SelStart = 0
    Me.txtMinute.SelLength = Len(Me.txtMinute.Text)
End Sub

Private Sub txtMinute_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtMinute.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtMinute.Text Then
        Beep
        Cancel = True
    Else
        If CInt(Me.txtMinute.Text) < 10 Then
            Me.txtMinute.Text = "0" & CInt(Me.txtMinute.Text)
        End If
    End If
    On Error GoTo 0
End Sub
