VERSION 5.00
Begin VB.Form frmWaitForTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wait for Time Action"
   ClientHeight    =   1995
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3750
   HelpContextID   =   1200
   Icon            =   "frmWaitForTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optTimeOption 
      Caption         =   "Relative"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   11
      Top             =   420
      Width           =   1575
   End
   Begin VB.OptionButton optTimeOption 
      Caption         =   "Absolute"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   10
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1700
      Begin VB.TextBox txtSecond 
         Height          =   252
         Left            =   1440
         TabIndex        =   2
         Text            =   "15"
         Top             =   420
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.VScrollBar vsHour 
         Height          =   252
         Left            =   60
         Max             =   23
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.VScrollBar vsMinute 
         Height          =   252
         Left            =   1920
         Max             =   59
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.TextBox txtMinute 
         Height          =   252
         Left            =   900
         TabIndex        =   1
         Text            =   "15"
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
      Begin VB.Label lblSecSeparator 
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
         Left            =   1300
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lbl24hformat 
         Alignment       =   2  'Center
         Caption         =   "(24 hour format)"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   780
         Width           =   1590
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
         TabIndex        =   6
         Top             =   360
         Width           =   96
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2460
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmWaitForTime"
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

Private Const RegistryName = "WaitForTime"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    Call GetSettings

    Me.ScaleMode = vbTwips

    If Me.optTimeOption(0).Value Then
        Me.lbl24hformat.Caption = "(24 hour format)"
'        Me.Frame1.Width = 1700
'        Me.lblSecSeparator.Visible = False
'        Me.txtSecond.Visible = False
        Me.Frame1.Width = 2200
        Me.lblSecSeparator.Visible = True
        Me.txtSecond.Visible = True
    Else
        Me.lbl24hformat.Caption = "(less than 24 hours)"
        Me.Frame1.Width = 2200
        Me.lblSecSeparator.Visible = True
        Me.txtSecond.Visible = True
    End If
    
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
    Me.vsHour.Value = CInt(Me.txtHour.Text)
    Me.txtMinute.Text = GetMySetting(RegistryName, "Minute", "00")
    Me.vsMinute.Value = CInt(Me.txtMinute.Text)
    Me.txtSecond.Text = GetMySetting(RegistryName, "Second", "00")
    
    Me.optTimeOption(CInt(GetMySetting(RegistryName, "AbsoluteRelativeTime", "0"))).Value = True
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtHour_Validate(Cancel)
    If Not Cancel Then Call txtMinute_Validate(Cancel)
    If Not Cancel Then Call txtSecond_Validate(Cancel)
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "Hour", Me.txtHour.Text)
    Call SaveMySetting(RegistryName, "Minute", Me.txtMinute.Text)
    Call SaveMySetting(RegistryName, "Second", Me.txtSecond.Text)
    If Me.optTimeOption(0).Value Then
        Call SaveMySetting(RegistryName, "AbsoluteRelativeTime", "0")
    Else
        Call SaveMySetting(RegistryName, "AbsoluteRelativeTime", "1")
    End If
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As WaitForTimeAction)
    clsAction.Hour = Me.txtHour.Text
    clsAction.Minute = Me.txtMinute.Text
    clsAction.Second = Me.txtSecond.Text
    clsAction.AbsoluteTime = Me.optTimeOption(0).Value
End Sub

Public Sub GetFormDataFromClass(clsAction As WaitForTimeAction)
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
    
    If clsAction.Second < 10 Then
        Me.txtSecond.Text = "0" & clsAction.Second
    Else
        Me.txtSecond.Text = clsAction.Second
    End If
    
    If clsAction.AbsoluteTime Then
        Me.optTimeOption(0).Value = True
    Else
        Me.optTimeOption(1).Value = True
    End If
    
    Call SaveMySetting(RegistryName, "Hour", Me.txtHour.Text)
    Call SaveMySetting(RegistryName, "Minute", Me.txtMinute.Text)
    Call SaveMySetting(RegistryName, "Second", Me.txtSecond.Text)
    If Me.optTimeOption(0).Value Then
        Call SaveMySetting(RegistryName, "AbsoluteRelativeTime", "0")
    Else
        Call SaveMySetting(RegistryName, "AbsoluteRelativeTime", "1")
    End If
End Sub

Private Sub optTimeOption_Click(Index As Integer)
    Me.ScaleMode = vbTwips
    If Me.optTimeOption(0).Value Then
        Me.lbl24hformat.Caption = "(24 hour format)"
'        Me.Frame1.Width = 1700
'        Me.lblSecSeparator.Visible = False
'        Me.txtSecond.Visible = False
        Me.Frame1.Width = 2200
        Me.lblSecSeparator.Visible = True
        Me.txtSecond.Visible = True
    Else
        Me.lbl24hformat.Caption = "(less than 24 hours)"
        Me.Frame1.Width = 2200
        Me.lblSecSeparator.Visible = True
        Me.txtSecond.Visible = True
    End If
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
        Me.vsHour.Value = CInt(Me.txtHour.Text)
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
        Me.vsMinute.Value = CInt(Me.txtMinute.Text)
        If CInt(Me.txtMinute.Text) < 10 Then
            Me.txtMinute.Text = "0" & CInt(Me.txtMinute.Text)
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub txtSecond_GotFocus()
    Me.txtSecond.SelStart = 0
    Me.txtSecond.SelLength = Len(Me.txtSecond.Text)
End Sub

Private Sub txtSecond_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtSecond.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtSecond.Text Then
        Beep
        Cancel = True
    Else
        Me.vsHour.Value = CInt(Me.txtSecond.Text)
        If CInt(Me.txtSecond.Text) < 10 Then
            Me.txtSecond.Text = "0" & CInt(Me.txtSecond.Text)
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub vsHour_Change()
    'Me.txtHour.Text = Me.vsHour.Value
End Sub

Private Sub vsMinute_Change()
    'Me.txtMinute.Text = Me.vsMinute.Value
End Sub

