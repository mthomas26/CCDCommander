VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRunActionList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run Action List"
   ClientHeight    =   2190
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6600
   HelpContextID   =   1400
   Icon            =   "frmRunActionList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpen 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6180
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmRunActionList.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   130
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.TextBox txtTime 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Top             =   1020
      Width           =   795
   End
   Begin VB.TextBox txtNumRepeat 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   795
   End
   Begin VB.OptionButton optRepeat 
      Caption         =   "Run until Aborted"
      Height          =   195
      Index           =   3
      Left            =   1020
      TabIndex        =   3
      Top             =   1380
      Width           =   2115
   End
   Begin VB.OptionButton optRepeat 
      Caption         =   "Run for a Period of Time"
      Height          =   195
      Index           =   2
      Left            =   1020
      TabIndex        =   2
      Top             =   1080
      Width           =   2115
   End
   Begin VB.OptionButton optRepeat 
      Caption         =   "Run Multiple Times"
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   780
      Width           =   1875
   End
   Begin VB.OptionButton optRepeat 
      Caption         =   "Run Once"
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   480
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   3150
      Top             =   1695
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.TextBox txtScript 
      BackColor       =   &H8000000F&
      Height          =   288
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   5115
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3983
      TabIndex        =   8
      Top             =   1695
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1403
      TabIndex        =   6
      Top             =   1695
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "minutes"
      Height          =   195
      Left            =   4380
      TabIndex        =   11
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "times"
      Height          =   195
      Left            =   4380
      TabIndex        =   10
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Action List:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmRunActionList"
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

Private Const RegistryName = "RunActionList"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub cmdOpen_Click()
    Me.ComDlg.FileName = Me.txtScript.Text
    Me.ComDlg.Filter = "Action List (*.act)|*.act"
    Me.ComDlg.FilterIndex = 1
    Me.ComDlg.CancelError = True
    Me.ComDlg.InitDir = Me.txtScript.Text
    Me.ComDlg.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.ComDlg.ShowOpen
    If Err.Number = 0 Then
        Me.txtScript.Text = Me.ComDlg.FileName
    Else
        Me.txtScript.Text = ""
    End If
    On Error GoTo 0
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
    Me.txtScript.Text = GetMySetting(RegistryName, "ActionListFileName", "")
    Me.optRepeat(GetMySetting(RegistryName, "RepeatOption", "0")).Value = True
    Me.txtNumRepeat.Text = GetMySetting(RegistryName, "TimesToRepeat", "2")
    Me.txtTime.Text = Format(GetMySetting(RegistryName, "RepeatTime", "1"), "0.00")
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtNumRepeat_Validate(Cancel)
    
    
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "ActionListFileName", Me.txtScript.Text)
    
    If Me.optRepeat(0).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "0")
    ElseIf Me.optRepeat(1).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "1")
    ElseIf Me.optRepeat(2).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "2")
    ElseIf Me.optRepeat(3).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "3")
    End If
    
    Call SaveMySetting(RegistryName, "TimesToRepeat", Me.txtNumRepeat.Text)
    Call SaveMySetting(RegistryName, "RepeatTime", Me.txtTime.Text)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As RunActionList)
    clsAction.ActionListName = Me.txtScript.Text
    
    If Me.optRepeat(0).Value Then
        clsAction.RepeatMode = 0
    ElseIf Me.optRepeat(1).Value Then
        clsAction.RepeatMode = 1
    ElseIf Me.optRepeat(2).Value Then
        clsAction.RepeatMode = 2
    ElseIf Me.optRepeat(3).Value Then
        clsAction.RepeatMode = 3
    End If
    
    clsAction.TimesToRepeat = CInt(Me.txtNumRepeat.Text)
    clsAction.RepeatTime = CDbl(Me.txtTime.Text)
    
End Sub

Public Sub GetFormDataFromClass(clsAction As RunActionList)
    Me.txtScript.Text = clsAction.ActionListName
    Me.optRepeat(clsAction.RepeatMode).Value = True
    Me.txtNumRepeat.Text = clsAction.TimesToRepeat
    Me.txtTime.Text = Format(clsAction.RepeatTime, "0.00")

    Call SaveMySetting(RegistryName, "ActionListFileName", Me.txtScript.Text)
    If Me.optRepeat(0).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "0")
    ElseIf Me.optRepeat(1).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "1")
    ElseIf Me.optRepeat(2).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "2")
    ElseIf Me.optRepeat(3).Value Then
        Call SaveMySetting(RegistryName, "RepeatOption", "3")
    End If
    Call SaveMySetting(RegistryName, "TimesToRepeat", Me.txtNumRepeat.Text)
    Call SaveMySetting(RegistryName, "RepeatTime", Me.txtTime.Text)
End Sub

Private Sub optRepeat_Click(Index As Integer)
    If Me.optRepeat(1).Value = True Then
        Me.txtNumRepeat.Enabled = True
        Me.txtTime.Enabled = False
    ElseIf Me.optRepeat(2).Value = True Then
        Me.txtNumRepeat.Enabled = False
        Me.txtTime.Enabled = True
    Else
        Me.txtNumRepeat.Enabled = False
        Me.txtTime.Enabled = False
    End If
End Sub

Private Sub txtNumRepeat_GotFocus()
    Me.txtNumRepeat.SelStart = 0
    Me.txtNumRepeat.SelLength = Len(Me.txtNumRepeat.Text)
End Sub

Private Sub txtNumRepeat_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtNumRepeat.Text)
    If Err.Number <> 0 Or Test < 2 Or Test <> Me.txtNumRepeat.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtTime_GotFocus()
    Me.txtTime.SelStart = 0
    Me.txtTime.SelLength = Len(Me.txtTime.Text)
End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtTime.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtTime.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub


