VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRunScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run Program Action"
   ClientHeight    =   3210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6780
   HelpContextID   =   1300
   Icon            =   "frmRunScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtArgs 
      Height          =   288
      Left            =   900
      TabIndex        =   9
      Top             =   540
      Width           =   5355
   End
   Begin VB.Frame Frame1 
      Caption         =   "Program is a:"
      Height          =   915
      Left            =   1800
      TabIndex        =   6
      Top             =   900
      Width           =   3255
      Begin VB.OptionButton optProgramType 
         Caption         =   "Executable (e.g: .exe, .com, .bat)"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   540
         Width           =   2715
      End
      Begin VB.OptionButton optProgramType 
         Caption         =   "Script (e.g. .vbs, .wsf)"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
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
      Left            =   6300
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmRunScript.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1470
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtScript 
      BackColor       =   &H8000000F&
      Height          =   288
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   180
      Width           =   5535
   End
   Begin VB.CheckBox chkWait 
      Caption         =   "Wait for program to finish"
      Height          =   252
      Left            =   2340
      TabIndex        =   1
      Top             =   1980
      Value           =   1  'Checked
      Width           =   2172
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   3210
      Top             =   2520
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Arguments:"
      Height          =   195
      Left            =   45
      TabIndex        =   10
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Program:"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmRunScript"
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

Private Const RegistryName = "RunScriptAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub


Private Sub cmdOpen_Click()
    Me.ComDlg.FileName = Me.txtScript.Text
    Me.ComDlg.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf|Executable (*.exe;*.com;*.bat)|*.exe;*.com;*.bat"
    Me.ComDlg.FilterIndex = 1
    Me.ComDlg.CancelError = True
    Me.ComDlg.InitDir = Me.txtScript.Text
    Me.ComDlg.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.ComDlg.ShowOpen
    If Err.Number = 0 Then
        Me.txtScript.Text = Me.ComDlg.FileName
        If Me.ComDlg.FilterIndex = 1 Then
            Me.optProgramType(0).Value = True
        Else
            Me.optProgramType(1).Value = True
        End If
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
    Me.txtScript.Text = GetMySetting(RegistryName, "ScriptFileName", "")
    Me.chkWait.Value = CInt(GetMySetting(RegistryName, "WaitForScript", CStr(vbChecked)))
    Me.optProgramType(0).Value = CBool(GetMySetting(RegistryName, "ProgramIsScript", "True"))
    Me.txtArgs.Text = GetMySetting(RegistryName, "ScriptArguments", "")
End Sub

Private Sub OKButton_Click()
    Call SaveMySetting(RegistryName, "ScriptFileName", Me.txtScript.Text)
    Call SaveMySetting(RegistryName, "WaitForScript", Me.chkWait.Value)
    Call SaveMySetting(RegistryName, "ProgramIsScript", Me.optProgramType(0).Value)
    Call SaveMySetting(RegistryName, "ScriptArguments", Me.txtArgs.Text)
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As RunScriptAction)
    clsAction.ScriptName = Me.txtScript.Text
    clsAction.WaitForScript = Me.chkWait.Value
    clsAction.ProgramIsScript = Me.optProgramType(0).Value
    clsAction.ScriptArguments = Me.txtArgs.Text
End Sub

Public Sub GetFormDataFromClass(clsAction As RunScriptAction)
    Me.txtScript.Text = clsAction.ScriptName
    Me.chkWait.Value = clsAction.WaitForScript
    Me.optProgramType(0).Value = clsAction.ProgramIsScript
    If Not Me.optProgramType(0).Value Then
        Me.optProgramType(1).Value = True
    End If
    Me.txtArgs.Text = clsAction.ScriptArguments

    Call SaveMySetting(RegistryName, "ScriptFileName", Me.txtScript.Text)
    Call SaveMySetting(RegistryName, "WaitForScript", Me.chkWait.Value)
    Call SaveMySetting(RegistryName, "ProgramIsScript", Me.optProgramType(0).Value)
    Call SaveMySetting(RegistryName, "ScriptArguments", Me.txtArgs.Text)
End Sub

