VERSION 5.00
Begin VB.Form frmComment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comment"
   ClientHeight    =   3810
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDoNotSkip 
      Caption         =   "Do not skip this action"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3420
      Width           =   4335
   End
   Begin VB.CheckBox chkSendAsEmail 
      Caption         =   "Send Comment as E-Mail"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox txtComment 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      ToolTipText     =   "Maximum of 80 characters will be visible in the action list.  Entire text will be visible in the log."
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmComment"
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

Private Const RegistryName = "Comment"

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
    Me.txtComment.Text = GetMySetting(RegistryName, "Comment", "")
    Me.chkSendAsEmail.Value = GetMySetting(RegistryName, "SendAsEMail", "0")
    Me.chkDoNotSkip.Value = GetMySetting(RegistryName, "DoNotSkip", "1")
End Sub

Private Sub OKButton_Click()
    Call SaveMySetting(RegistryName, "Comment", Me.txtComment.Text)
    Call SaveMySetting(RegistryName, "SendAsEmail", Me.chkSendAsEmail.Value)
    Call SaveMySetting(RegistryName, "DoNotSkip", Me.chkDoNotSkip.Value)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As clsCommentAction)
    clsAction.Comment = Me.txtComment.Text
    If Me.chkSendAsEmail.Value = vbChecked Then
        clsAction.SendAsEmail = True
    Else
        clsAction.SendAsEmail = False
    End If
    If Me.chkDoNotSkip.Value = vbChecked Then
        clsAction.DoNotSkip = True
    Else
        clsAction.DoNotSkip = False
    End If
End Sub

Public Sub GetFormDataFromClass(clsAction As clsCommentAction)
    Me.txtComment.Text = clsAction.Comment
    
    If clsAction.SendAsEmail Then
        Me.chkSendAsEmail.Value = vbChecked
    Else
        Me.chkSendAsEmail.Value = vbUnchecked
    End If
    
    If clsAction.DoNotSkip Then
        Me.chkDoNotSkip.Value = vbChecked
    Else
        Me.chkDoNotSkip.Value = vbUnchecked
    End If
    
    Call SaveMySetting(RegistryName, "Comment", Me.txtComment.Text)
    Call SaveMySetting(RegistryName, "SendAsEmail", Me.chkSendAsEmail.Value)
    Call SaveMySetting(RegistryName, "DoNotSkip", Me.chkDoNotSkip.Value)
End Sub


