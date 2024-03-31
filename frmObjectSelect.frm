VERSION 5.00
Begin VB.Form frmObjectSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TheSky6 Selected Objects"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   HelpContextID   =   800
   Icon            =   "frmObjectSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Default         =   -1  'True
      Height          =   375
      Left            =   2288
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   488
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox lstObjects 
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Select the object you wish to use from the list below and click the Ok button:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmObjectSelect"
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

Private Sub cmdRefresh_Click()
    Me.Tag = "1"
    Me.Hide
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
End Sub

Private Sub OKButton_Click()
    Me.Tag = "0"
    Me.Hide
End Sub
