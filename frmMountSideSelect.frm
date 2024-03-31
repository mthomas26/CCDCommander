VERSION 5.00
Begin VB.Form frmMountSideSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GEM Side Selection"
   ClientHeight    =   2430
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   ControlBox      =   0   'False
   HelpContextID   =   204
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmMountSideSelect.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   420
      Width           =   480
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdWest 
      Caption         =   "West"
      Height          =   375
      Left            =   2340
      TabIndex        =   1
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdEast 
      Caption         =   "East"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Test"
      Height          =   1755
      Left            =   720
      TabIndex        =   3
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmMountSideSelect"
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

Private Sub cmdAuto_Click()
    Me.Tag = "Auto"
    Me.Hide
End Sub

Private Sub cmdEast_Click()
    Me.Tag = "East"
    Me.Hide
End Sub

Private Sub cmdWest_Click()
    Me.Tag = "West"
    Me.Hide
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    Me.Label1 = "CCD Commander cannot determine which side of the mount your telescope is on." & vbCrLf & vbCrLf & _
        "Please select below the side of the mount your telescope is currently on." & vbCrLf & vbCrLf & _
        "Or push the auto button to allow CCD Commander to force the mount to one side or the other."
End Sub
