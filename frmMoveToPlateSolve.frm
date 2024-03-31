VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMoveToPlateSolve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get RA/Dec from Current Telescope Position"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5490
   Icon            =   "frmMoveToPlateSolve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   480
      Top             =   2580
   End
   Begin VB.CommandButton cmdAbortClose 
      Caption         =   "Abort"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2580
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   2415
      HelpContextID   =   110
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Current Status"
      Top             =   0
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4260
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMoveToPlateSolve.frx":030A
   End
End
Attribute VB_Name = "frmMoveToPlateSolve"
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

Public myImageLink As ImageLinkSyncAction
Public Precess As Boolean

Private Sub cmdAbortClose_Click()
    If Me.cmdAbortClose.Caption = "Abort" Then
        'abort!
        MainMod.Aborted = True
    Else
        'close!
        MainMod.Aborted = False
        Me.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    MainMod.MoveToRADecPlateSolveStatus = True
    
    Call AddToStatus("Connecting to camera...")
    Call Camera.CameraSetup
    Call AddToStatus("Connecting to mount...")
    Call Mount.MountSetup
    Call Mount.ConnectToTelescope
    If Not Camera.TakeImageAndLink(myImageLink, True, Precess) Then
        Call MsgBox("Plate solve errror!", vbOKOnly + vbCritical, "Get RA/Dec from Telescope")
        Me.Tag = "0"
    Else
        Call MsgBox("Plate solve successful.", vbOKOnly + vbInformation, "Get RA/Dec from Telescope")
        Me.Tag = "1"
    End If
    frmMoveToPlateSolve.cmdAbortClose.Caption = "Close"

    MainMod.MoveToRADecPlateSolveStatus = False
End Sub
