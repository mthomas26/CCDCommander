VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImageLinkSync 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plate Solve and Sync or Offset Action"
   ClientHeight    =   4950
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6165
   HelpContextID   =   700
   Icon            =   "frmImageLinkSync.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSaveTo 
      Height          =   315
      Left            =   1680
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "C:\"
      Top             =   4020
      Width           =   4095
   End
   Begin VB.TextBox txtFileNamePrefix 
      Height          =   285
      Left            =   1680
      TabIndex        =   28
      Top             =   4620
      Width           =   4095
   End
   Begin VB.CheckBox chkUseGlobalImageSaveLocation 
      Caption         =   "Use Global Image Save Location"
      Height          =   312
      Left            =   60
      TabIndex        =   32
      Top             =   3720
      Value           =   1  'Checked
      Width           =   3135
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
      Left            =   5820
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmImageLinkSync.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4080
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.CheckBox chkAutosave 
      Caption         =   "Autosave Exposures"
      Height          =   312
      Left            =   60
      TabIndex        =   29
      Top             =   4380
      Width           =   1812
   End
   Begin VB.CommandButton cmdAutoFileNameBuilder 
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
      Left            =   5820
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmImageLinkSync.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "File Name Builder"
      Top             =   4620
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.Frame fraRetry 
      Caption         =   "Retry"
      Height          =   1215
      Left            =   60
      TabIndex        =   25
      Top             =   2400
      Width           =   6015
      Begin VB.CheckBox chkSkip 
         Caption         =   "Skip to Next Target if Second Solve Succeeds"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   5175
      End
      Begin VB.TextBox txtSlewOnFailAmount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   240
         Left            =   2760
         TabIndex        =   13
         Text            =   "5"
         Top             =   520
         Width           =   495
      End
      Begin VB.CheckBox chkSlewOnFail 
         Caption         =   "Slew Mount if First Solve Fails"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   540
         Width           =   2535
      End
      Begin VB.CheckBox chkRetry 
         Caption         =   "Retry Plate Solve on Failure"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label7 
         Caption         =   "arcminutes toward the zenith."
         Height          =   255
         Left            =   3300
         TabIndex        =   26
         Top             =   540
         Width           =   2475
      End
   End
   Begin VB.CheckBox chkAbort 
      Alignment       =   1  'Right Justify
      Caption         =   "Abort List/Sub-List if Solve Fails"
      Height          =   195
      Left            =   2940
      TabIndex        =   10
      Top             =   1980
      Width           =   3075
   End
   Begin VB.CheckBox chkSlew 
      Alignment       =   1  'Right Justify
      Caption         =   "Slew to Original Location after Solve"
      Height          =   252
      Left            =   2940
      TabIndex        =   9
      Top             =   1440
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin VB.Frame fraSync 
      Caption         =   "Sync Selection"
      Height          =   1212
      Left            =   2880
      TabIndex        =   22
      Top             =   60
      Width           =   1812
      Begin VB.OptionButton optSync 
         Caption         =   "Offset Position"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   660
         Width           =   1512
      End
      Begin VB.OptionButton optSync 
         Caption         =   "Sync Mount"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1212
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plate Solve Exposure Information"
      Height          =   2295
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   2775
      Begin VB.TextBox txtDelay 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   1380
         TabIndex        =   1
         Text            =   "0"
         Top             =   540
         Width           =   372
      End
      Begin VB.OptionButton optFrameSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Frame"
         Height          =   252
         Index           =   0
         Left            =   420
         TabIndex        =   4
         Top             =   1500
         Width           =   1152
      End
      Begin VB.ComboBox cmbBin 
         Height          =   315
         ItemData        =   "frmImageLinkSync.frx":05A2
         Left            =   1380
         List            =   "frmImageLinkSync.frx":05B2
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   912
      End
      Begin VB.TextBox txtExposureTime 
         Alignment       =   1  'Right Justify
         Height          =   252
         Left            =   1380
         TabIndex        =   0
         Text            =   "5"
         Top             =   240
         Width           =   372
      End
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         ItemData        =   "frmImageLinkSync.frx":05CA
         Left            =   1380
         List            =   "frmImageLinkSync.frx":05DD
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   912
      End
      Begin VB.OptionButton optFrameSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Quarter Frame"
         Height          =   252
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1980
         Width           =   1392
      End
      Begin VB.OptionButton optFrameSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Half Frame"
         Height          =   252
         Index           =   1
         Left            =   420
         TabIndex        =   5
         Top             =   1740
         Value           =   -1  'True
         Width           =   1152
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Delay Before Exp:"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   570
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   1800
         TabIndex        =   23
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Bin:"
         Height          =   195
         Left            =   1020
         TabIndex        =   21
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Exposure Time:"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Filter:"
         Height          =   195
         Left            =   915
         TabIndex        =   18
         Top             =   1140
         Width           =   435
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   16
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4860
      TabIndex        =   15
      Top             =   180
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog MSComm 
      Left            =   5280
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblImageSaveLocation 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Image Save Location:"
      Height          =   195
      Left            =   60
      TabIndex        =   36
      Top             =   4080
      Width           =   1560
   End
   Begin VB.Label lblFilenamePrefix 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Filename Prefix"
      Height          =   195
      Left            =   420
      TabIndex        =   35
      Top             =   4680
      Width           =   1110
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Image Save Location:"
      Height          =   195
      Left            =   60
      TabIndex        =   34
      Top             =   360
      Width           =   1560
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Filename Prefix"
      Height          =   195
      Left            =   480
      TabIndex        =   33
      Top             =   960
      Width           =   1110
   End
End
Attribute VB_Name = "frmImageLinkSync"
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

Private Const RegistryName = "ImageLinkAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub GetSettings()
    Me.cmbBin.ListIndex = CInt(GetMySetting(RegistryName, "Bin", "0"))
    If CInt(GetMySetting(RegistryName, "Filter", "0")) < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = CInt(GetMySetting(RegistryName, "Filter", "0"))
    End If
    Me.optFrameSize(CInt(GetMySetting(RegistryName, "FrameSize", "0"))).Value = True
    Me.txtExposureTime.Text = GetMySetting(RegistryName, "ExposureTime", "5")
    Me.txtDelay.Text = GetMySetting(RegistryName, "DelayTime", "0")
    Me.optSync(CInt(GetMySetting(RegistryName, "SyncMode", "0"))).Value = True
    Me.chkSlew.Value = GetMySetting(RegistryName, "SlewToOriginalLocation", "1")
    
    Me.chkAbort.Value = GetMySetting(RegistryName, "AbortIfFail", "0")
    Me.chkRetry.Value = GetMySetting(RegistryName, "Retry", "0")
    Me.chkSlewOnFail.Value = GetMySetting(RegistryName, "SlewOnFail", "0")
    Me.txtSlewOnFailAmount.Text = GetMySetting(RegistryName, "SlewOnFailAmount", "1")
    Me.chkSkip.Value = GetMySetting(RegistryName, "SkipOnFail", "0")
    
    Me.chkAutosave.Value = CInt(GetMySetting(RegistryName, "Autosave", "0"))
    Me.txtFileNamePrefix.Text = GetMySetting(RegistryName, "FileNamePrefix", "")
    Me.chkUseGlobalImageSaveLocation.Value = CInt(GetMySetting(RegistryName, "UseGlobalImageSaveLocation", "1"))
    Me.txtSaveTo.Text = GetMySetting(RegistryName, "SaveToPath", frmOptions.txtSaveTo.Text)
    Call chkUseGlobalImageSaveLocation_Click
End Sub

Private Sub chkAutosave_Click()
    If Me.chkAutosave.Value = vbChecked Then
        Me.txtFileNamePrefix.Enabled = True
        Me.cmdAutoFileNameBuilder.Enabled = True
    Else
        Me.txtFileNamePrefix.Enabled = False
        Me.cmdAutoFileNameBuilder.Enabled = False
    End If
End Sub

Private Sub chkRetry_Click()
    If chkRetry.Value = vbChecked Then
        Me.chkSlewOnFail.Enabled = True
        Me.chkSkip.Enabled = True
    Else
        Me.chkSlewOnFail.Value = vbUnchecked
        Me.chkSlewOnFail.Enabled = False
        Me.chkSkip.Value = vbUnchecked
        Me.chkSkip.Enabled = False
    End If
End Sub

Private Sub chkSlewOnFail_Click()
    If Me.chkSlewOnFail.Value = vbChecked Then
        Me.txtSlewOnFailAmount.Enabled = True
    Else
        Me.txtSlewOnFailAmount.Enabled = False
    End If
End Sub

Private Sub chkUseGlobalImageSaveLocation_Click()
    If Me.chkUseGlobalImageSaveLocation.Value = vbChecked Then
        Me.txtSaveTo.Text = frmOptions.txtSaveTo.Text
        Me.cmdOpen.Enabled = False
        Me.txtSaveTo.Enabled = False
    Else
        Me.cmdOpen.Enabled = True
        Me.txtSaveTo.Enabled = True
    End If
End Sub

Private Sub cmdAutoFileNameBuilder_Click()
    frmFileNameBuilder.txtFileNamePrefix.Text = Me.txtFileNamePrefix.Text
    frmFileNameBuilder.Show vbModal, Me
    If frmFileNameBuilder.Tag <> "" Then
        Me.txtFileNamePrefix.Text = frmFileNameBuilder.Tag
    End If
    Unload frmFileNameBuilder
End Sub

Private Sub cmdOpen_Click()
    Me.MSComm.Filter = "Folders|Folders"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.DialogTitle = "Select the folder to save your FITS files to..."
    Me.MSComm.FileName = "Select folder"
    Me.MSComm.InitDir = Me.txtSaveTo.Text
    Me.MSComm.flags = cdlOFNHideReadOnly + cdlOFNNoValidate + cdlOFNPathMustExist
    Me.MSComm.CancelError = True
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtSaveTo.Text = Left(Me.MSComm.FileName, InStrRev(Me.MSComm.FileName, "\"))
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    Dim Counter As Integer
    
    Call MainMod.SetOnTopMode(Me)
    
    Me.cmbFilter.Clear
    For Counter = 0 To frmOptions.lstFilters.ListCount - 1
        Call Me.cmbFilter.AddItem(frmOptions.lstFilters.List(Counter))
    Next Counter
    
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

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtExposureTime_Validate(Cancel)
    If Not Cancel Then Call txtDelay_Validate(Cancel)
    
    If Cancel Then
        Beep
        Exit Sub
    End If

    Call SaveMySetting(RegistryName, "Bin", Me.cmbBin.ListIndex)
    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    If Me.optFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "0")
    ElseIf Me.optFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "1")
    ElseIf Me.optFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "2")
    End If
    Call SaveMySetting(RegistryName, "ExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "DelayTime", Me.txtDelay.Text)
    
    If Me.optSync(0).Value = True Then
        Call SaveMySetting(RegistryName, "SyncMode", "0")
    Else
        Call SaveMySetting(RegistryName, "SyncMode", "1")
    End If
    Call SaveMySetting(RegistryName, "SlewToOriginalLocation", Me.chkSlew.Value)
    
    Call SaveMySetting(RegistryName, "AbortIfFail", Me.chkAbort.Value)
    Call SaveMySetting(RegistryName, "Retry", Me.chkRetry.Value)
    Call SaveMySetting(RegistryName, "SlewOnFail", Me.chkSlewOnFail.Value)
    Call SaveMySetting(RegistryName, "SlewOnFailAmount", Me.txtSlewOnFailAmount.Text)
    Call SaveMySetting(RegistryName, "SkipOnFail", Me.chkSkip.Value)
    
    Call SaveMySetting(RegistryName, "Autosave", Me.chkAutosave.Value)
    Call SaveMySetting(RegistryName, "FileNamePrefix", Me.txtFileNamePrefix.Text)
    Call SaveMySetting(RegistryName, "UseGlobalImageSaveLocation", Me.chkUseGlobalImageSaveLocation.Value)
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub optFrameSize_Click(Index As Integer)
    Me.optFrameSize(Index).Value = True
    
    If Index = 0 Then
        Me.optFrameSize(1).Value = False
        Me.optFrameSize(2).Value = False
    ElseIf Index = 1 Then
        Me.optFrameSize(0).Value = False
        Me.optFrameSize(2).Value = False
    ElseIf Index = 2 Then
        Me.optFrameSize(0).Value = False
        Me.optFrameSize(1).Value = False
    End If
End Sub

Public Sub PutFormDataIntoClass(clsAction As ImageLinkSyncAction)
    clsAction.Bin = Me.cmbBin.ListIndex
    clsAction.ExpTime = CDbl(Me.txtExposureTime.Text)
    clsAction.DelayTime = CDbl(Me.txtDelay.Text)
    clsAction.Filter = Me.cmbFilter.ListIndex
    If Me.optFrameSize(0).Value Then
        clsAction.FrameSize = FullFrame
    ElseIf Me.optFrameSize(1).Value Then
        clsAction.FrameSize = HalfFrame
    ElseIf Me.optFrameSize(2).Value Then
        clsAction.FrameSize = QuarterFrame
    End If
    If Me.optSync(0).Value Then
        clsAction.SyncMode = MountSync
    ElseIf Me.optSync(1).Value Then
        clsAction.SyncMode = Offset
    End If
    clsAction.SlewToOriginalLocation = Me.chkSlew.Value
    
    clsAction.AbortListOnFailure = CBool(Me.chkAbort.Value)
    clsAction.RetryPlateSolveOnFailure = CBool(Me.chkRetry.Value)
    clsAction.SlewMountForRetry = CBool(Me.chkSlewOnFail.Value)
    clsAction.ArcminutesToSlew = CLng(Me.txtSlewOnFailAmount.Text)
    clsAction.SkipIfRetrySucceeds = CBool(Me.chkSkip.Value)
    
    clsAction.AutosaveExposure = Me.chkAutosave.Value
    clsAction.FileNamePrefix = Me.txtFileNamePrefix.Text
    clsAction.UseGlobalSaveToLocation = Me.chkUseGlobalImageSaveLocation.Value
    clsAction.FileSavePath = Me.txtSaveTo.Text
End Sub

Public Sub GetFormDataFromClass(clsAction As ImageLinkSyncAction)
    Me.cmbBin.ListIndex = clsAction.Bin
    Me.txtExposureTime.Text = clsAction.ExpTime
    Me.txtDelay.Text = clsAction.DelayTime
    If clsAction.Filter < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = clsAction.Filter
    End If
    If clsAction.FrameSize = FullFrame Then
        Me.optFrameSize(0).Value = True
    ElseIf clsAction.FrameSize = HalfFrame Then
         Me.optFrameSize(1).Value = True
    ElseIf clsAction.FrameSize = QuarterFrame Then
        Me.optFrameSize(2).Value = True
    End If
    If clsAction.SyncMode = MountSync Then
        Me.optSync(0).Value = True
    ElseIf clsAction.SyncMode = Offset Then
        Me.optSync(1).Value = True
    End If
    Me.chkSlew.Value = clsAction.SlewToOriginalLocation

    If clsAction.AbortListOnFailure Then
        Me.chkAbort.Value = vbChecked
    Else
        Me.chkAbort.Value = vbUnchecked
    End If
    
    If clsAction.RetryPlateSolveOnFailure Then
        Me.chkRetry.Value = vbChecked
    Else
        Me.chkRetry.Value = vbUnchecked
    End If
    
    If clsAction.SlewMountForRetry Then
        Me.chkSlewOnFail.Value = vbChecked
    Else
        Me.chkSlewOnFail.Value = vbUnchecked
    End If
    
    Me.txtSlewOnFailAmount.Text = clsAction.ArcminutesToSlew
    
    If clsAction.SkipIfRetrySucceeds Then
        Me.chkSkip.Value = vbChecked
    Else
        Me.chkSkip.Value = vbUnchecked
    End If
    
    Me.chkAutosave.Value = clsAction.AutosaveExposure
    Me.txtFileNamePrefix.Text = clsAction.FileNamePrefix
    Me.chkUseGlobalImageSaveLocation.Value = clsAction.UseGlobalSaveToLocation
    Me.txtSaveTo.Text = clsAction.FileSavePath

    Call SaveMySetting(RegistryName, "Bin", Me.cmbBin.ListIndex)
    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    If Me.optFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "0")
    ElseIf Me.optFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "1")
    ElseIf Me.optFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "2")
    End If
    Call SaveMySetting(RegistryName, "ExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "DelayTime", Me.txtDelay.Text)
    If Me.optSync(0).Value = True Then
        Call SaveMySetting(RegistryName, "SyncMode", "0")
    Else
        Call SaveMySetting(RegistryName, "SyncMode", "1")
    End If
    Call SaveMySetting(RegistryName, "SlewToOriginalLocation", Me.chkSlew.Value)
    Call SaveMySetting(RegistryName, "AbortIfFail", Me.chkAbort.Value)
    Call SaveMySetting(RegistryName, "Retry", Me.chkRetry.Value)
    Call SaveMySetting(RegistryName, "SlewOnFail", Me.chkSlewOnFail.Value)
    Call SaveMySetting(RegistryName, "SlewOnFailAmount", Me.txtSlewOnFailAmount.Text)
    Call SaveMySetting(RegistryName, "SkipOnFail", Me.chkSkip.Value)
    
    Call SaveMySetting(RegistryName, "Autosave", Me.chkAutosave.Value)
    Call SaveMySetting(RegistryName, "FileNamePrefix", Me.txtFileNamePrefix.Text)
    Call SaveMySetting(RegistryName, "UseGlobalImageSaveLocation", Me.chkUseGlobalImageSaveLocation.Value)
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)
End Sub

Private Sub optSync_Click(Index As Integer)
    Me.optSync(Index).Value = True
    
    If Index = 0 Then
        Me.optSync(1).Value = False
    Else
        Me.optSync(0).Value = False
    End If
End Sub

Private Sub txtDelay_GotFocus()
    Me.txtDelay.SelStart = 0
    Me.txtDelay.SelLength = Len(Me.txtDelay.Text)
End Sub

Private Sub txtDelay_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDelay.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtDelay.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtExposureTime_GotFocus()
    Me.txtExposureTime.SelStart = 0
    Me.txtExposureTime.SelLength = Len(Me.txtExposureTime.Text)
End Sub

Private Sub txtExposureTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtExposureTime.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtExposureTime.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtFileNamePrefix_GotFocus()
    Me.txtFileNamePrefix.SelStart = 0
    Me.txtFileNamePrefix.SelLength = Len(Me.txtFileNamePrefix.Text)
End Sub

Private Sub txtSlewOnFailAmount_GotFocus()
    Me.txtSlewOnFailAmount.SelStart = 0
    Me.txtSlewOnFailAmount.SelLength = Len(Me.txtSlewOnFailAmount.Text)
End Sub

Private Sub txtSlewOnFailAmount_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CLng(Me.txtSlewOnFailAmount.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtSlewOnFailAmount.Text Or Test > (90 * 60) Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub
