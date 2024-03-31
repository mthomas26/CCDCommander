VERSION 5.00
Begin VB.Form frmFocusAction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Focus Action"
   ClientHeight    =   1815
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4305
   HelpContextID   =   600
   Icon            =   "frmFocusAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraAverages 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CheckBox chkTempComp 
         Alignment       =   1  'Right Justify
         Caption         =   "Temperature Compensation"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtAverages 
         Height          =   285
         Left            =   1050
         TabIndex        =   16
         Text            =   "1"
         Top             =   120
         Width           =   500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Averages:"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.CheckBox chkSpecifyStarPos 
      Alignment       =   1  'Right Justify
      Caption         =   "Specify Focus Star Position"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   780
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtExposureTime 
      Height          =   252
      Left            =   1440
      TabIndex        =   1
      Text            =   "300"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      ItemData        =   "frmFocusAction.frx":030A
      Left            =   885
      List            =   "frmFocusAction.frx":031D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
   Begin VB.Frame fraFocusMaxStarPos 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtStarXPos 
         Height          =   285
         Left            =   1050
         TabIndex        =   11
         Text            =   "50"
         Top             =   120
         Width           =   500
      End
      Begin VB.TextBox txtStarYPos 
         Height          =   285
         Left            =   1050
         TabIndex        =   10
         Text            =   "50"
         Top             =   420
         Width           =   500
      End
      Begin VB.CommandButton cmdGetStarPos 
         Caption         =   "Get Position"
         Height          =   555
         Left            =   1560
         TabIndex        =   9
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Star X pos:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Star Y pos:"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   450
         Width           =   780
      End
   End
   Begin VB.Label lblExposureTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Exposure Time:"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblExposureTimeTermination 
      AutoSize        =   -1  'True
      Caption         =   "sec"
      Height          =   195
      Left            =   2220
      TabIndex        =   7
      Top             =   530
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Offset:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Filter:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   180
      Width           =   435
   End
End
Attribute VB_Name = "frmFocusAction"
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

Private Const RegistryName = "FocusAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub chkSpecifyStarPos_Click()
    If Me.chkSpecifyStarPos.Value = vbChecked Then
        Me.fraFocusMaxStarPos.Enabled = True
    Else
        Me.fraFocusMaxStarPos.Enabled = False
    End If
End Sub

Private Sub cmdGetStarPos_Click()
    Dim myFocusMax As Object

    Set myFocusMax = CreateObject("FocusMax.FocusControl")
    
    Me.txtStarXPos.Text = CInt(myFocusMax.StarXCenter)
    Me.txtStarYPos.Text = CInt(myFocusMax.StarYCenter)
    
    If myFocusMax.StarXCenter = 0 And myFocusMax.StarYCenter = 0 Then
        MsgBox "You need to push the Select button in FocusMax prior to getting the star position.", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    Dim Counter As Integer
    
    Call MainMod.SetOnTopMode(Me)
    
    Me.cmbFilter.Clear
    
    If frmOptions.lstFocuserControl.ListIndex = FocusControl.CCDSoftAtFocus2 Then
        Me.lblExposureTime.Visible = False
        Me.lblExposureTimeTermination.Visible = False
        Me.txtExposureTime.Visible = False
    ElseIf frmOptions.lstFocuserControl.ListIndex = FocusControl.CCDSoftAtFocus Then
        Me.lblExposureTime.Visible = True
        Me.lblExposureTimeTermination.Visible = True
        Me.txtExposureTime.Visible = True
    Else
        Me.lblExposureTime.Visible = True
        Me.lblExposureTimeTermination.Visible = True
        Me.txtExposureTime.Visible = True
    End If
    
    If frmOptions.lstFocuserControl.ListIndex = FocusControl.FocusMax Then
        Me.chkSpecifyStarPos.Visible = True
        Me.fraFocusMaxStarPos.Visible = True
        Me.fraAverages.Visible = False
        Me.fraFocusMaxStarPos.ZOrder 1
    ElseIf frmOptions.lstFocuserControl.ListIndex = FocusControl.TheSkyXAtFocus3 Then
        Me.chkSpecifyStarPos.Visible = False
        Me.fraFocusMaxStarPos.Visible = False
        Me.fraAverages.Visible = True
        Me.fraAverages.ZOrder 1
    Else
        Me.chkSpecifyStarPos.Visible = False
        Me.fraFocusMaxStarPos.Visible = False
        Me.fraAverages.Visible = False
    End If
    
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

Private Sub GetSettings()
    If CInt(GetMySetting(RegistryName, "Filter", "0")) < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = CInt(GetMySetting(RegistryName, "Filter", "0"))
    End If
    
    Me.txtExposureTime.Text = GetMySetting(RegistryName, "@FocusExposureTime", Format(0.11, "0.00"))
    
    Me.chkSpecifyStarPos.Value = CInt(GetMySetting(RegistryName, "SpecifyStarPosition", "0"))
    Me.txtStarXPos.Text = GetMySetting(RegistryName, "StarXPosition", "0")
    Me.txtStarYPos.Text = GetMySetting(RegistryName, "StarYPosition", "0")
    
    Me.txtAverages.Text = GetMySetting(RegistryName, "FocusAverages", "1")
    Me.chkTempComp.Value = CInt(GetMySetting(RegistryName, "TempComp", "0"))
    
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    
    Cancel = False
    If Not Cancel Then Call txtStarXPos_Validate(Cancel)
    If Not Cancel Then Call txtStarYPos_Validate(Cancel)
    If Not Cancel Then Call txtAverages_Validate(Cancel)
    
    If Cancel Then
        Beep
        Exit Sub
    End If

    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    Call SaveMySetting(RegistryName, "@FocusExposureTime", Me.txtExposureTime.Text)
    
    Call SaveMySetting(RegistryName, "SpecifyStarPosition", Me.chkSpecifyStarPos.Value)
    Call SaveMySetting(RegistryName, "StarXPosition", Me.txtStarXPos.Text)
    Call SaveMySetting(RegistryName, "StarYPosition", Me.txtStarYPos.Text)
    
    Call SaveMySetting(RegistryName, "FocusAverages", Me.txtAverages.Text)
    Call SaveMySetting(RegistryName, "TempComp", Me.chkTempComp.Value)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Public Sub PutFormDataIntoClass(clsAction As FocusAction)
    clsAction.ImagerFilter = Me.cmbFilter.ListIndex
    clsAction.UseOffset = 0
    clsAction.Offset = 0
    clsAction.ExposureTime = CDbl(Me.txtExposureTime.Text)
    
    clsAction.SpecifyFocusStarPosition = CBool(Me.chkSpecifyStarPos.Value)
    clsAction.StarXPosition = CLng(Me.txtStarXPos.Text)
    clsAction.StarYPosition = CLng(Me.txtStarYPos.Text)
    
    clsAction.FocusAverages = CLng(Me.txtAverages.Text)
    clsAction.TempComp = CBool(Me.chkTempComp.Value)
    
    clsAction.FastReadout = False
End Sub

Public Sub GetFormDataFromClass(clsAction As FocusAction)
    If clsAction.ImagerFilter < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = clsAction.ImagerFilter
    End If
    
    Me.txtExposureTime.Text = clsAction.ExposureTime
    
    Me.chkSpecifyStarPos.Value = -clsAction.SpecifyFocusStarPosition
    Me.txtStarXPos.Text = clsAction.StarXPosition
    Me.txtStarYPos.Text = clsAction.StarYPosition
    
    Me.txtAverages.Text = clsAction.FocusAverages
    
    Me.chkTempComp.Value = -clsAction.TempComp

    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    Call SaveMySetting(RegistryName, "@FocusExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "SpecifyStarPosition", Me.chkSpecifyStarPos.Value)
    Call SaveMySetting(RegistryName, "StarXPosition", Me.txtStarXPos.Text)
    Call SaveMySetting(RegistryName, "StarYPosition", Me.txtStarYPos.Text)
    
    Call SaveMySetting(RegistryName, "FocusAverages", Me.txtAverages.Text)
    Call SaveMySetting(RegistryName, "TempComp", Me.chkTempComp.Value)
End Sub

Private Sub txtExposureTime_GotFocus()
    Me.txtExposureTime.SelStart = 0
    Me.txtExposureTime.SelLength = Len(Me.txtExposureTime.Text)
End Sub

Private Sub txtExposureTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtExposureTime.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtExposureTime.Text Then
        Beep
        Cancel = True
    Else
        Me.txtExposureTime.Text = Format(Me.txtExposureTime.Text, "0.000")
    End If
    On Error GoTo 0
End Sub

Private Sub txtStarXPos_GotFocus()
    Me.txtStarXPos.SelStart = 0
    Me.txtStarXPos.SelLength = Len(Me.txtStarXPos.Text)
End Sub

Private Sub txtStarXPos_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtStarXPos.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtStarXPos.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtStarYPos_GotFocus()
    Me.txtStarYPos.SelStart = 0
    Me.txtStarYPos.SelLength = Len(Me.txtStarYPos.Text)
End Sub

Private Sub txtStarYPos_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtStarYPos.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtStarYPos.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAverages_GotFocus()
    Me.txtAverages.SelStart = 0
    Me.txtAverages.SelLength = Len(Me.txtAverages.Text)
End Sub

Private Sub txtAverages_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtAverages.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtAverages.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

