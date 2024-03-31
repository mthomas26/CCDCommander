VERSION 5.00
Begin VB.Form frmMoveToRADec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move To RA & Dec Action"
   ClientHeight    =   3795
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5370
   HelpContextID   =   800
   Icon            =   "frmMoveToRADec.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGetFromTelescope 
      Caption         =   "Get RA/Dec from Current Telescope Position"
      Height          =   495
      Left            =   780
      TabIndex        =   9
      Top             =   1740
      Width           =   2055
   End
   Begin VB.OptionButton optEpoch 
      Caption         =   "J2000"
      Height          =   315
      Index           =   1
      Left            =   2580
      TabIndex        =   7
      Top             =   600
      Width           =   915
   End
   Begin VB.OptionButton optEpoch 
      Caption         =   "JNow"
      Height          =   315
      Index           =   0
      Left            =   2580
      TabIndex        =   6
      Top             =   300
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.CheckBox chkRecomputeCoordinates 
      Caption         =   "Recompute Object Coordinates before Slew"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton cmdFindObject 
      Caption         =   "Find Object in TheSky"
      Height          =   432
      Left            =   1440
      TabIndex        =   11
      Top             =   2820
      Width           =   2055
   End
   Begin VB.TextBox txtObjectName 
      Height          =   315
      Left            =   1740
      TabIndex        =   10
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdGetFromSky 
      Caption         =   "Get RA/Dec from TheSky"
      Height          =   495
      Left            =   780
      TabIndex        =   8
      Top             =   1140
      Width           =   2055
   End
   Begin VB.TextBox txtDecS 
      Height          =   312
      Left            =   1920
      TabIndex        =   5
      Text            =   "15"
      Top             =   660
      Width           =   312
   End
   Begin VB.TextBox txtDecM 
      Height          =   312
      Left            =   1260
      TabIndex        =   4
      Text            =   "23"
      Top             =   660
      Width           =   312
   End
   Begin VB.TextBox txtDecD 
      Height          =   312
      Left            =   660
      TabIndex        =   3
      Text            =   "5"
      Top             =   660
      Width           =   312
   End
   Begin VB.TextBox txtRAS 
      Height          =   312
      Left            =   1920
      TabIndex        =   2
      Text            =   "58"
      Top             =   180
      Width           =   312
   End
   Begin VB.TextBox txtRAM 
      Height          =   312
      Left            =   1260
      TabIndex        =   1
      Text            =   "10"
      Top             =   180
      Width           =   312
   End
   Begin VB.TextBox txtRAH 
      Height          =   312
      Left            =   660
      TabIndex        =   0
      Text            =   "23"
      Top             =   180
      Width           =   312
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3900
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Object Name:"
      Height          =   195
      Left            =   660
      TabIndex        =   23
      Top             =   2460
      Width           =   1035
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "s"
      Height          =   195
      Left            =   2280
      TabIndex        =   22
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   1620
      TabIndex        =   21
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "d"
      Height          =   195
      Left            =   1020
      TabIndex        =   20
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Dec:"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "s"
      Height          =   195
      Left            =   2280
      TabIndex        =   18
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   1620
      TabIndex        =   17
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "h"
      Height          =   195
      Left            =   1020
      TabIndex        =   16
      Top             =   240
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RA:"
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   240
      Width           =   270
   End
End
Attribute VB_Name = "frmMoveToRADec"
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

Private Const RegistryName = "MoveRADecAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub cmdFindObject_Click()
    Dim RA As Double
    Dim Dec As Double
    
    If Me.optEpoch(1).Value Then
        Call Planetarium.GetObjectRADec(Me.txtObjectName.Text, RA, Dec, True)
    Else
        Call Planetarium.GetObjectRADec(Me.txtObjectName.Text, RA, Dec, False)
    End If
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
    Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
    
    If Dec < 0 And Dec > -1 Then
        Me.txtDecD.Text = "-" & Fix(Dec + (0.5 / 3600))
    Else
        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
    End If
    
    Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
    Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
End Sub

Private Sub cmdGetFromSky_Click()
    Dim RA As Double
    Dim Dec As Double
    
    If Me.optEpoch(1).Value Then
        Me.txtObjectName.Text = Planetarium.GetSelectedObject(RA, Dec, True)
    Else
        Me.txtObjectName.Text = Planetarium.GetSelectedObject(RA, Dec, False)
    End If
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
    Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
    
    If Dec < 0 And Dec > -1 Then
        Me.txtDecD.Text = "-" & Fix(Dec + (0.5 / 3600))
    Else
        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
    End If
    Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
    Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
End Sub

Private Sub cmdGetFromTelescope_Click()
    Dim RA As Double
    Dim Dec As Double
    Dim myImageLink As New ImageLinkSyncAction
    Dim Result As VbMsgBoxResult
    
    Result = MsgBox("Perform Plate Solve to determine exact telescope position?", vbYesNoCancel + vbQuestion, "Get RA/Dec from Telescope")
    
    If Result = vbCancel Then
        Exit Sub
    ElseIf Result = vbNo Then
        Call Mount.GetTelescopeRADec(RA, Dec)
        
        Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
        Me.txtRAM.Text = Fix((RA - CDbl(Me.txtRAH.Text)) * 60# + (0.5 / 60))
        Me.txtRAS.Text = Fix(((((RA - CDbl(Me.txtRAH.Text)) * 60#) - CDbl(Me.txtRAM.Text)) * 60#) + 0.5)
        
        If Dec < 0 And Dec > -1 Then
            Me.txtDecD.Text = "-" & Fix(Dec + (0.5 / 3600))
        Else
            Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
        End If
        Me.txtDecM.Text = Fix(Abs(Dec - CDbl(Me.txtDecD.Text)) * 60# + (0.5 / 60))
        Me.txtDecS.Text = Fix((((Abs(Dec - CDbl(Me.txtDecD.Text)) * 60#) - CDbl(Me.txtDecM.Text)) * 60#) + 0.5)
        
        Me.txtObjectName.Text = ""
        
        Me.optEpoch(0).Value = True
        Me.chkRecomputeCoordinates.Value = vbUnchecked
    ElseIf Result = vbYes Then
        'Show Image Link window with somethings disabled
        Call Load(frmImageLinkSync)
        frmImageLinkSync.fraSync.Visible = False
        frmImageLinkSync.chkSlew.Visible = False
        frmImageLinkSync.chkAbort.Visible = False
        frmImageLinkSync.fraRetry.Visible = False
        frmImageLinkSync.chkUseGlobalImageSaveLocation.Visible = False
        frmImageLinkSync.lblImageSaveLocation.Visible = False
        frmImageLinkSync.txtSaveTo.Visible = False
        frmImageLinkSync.cmdOpen.Visible = False
        frmImageLinkSync.chkAutosave.Visible = False
        frmImageLinkSync.lblFilenamePrefix.Visible = False
        frmImageLinkSync.txtFileNamePrefix.Visible = False
        frmImageLinkSync.cmdAutoFileNameBuilder.Visible = False
        frmImageLinkSync.Caption = "Get RA/Dec from Telescope Plate Solve Settings"
        
        Call frmImageLinkSync.Show(vbModal, Me)
        If frmImageLinkSync.Tag = "False" Then
            Exit Sub
        End If
        
        Me.MousePointer = vbHourglass
        DoEvents
        Me.Enabled = False
        
        Call frmImageLinkSync.PutFormDataIntoClass(myImageLink)
        
        Call Unload(frmImageLinkSync)
        
        myImageLink.AbortListOnFailure = False
        myImageLink.ArcminutesToSlew = 0
        myImageLink.AutosaveExposure = vbUnchecked
        myImageLink.DelayTime = 0
        myImageLink.FileNamePrefix = ""
        myImageLink.FileSavePath = ""
        myImageLink.RetryPlateSolveOnFailure = False
        myImageLink.SkipIfRetrySucceeds = False
        myImageLink.SlewMountForRetry = False
        myImageLink.SyncMode = NoSync
        myImageLink.UseGlobalSaveToLocation = vbChecked
        
        Call Load(frmMoveToPlateSolve)
        Set frmMoveToPlateSolve.myImageLink = myImageLink
        If Me.optEpoch(0).Value Then
            frmMoveToPlateSolve.Precess = True
        Else
            frmMoveToPlateSolve.Precess = False
        End If
        Call frmMoveToPlateSolve.Show(vbModal, Me)
        
        If frmMoveToPlateSolve.Tag = "1" Then
            Me.txtObjectName.Text = ""
            Me.chkRecomputeCoordinates.Value = vbUnchecked
        End If
        
        Unload frmMoveToPlateSolve
    
        Me.MousePointer = vbNormal
        Me.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    If frmOptions.lstPlanetarium.ListIndex = 0 Then
        Me.cmdFindObject.Enabled = False
        Me.cmdGetFromSky.Enabled = False
        Me.chkRecomputeCoordinates.Enabled = False
    Else
        Me.cmdFindObject.Enabled = True
        Me.cmdGetFromSky.Enabled = True
        Me.chkRecomputeCoordinates.Enabled = True
    End If
        
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
    Me.txtRAH.Text = GetMySetting(RegistryName, "RAH", "0")
    Me.txtRAM.Text = GetMySetting(RegistryName, "RAM", "0")
    Me.txtRAS.Text = GetMySetting(RegistryName, "RAS", "0")
    Me.txtDecD.Text = GetMySetting(RegistryName, "DecD", "0")
    Me.txtDecM.Text = GetMySetting(RegistryName, "DecM", "0")
    Me.txtDecS.Text = GetMySetting(RegistryName, "DecS", "0")
    Me.txtObjectName.Text = GetMySetting(RegistryName, "ObjectName", "")
    Me.chkRecomputeCoordinates.Value = CInt(GetMySetting(RegistryName, "RecomputeCoordinates", "0"))
    
    If GetMySetting(RegistryName, "Epoch", "0") = "0" Then
        Me.optEpoch(0).Value = True
    Else
        Me.optEpoch(1).Value = True
    End If
End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    If Not Cancel Then Call txtRAH_Validate(Cancel)
    If Not Cancel Then Call txtRAM_Validate(Cancel)
    If Not Cancel Then Call txtRAS_Validate(Cancel)
    If Not Cancel Then Call txtDecD_Validate(Cancel)
    If Not Cancel Then Call txtDecM_Validate(Cancel)
    If Not Cancel Then Call txtDecS_Validate(Cancel)
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    If Me.chkRecomputeCoordinates.Value = 1 Then
        'validate object name
        Call cmdFindObject_Click
        
        If Me.txtRAH.Text = "0" And Me.txtRAM.Text = "0" And Me.txtRAS.Text = "0" And _
            Me.txtDecD.Text = "0" And Me.txtDecM.Text = "0" And Me.txtDecS.Text = "0" Then
            
            'object was not found
            If MsgBox("Object Name not found." & vbCrLf & "Continue?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    Call SaveMySetting(RegistryName, "RAH", Me.txtRAH.Text)
    Call SaveMySetting(RegistryName, "RAM", Me.txtRAM.Text)
    Call SaveMySetting(RegistryName, "RAS", Me.txtRAS.Text)
    Call SaveMySetting(RegistryName, "DecD", Me.txtDecD.Text)
    Call SaveMySetting(RegistryName, "DecM", Me.txtDecM.Text)
    Call SaveMySetting(RegistryName, "DecS", Me.txtDecS.Text)
    Call SaveMySetting(RegistryName, "ObjectName", Me.txtObjectName.Text)
    Call SaveMySetting(RegistryName, "RecomputeCoordinates", Me.chkRecomputeCoordinates.Value)
    
    If Me.optEpoch(0).Value Then
        Call SaveMySetting(RegistryName, "Epoch", "0")
    Else
        Call SaveMySetting(RegistryName, "Epoch", "1")
    End If
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub optEpoch_Click(Index As Integer)
    If Index = 0 Then
        Me.optEpoch(0).Value = True
        Me.optEpoch(1).Value = False
        Me.chkRecomputeCoordinates.Enabled = True
    Else
        Me.optEpoch(0).Value = False
        Me.optEpoch(1).Value = True
        Me.chkRecomputeCoordinates.Enabled = False
        Me.chkRecomputeCoordinates.Value = vbUnchecked
    End If
End Sub

Private Sub txtRAH_GotFocus()
    Me.txtRAH.SelStart = 0
    Me.txtRAH.SelLength = Len(Me.txtRAH.Text)
End Sub

Private Sub txtRAH_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtRAH.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 24 Or Test <> Me.txtRAH.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtRAM_GotFocus()
    Me.txtRAM.SelStart = 0
    Me.txtRAM.SelLength = Len(Me.txtRAM.Text)
End Sub

Private Sub txtRAM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtRAM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtRAM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtRAS_GotFocus()
    Me.txtRAS.SelStart = 0
    Me.txtRAS.SelLength = Len(Me.txtRAS.Text)
End Sub

Private Sub txtRAS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtRAS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtRAS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDecD_GotFocus()
    Me.txtDecD.SelStart = 0
    Me.txtDecD.SelLength = Len(Me.txtDecD.Text)
End Sub

Private Sub txtDecD_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDecD.Text)
    If Err.Number <> 0 Or Test < -90 Or Test > 90 Or Test <> Me.txtDecD.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDecM_GotFocus()
    Me.txtDecM.SelStart = 0
    Me.txtDecM.SelLength = Len(Me.txtDecM.Text)
End Sub

Private Sub txtDecM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDecM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtDecM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDecS_GotFocus()
    Me.txtDecS.SelStart = 0
    Me.txtDecS.SelLength = Len(Me.txtDecS.Text)
End Sub

Private Sub txtDecS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDecS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtDecS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Public Sub PutFormDataIntoClass(clsAction As MoveRADecAction)
    clsAction.RA = CDbl(Me.txtRAH.Text) + (CDbl(Me.txtRAM.Text) / 60) + (CDbl(Me.txtRAS.Text) / 3600)
    clsAction.Dec = Abs(CDbl(Me.txtDecD.Text)) + (CDbl(Me.txtDecM.Text) / 60) + (CDbl(Me.txtDecS.Text) / 3600)
    If InStr(Me.txtDecD.Text, "-") <> 0 Then
        clsAction.Dec = -clsAction.Dec
    End If
    
    clsAction.Name = Me.txtObjectName.Text
    
    clsAction.RecomputeObjectCoordinates = Me.chkRecomputeCoordinates.Value
    
    If Me.optEpoch(0).Value Then
        clsAction.Epoch = JNow
    Else
        clsAction.Epoch = J2000
    End If
End Sub

Public Sub GetFormDataFromClass(clsAction As MoveRADecAction)
    Dim RA As Double
    Dim Dec As Double
    
    RA = clsAction.RA
    Dec = clsAction.Dec
    
    Me.txtRAH.Text = Fix(RA + (0.5 / 3600))
    RA = (RA - Fix(RA + (0.5 / 3600))) * 60
    Me.txtRAM.Text = Fix(RA + (0.5 / 60))
    RA = (RA - Fix(RA + (0.5 / 60))) * 60
    Me.txtRAS.Text = Fix(RA + 0.5)

    If Dec < 0 And Dec > -1 Then
        Dec = -Dec
        Me.txtDecD.Text = "-0"
    ElseIf Dec < 0 Then
        Dec = -Dec
        Me.txtDecD.Text = -Fix(Dec + (0.5 / 3600))
    Else
        Me.txtDecD.Text = Fix(Dec + (0.5 / 3600))
    End If
    Dec = (Dec - Fix(Dec + (0.5 / 3600))) * 60
    Me.txtDecM.Text = Fix(Dec + (0.5 / 60))
    Dec = (Dec - Fix(Dec + (0.5 / 60))) * 60
    Me.txtDecS.Text = Fix(Dec + 0.5)
    
    Me.txtObjectName.Text = clsAction.Name
    
    Me.chkRecomputeCoordinates.Value = clsAction.RecomputeObjectCoordinates
    
    If clsAction.Epoch = JNow Then
        Me.optEpoch(0).Value = True
    Else
        Me.optEpoch(1).Value = True
    End If

    Call SaveMySetting(RegistryName, "RAH", Me.txtRAH.Text)
    Call SaveMySetting(RegistryName, "RAM", Me.txtRAM.Text)
    Call SaveMySetting(RegistryName, "RAS", Me.txtRAS.Text)
    Call SaveMySetting(RegistryName, "DecD", Me.txtDecD.Text)
    Call SaveMySetting(RegistryName, "DecM", Me.txtDecM.Text)
    Call SaveMySetting(RegistryName, "DecS", Me.txtDecS.Text)
    Call SaveMySetting(RegistryName, "ObjectName", Me.txtObjectName.Text)
    Call SaveMySetting(RegistryName, "RecomputeCoordinates", Me.chkRecomputeCoordinates.Value)
    
    If Me.optEpoch(0).Value Then
        Call SaveMySetting(RegistryName, "Epoch", "0")
    Else
        Call SaveMySetting(RegistryName, "Epoch", "1")
    End If
End Sub

