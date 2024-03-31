VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Target List"
   ClientHeight    =   3570
   ClientLeft      =   2760
   ClientTop       =   4050
   ClientWidth     =   7515
   HelpContextID   =   10000
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6240
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":0976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImport.frx":0B9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImportToSubList 
      Caption         =   "Import as Sub-List"
      Height          =   375
      Left            =   3900
      TabIndex        =   8
      Top             =   3120
      Width           =   1605
   End
   Begin VB.CommandButton cmdImportToMainList 
      Caption         =   "Import to Main List"
      Height          =   375
      Left            =   1937
      TabIndex        =   7
      Top             =   3120
      Width           =   1605
   End
   Begin VB.CommandButton cmdImportToFile 
      Caption         =   "Import To File"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   3120
      Width           =   1605
   End
   Begin VB.CommandButton cmdOpenTargetList 
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
      Left            =   7200
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmImport.frx":0CB0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5820
      TabIndex        =   0
      Top             =   3120
      Width           =   1605
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   6840
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstAction 
      Height          =   1635
      Left            =   60
      TabIndex        =   4
      Top             =   1380
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2884
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   12347
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy Action"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut Action"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste Action"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Delete"
            Object.ToolTipText     =   "Delete Action"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Move Up"
            Object.ToolTipText     =   "Move Action Up"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Move Down"
            Object.ToolTipText     =   "Move Action Down"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Actions for Each Target:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Target List:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   795
   End
   Begin VB.Label lblFileName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\Program Files\CCD Commander\Target Lists\Sample.txt"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   7095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close"
         Index           =   4
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "C&ut"
         Index           =   1
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "De&lete"
         Index           =   4
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Move Up"
         Index           =   6
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Move &Down"
         Index           =   7
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuActionItems 
         Caption         =   "Move To RA && Dec"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Rotate Camera"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Plate Solve"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Focus"
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Take Images"
         Index           =   5
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Wait for Altitude"
         Index           =   6
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Wait for Time"
         Index           =   7
         Shortcut        =   {F8}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Skip Ahead at Time"
         Index           =   8
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Skip Ahead at Altitude"
         Index           =   9
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Skip Ahead at Hour Angle"
         Index           =   10
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Automatic Flats"
         Index           =   11
         Shortcut        =   {F11}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Temperature Control"
         Index           =   12
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Dome Control"
         Index           =   13
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Cloud Monitor Control"
         Index           =   14
         Shortcut        =   +{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Run External Script"
         Index           =   15
         Shortcut        =   +{F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Run Sub-Action List"
         Index           =   16
         Shortcut        =   +{F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Park Mount"
         Index           =   17
         Shortcut        =   +{F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmImport"
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

Private Const MoveToRADecIndex = 1
Private Const RotatorIndex = 2
Private Const SyncIndex = 3
Private Const FocusMaxIndex = 4
Private Const ImagerIndex = 5
Private Const WaitForAlt = 6
Private Const WaitForTime = 7
Private Const SkipAheadAtTimeIndex = 8
Private Const SkipAheadAtAltIndex = 9
Private Const SkipAheadAtHAIndex = 10
Private Const AutoFlatIndex = 11
Private Const IntelligentTempIndex = 12
Private Const DomeActionIndex = 13
Private Const CloudMonitorActionIndex = 14
Private Const RunScript = 15
Private Const RunActionList = 16
Private Const ParkMountIndex = 17
    
Private ImportActionList As Collection

Private Modified As Boolean

Private Enum ToolBarButtons
    NewFileBtn = 1
    OpenFileBtn = 2
    SaveFileBtn = 3
    CopyBtn = 5
    CutBtn = 6
    PasteBtn = 7
    DeleteBtn = 9
    MoveUpBtn = 11
    MoveDownBtn = 12
End Enum

Private Enum FileMenuItems
    NewFileMenu = 0
    OpenFileMenu = 1
    SaveFileMenu = 2
    CloseFileMenu = 4
End Enum

Private Enum EditMenuItems
    CopyMenu = 0
    CutMenu = 1
    PasteMenu = 2
    DeleteMenu = 4
    MoveUpMenu = 6
    MoveDownMenu = 7
End Enum

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
    ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Private Sub cmdClose_Click()
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmdImportToFile_Click()
    Dim ImportedActionCollection As New Collection
    
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    On Error Resume Next
    Call MkDir(App.Path & "\Actions")
    On Error GoTo 0
    Me.ComDlg.DialogTitle = "Save Action List"
    Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
    Me.ComDlg.FilterIndex = 1
    Me.ComDlg.InitDir = App.Path & "\Actions\"
    Me.ComDlg.FileName = Format(Now, "yymmdd") & ".act"
    Me.ComDlg.CancelError = True
    Me.ComDlg.flags = cdlOFNOverwritePrompt
    On Error Resume Next
    Me.ComDlg.ShowSave
    If Err.Number = 0 Then
        On Error GoTo 0
        
        Me.MousePointer = vbHourglass
        Call Me.ImportToList(ImportedActionCollection)
        
        Call SaveAction(Me.ComDlg.FileName, ImportedActionCollection)
        
        Me.MousePointer = vbNormal
    Else
        On Error GoTo 0
    End If
End Sub

Private Sub cmdImportToMainList_Click()
    Dim clsAction As Object
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If

    MainMod.ClearAll
    
    Me.MousePointer = vbHourglass
    
    Call Me.ImportToList(frmMain.ActionCollections(1))

    For Each clsAction In frmMain.ActionCollections(1)
        DoEvents
        frmMain.lstAction(1).ListItems.Add , , clsAction.BuildActionListString
        If clsAction.Selected Then
            frmMain.lstAction(1).ListItems(frmMain.lstAction(1).ListItems.Count).Checked = True
        End If
    Next clsAction
    
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdImportToSubList_Click()
    Dim clsAction As Object
    
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If

    Call frmMain.AddAction(RunActionList)

    Me.MousePointer = vbHourglass
    
    Call Me.ImportToList(frmMain.ActionCollections(frmMain.TabStrip.SelectedItem.Index))

    For Each clsAction In frmMain.ActionCollections(frmMain.TabStrip.SelectedItem.Index)
        DoEvents
        frmMain.lstAction(frmMain.TabStrip.SelectedItem.Index).ListItems.Add , , clsAction.BuildActionListString
        If clsAction.Selected Then
            frmMain.lstAction(frmMain.TabStrip.SelectedItem.Index).ListItems(frmMain.lstAction(frmMain.TabStrip.SelectedItem.Index).ListItems.Count).Checked = True
        End If
    Next clsAction
    
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdOpenTargetList_Click()
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    Me.ComDlg.FileName = ""
    Me.ComDlg.Filter = "Target List (*.txt)|*.txt"
    Me.ComDlg.FilterIndex = 1
    Me.ComDlg.CancelError = True
    Me.ComDlg.InitDir = Me.lblFileName.Caption
    Me.ComDlg.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.ComDlg.ShowOpen
    If Err.Number = 0 Then
        Me.lblFileName.Caption = Me.ComDlg.FileName
    Else
        Me.lblFileName.Caption = ""
    End If
    On Error GoTo 0
    
    Call SaveMySetting("Import Utility", "Import File", Me.lblFileName.Caption)
End Sub

Private Sub Form_Load()
    Dim actFileNum As Integer
    Dim ByteData() As Byte
    Dim NumberOfBytes As Long
    Dim clsAction As MoveRADecAction
    
    Call MainMod.SetOnTopMode(Me)
    
    Set ImportActionList = New Collection
    
    On Error Resume Next
    Call MkDir(App.Path & "\ImportActionTemplates")
    On Error GoTo 0
    actFileNum = FreeFile
    On Error GoTo FormLoadError
    Open App.Path & "\ImportActionTemplates\DefaultTemplate.alt" For Binary Access Read As #actFileNum
    
    NumberOfBytes = LOF(actFileNum)
    
    ReDim ByteData(0 To NumberOfBytes - 1)
    Get #actFileNum, , ByteData()
    Close actFileNum
    
    Call LoadActionLists(ImportActionList, ByteData, 0, UBound(ByteData), , Me.lstAction)

    Me.lblFileName.Caption = GetMySetting("Import Utility", "Import File", "C:\")
    
    On Error GoTo 0
    Exit Sub
    
FormLoadError:
    Close actFileNum
    
    Set ImportActionList = New Collection
    Me.lstAction.ListItems.Clear
    
    Set clsAction = New MoveRADecAction
    clsAction.Name = ""
    clsAction.RA = 0
    clsAction.Dec = 0
    clsAction.Epoch = J2000
    clsAction.Selected = True
    
    ImportActionList.Add clsAction
    Me.lstAction.ListItems.Add , , clsAction.BuildActionListString()
    On Error GoTo 0
End Sub

Private Sub lstAction_DblClick()
    Dim ActionInfo As Object
    
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    If Me.lstAction.SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first.", vbExclamation
        Exit Sub
    End If
    
    Set ActionInfo = ImportActionList.Item(Me.lstAction.SelectedItem.Index)
    
    If TypeName(ActionInfo) = "ImagerAction" Then
        Call frmCameraAction.GetFormDataFromClass(ActionInfo)
        Call SetupActionFormForImportList(ImagerIndex)
        Call frmCameraAction.Show(1, Me)
        Call frmCameraAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmCameraAction.Tag) Then Modified = True
        Unload frmCameraAction
    ElseIf TypeName(ActionInfo) = "MoveRADecAction" Then
        Call frmMoveToRADec.GetFormDataFromClass(ActionInfo)
        Call SetupActionFormForImportList(MoveToRADecIndex)
        Call frmMoveToRADec.Show(1, Me)
        Call frmMoveToRADec.PutFormDataIntoClass(ActionInfo)
        If CBool(frmMoveToRADec.Tag) Then Modified = True
        Unload frmMoveToRADec
    ElseIf TypeName(ActionInfo) = "FocusAction" Then
        Call frmFocusAction.GetFormDataFromClass(ActionInfo)
        Call frmFocusAction.Show(1, Me)
        Call frmFocusAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmFocusAction.Tag) Then Modified = True
        Unload frmFocusAction
    ElseIf TypeName(ActionInfo) = "RotatorAction" Then
        Call frmRotator.GetFormDataFromClass(ActionInfo)
        Call frmRotator.Show(1, Me)
        Call frmRotator.PutFormDataIntoClass(ActionInfo)
        If CBool(frmRotator.Tag) Then Modified = True
        Unload frmRotator
    ElseIf TypeName(ActionInfo) = "IntelligentTempAction" Then
        Call frmTempControl.GetFormDataFromClass(ActionInfo)
        Call frmTempControl.Show(1, Me)
        Call frmTempControl.PutFormDataIntoClass(ActionInfo)
        If CBool(frmTempControl.Tag) Then Modified = True
        Unload frmTempControl
    ElseIf TypeName(ActionInfo) = "ImageLinkSyncAction" Then
        Call frmImageLinkSync.GetFormDataFromClass(ActionInfo)
        Call frmImageLinkSync.Show(1, Me)
        Call frmImageLinkSync.PutFormDataIntoClass(ActionInfo)
        If CBool(frmImageLinkSync.Tag) Then Modified = True
        Unload frmImageLinkSync
    ElseIf TypeName(ActionInfo) = "WaitForAltAction" Then
        Call frmWaitForAlt.GetFormDataFromClass(ActionInfo)
        Call SetupActionFormForImportList(WaitForAlt)
        Call frmWaitForAlt.Show(1, Me)
        Call frmWaitForAlt.PutFormDataIntoClass(ActionInfo)
        If CBool(frmWaitForAlt.Tag) Then Modified = True
        Unload frmWaitForAlt
    ElseIf TypeName(ActionInfo) = "WaitForTimeAction" Then
        Call frmWaitForTime.GetFormDataFromClass(ActionInfo)
        Call frmWaitForTime.Show(1, Me)
        Call frmWaitForTime.PutFormDataIntoClass(ActionInfo)
        If CBool(frmWaitForTime.Tag) Then Modified = True
        Unload frmWaitForTime
    ElseIf TypeName(ActionInfo) = "ParkMountAction" Then
        Call frmPark.GetFormDataFromClass(ActionInfo)
        Call frmPark.Show(1, Me)
        Call frmPark.PutFormDataIntoClass(ActionInfo)
        If CBool(frmPark.Tag) Then Modified = True
        Unload frmPark
    ElseIf TypeName(ActionInfo) = "RunScriptAction" Then
        Call frmRunScript.GetFormDataFromClass(ActionInfo)
        Call frmRunScript.Show(1, Me)
        Call frmRunScript.PutFormDataIntoClass(ActionInfo)
        If CBool(frmRunScript.Tag) Then Modified = True
        Unload frmRunScript
    ElseIf TypeName(ActionInfo) = "RunActionList" Then
'        If RunningAction And Index < RunningActionListLevel Then
'            If ActionInfo Is SubActionInfo(Index + 1) Then
'                'clicking on the same action that is already open
'                Set Me.TabStrip.SelectedItem = Me.TabStrip.Tabs(Index + 1)
'            Else
'                MsgBox "You must wait for the running sub-action to complete before editing this sub-action.", vbInformation
'            End If
'        Else
'            Call RemoveSubActionTabsAfter
'
'            'Copy link to the ActionInfo object into the array so I can access the object when the user makes changes on the window
'            Call SubActionInfo.Add(ActionInfo, , , Index)
'
'            'Copy link to the Collection into the collection array
'            Call ImportActionList.Add(SubActionInfo(Index + 1).ActionCollection, , , Index)
'
'            Call UpperActionIndexes.Add(Me.lstAction(Index).SelectedItem.Index, , , Index)
'
'            Call AllowSubActionEdit.Add(AllowEdits, , , Index)
'
'            'Create new tab for the sub-action
'            Call AddStuffForSubAction(ActionInfo)
'        End If
    ElseIf TypeName(ActionInfo) = "AutoFlatAction" Then
        Call frmAutoFlat.GetFormDataFromClass(ActionInfo)
        Call frmAutoFlat.Show(1, Me)
        Call frmAutoFlat.PutFormDataIntoClass(ActionInfo)
        If CBool(frmAutoFlat.Tag) Then Modified = True
        Unload frmAutoFlat
    ElseIf TypeName(ActionInfo) = "SkipAheadAtTimeAction" Then
        Call frmSkipAtTime.GetFormDataFromClass(ActionInfo)
        Call frmSkipAtTime.Show(1, Me)
        Call frmSkipAtTime.PutFormDataIntoClass(ActionInfo)
        If CBool(frmSkipAtTime.Tag) Then Modified = True
        Unload frmSkipAtTime
    ElseIf TypeName(ActionInfo) = "SkipAheadAtAltAction" Then
        Call frmSkipAtAlt.GetFormDataFromClass(ActionInfo)
        Call frmSkipAtAlt.Show(1, Me)
        Call frmSkipAtAlt.PutFormDataIntoClass(ActionInfo)
        If CBool(frmSkipAtAlt.Tag) Then Modified = True
        Unload frmSkipAtAlt
    ElseIf TypeName(ActionInfo) = "SkipAheadAtHAAction" Then
        Call frmSkipAtHA.GetFormDataFromClass(ActionInfo)
        Call frmSkipAtHA.Show(1, Me)
        Call frmSkipAtHA.PutFormDataIntoClass(ActionInfo)
        If CBool(frmSkipAtHA.Tag) Then Modified = True
        Unload frmSkipAtHA
    ElseIf TypeName(ActionInfo) = "DomeAction" Then
        Call frmDomeAction.GetFormDataFromClass(ActionInfo)
        Call frmDomeAction.Show(1, Me)
        Call frmDomeAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmDomeAction.Tag) Then Modified = True
        Unload frmDomeAction
    ElseIf TypeName(ActionInfo) = "CloudMonitorAction" Then
        Call frmCloudMonitorAction.GetFormDataFromClass(ActionInfo)
        Call frmCloudMonitorAction.Show(1, Me)
        Call frmCloudMonitorAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmCloudMonitorAction.Tag) Then Modified = True
        Unload frmCloudMonitorAction
    End If
    
    Call ImportActionList.Remove(Me.lstAction.SelectedItem.Index)
    If ImportActionList.Count = 0 Then
        Call ImportActionList.Add(ActionInfo)
    ElseIf Me.lstAction.SelectedItem.Index > 1 Then
        Call ImportActionList.Add(ActionInfo, , , Me.lstAction.SelectedItem.Index - 1)
    Else
        Call ImportActionList.Add(ActionInfo, , Me.lstAction.SelectedItem.Index)
    End If
    Me.lstAction.ListItems(Me.lstAction.SelectedItem.Index).Text = ActionInfo.BuildActionListString()
    
End Sub

Private Sub lstAction_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    If KeyCode = vbKeyDelete Then
        Call DeleteActions
    End If
End Sub

Private Sub lstAction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuEdit
    End If
End Sub

Private Sub SetupActionFormForImportList(WhatAction As Integer)
    Dim myObject As Object
    
    Select Case WhatAction
        Case ImagerIndex
'            frmCameraAction.chkAutosave.Value = vbChecked
'            frmCameraAction.chkAutosave.Enabled = False
'            frmCameraAction.txtFileNamePrefix = ""
'            frmCameraAction.txtFileNamePrefix.Enabled = False
        Case MoveToRADecIndex
            For Each myObject In frmMoveToRADec.Controls
                If myObject.Name <> "OKButton" And myObject.Name <> "CancelButton" And myObject.Name <> "chkRecomputeCoordinates" And myObject.Name <> "optEpoch" Then
                    myObject.Enabled = False
                End If
            Next myObject
            
            frmMoveToRADec.txtRAH = "0"
            frmMoveToRADec.txtRAM = "0"
            frmMoveToRADec.txtRAS = "0"
            frmMoveToRADec.txtDecD = "0"
            frmMoveToRADec.txtDecM = "0"
            frmMoveToRADec.txtDecS = "0"
            frmMoveToRADec.txtObjectName = ""
        Case WaitForAlt
            frmWaitForAlt.fraObjectCoordinates.Enabled = False
            frmWaitForAlt.fraObjectName.Enabled = False
            frmWaitForAlt.txtRAH = "0"
            frmWaitForAlt.txtRAM = "0"
            frmWaitForAlt.txtRAS = "0"
            frmWaitForAlt.txtDecD = "0"
            frmWaitForAlt.txtDecM = "0"
            frmWaitForAlt.txtDecS = "0"
            frmWaitForAlt.txtObjectName = ""
        Case SkipAheadAtAltIndex
            frmSkipAtAlt.fraObjectCoordinates.Enabled = False
            frmSkipAtAlt.fraObjectName.Enabled = False
            frmSkipAtAlt.txtRAH = "0"
            frmSkipAtAlt.txtRAM = "0"
            frmSkipAtAlt.txtRAS = "0"
            frmSkipAtAlt.txtDecD = "0"
            frmSkipAtAlt.txtDecM = "0"
            frmSkipAtAlt.txtDecS = "0"
            frmSkipAtAlt.txtObjectName = ""
        Case SkipAheadAtHAIndex
            frmSkipAtHA.fraObjectCoordinates.Enabled = False
            frmSkipAtHA.fraObjectName.Enabled = False
            frmSkipAtHA.txtRAH = "0"
            frmSkipAtHA.txtRAM = "0"
            frmSkipAtHA.txtRAS = "0"
            frmSkipAtHA.txtObjectName = ""
    End Select
End Sub


Private Sub mnuActionItems_Click(Index As Integer)
    Dim ActionInfo As Object
    Dim SelectedActionIndex As Long
    Dim ActionIndex As Long
        
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    'first need to figure out where to put the pasted actions
    'find the last selected action
    SelectedActionIndex = 0
    For ActionIndex = 1 To Me.lstAction.ListItems.Count
        If Me.lstAction.ListItems(ActionIndex).Selected Then
            SelectedActionIndex = ActionIndex
            Me.lstAction.ListItems(ActionIndex).Selected = False
        End If
    Next ActionIndex
    
    If SelectedActionIndex = 0 Then
        'nothing is selected - put it at the end
        SelectedActionIndex = Me.lstAction.ListItems.Count
    End If
            
    Call SetupActionFormForImportList(Index)
        
    If Index = ImagerIndex Then
        Call frmCameraAction.Show(1, Me)
        
        If CBool(frmCameraAction.Tag) Then
            Set ActionInfo = New ImagerAction
            Call frmCameraAction.PutFormDataIntoClass(ActionInfo)
            
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmCameraAction
    ElseIf Index = MoveToRADecIndex Then
        Call frmMoveToRADec.Show(1, Me)
        
        If CBool(frmMoveToRADec.Tag) Then
            Set ActionInfo = New MoveRADecAction
            Call frmMoveToRADec.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmMoveToRADec
    ElseIf Index = SyncIndex Then
        Call frmImageLinkSync.Show(1, Me)
        
        If CBool(frmImageLinkSync.Tag) Then
            Set ActionInfo = New ImageLinkSyncAction
            Call frmImageLinkSync.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmImageLinkSync
    ElseIf Index = FocusMaxIndex Then
        Call frmFocusAction.Show(1, Me)
        
        If CBool(frmFocusAction.Tag) Then
            Set ActionInfo = New FocusAction
            Call frmFocusAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmFocusAction
    ElseIf Index = RotatorIndex Then
        Call frmRotator.Show(1, Me)
        
        If CBool(frmRotator.Tag) Then
            Set ActionInfo = New RotatorAction
            Call frmRotator.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmRotator
    ElseIf Index = IntelligentTempIndex Then
        Call frmTempControl.Show(1, Me)
        
        If CBool(frmTempControl.Tag) Then
            Set ActionInfo = New IntelligentTempAction
            Call frmTempControl.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmTempControl
    ElseIf Index = WaitForAlt Then
        Call frmWaitForAlt.Show(1, Me)
        
        If CBool(frmWaitForAlt.Tag) Then
            Set ActionInfo = New WaitForAltAction
            Call frmWaitForAlt.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmWaitForAlt
    ElseIf Index = WaitForTime Then
        Call frmWaitForTime.Show(1, Me)
        
        If CBool(frmWaitForTime.Tag) Then
            Set ActionInfo = New WaitForTimeAction
            Call frmWaitForTime.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmWaitForTime
    ElseIf Index = ParkMountIndex Then
        Call frmPark.Show(1, Me)
        
        If CBool(frmPark.Tag) Then
            Set ActionInfo = New ParkMountAction
            Call frmPark.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmPark
    ElseIf Index = RunScript Then
        Call frmRunScript.Show(1, Me)
        
        If CBool(frmRunScript.Tag) Then
            Set ActionInfo = New RunScriptAction
            Call frmRunScript.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmRunScript
'    ElseIf Index = RunActionList Then
'        Set ActionInfo = New RunActionList
'
'        Call RemoveSubActionTabsAfter
'
'        'Copy link to the ActionInfo object into the array so I can access the object when the user makes changes on the window
'        Call SubActionInfo.Add(ActionInfo, , , Me.TabStrip.SelectedItem.Index)
'
'        'Create collection for Action object
'        Set SubActionInfo(Me.TabStrip.SelectedItem.Index + 1).ActionCollection = New Collection
'
'        'Copy link to the Collection into the collection array
'        Call ImportActionList.Add(SubActionInfo(Me.TabStrip.SelectedItem.Index + 1).ActionCollection, , , Me.TabStrip.SelectedItem.Index)
'
'        ActionInfo.Selected = True
'
'        'Create entry in the current collection and list
'        ImportActionList.Add ActionInfo
'        Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
'        Me.lstAction.ListItems(Me.lstAction.ListItems.Count).Checked = True
'
'        Call UpperActionIndexes.Add(Me.lstAction.ListItems.Count, , , Me.TabStrip.SelectedItem.Index)
'
'        Call AllowSubActionEdit.Add(True, , , Me.TabStrip.SelectedItem.Index)
'
'        If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
'            MsgBox "Error!  Counts don't match!"
'            End
'        End If
'        Set ActionInfo = Nothing
'
'        'Create new tab for the sub-action
'        Call AddStuffForSubAction(Nothing)
    ElseIf Index = AutoFlatIndex Then
        Call frmAutoFlat.Show(1, Me)
        
        If CBool(frmAutoFlat.Tag) Then
            Set ActionInfo = New AutoFlatAction
            Call frmAutoFlat.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmAutoFlat
    ElseIf Index = SkipAheadAtTimeIndex Then
        Call frmSkipAtTime.Show(1, Me)
        
        If CBool(frmSkipAtTime.Tag) Then
            Set ActionInfo = New SkipAheadAtTimeAction
            Call frmSkipAtTime.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmSkipAtTime
    ElseIf Index = SkipAheadAtAltIndex Then
        Call frmSkipAtAlt.Show(1, Me)
        
        If CBool(frmSkipAtAlt.Tag) Then
            Set ActionInfo = New SkipAheadAtAltAction
            Call frmSkipAtAlt.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmSkipAtAlt
    ElseIf Index = SkipAheadAtHAIndex Then
        Call frmSkipAtHA.Show(1, Me)
        
        If CBool(frmSkipAtHA.Tag) Then
            Set ActionInfo = New SkipAheadAtHAAction
            Call frmSkipAtHA.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmSkipAtHA
    ElseIf Index = DomeActionIndex Then
        Call frmDomeAction.Show(1, Me)
        
        If CBool(frmDomeAction.Tag) Then
            Set ActionInfo = New DomeAction
            Call frmDomeAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmDomeAction
    ElseIf Index = CloudMonitorActionIndex Then
        Call frmCloudMonitorAction.Show(1, Me)
        
        If CBool(frmCloudMonitorAction.Tag) Then
            Set ActionInfo = New CloudMonitorAction
            Call frmCloudMonitorAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ImportActionList.Add ActionInfo
                Call Me.lstAction.ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(1).Checked = True
            Else
                ImportActionList.Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction.ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction.ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction.ListItems.Count <> ImportActionList.Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmCloudMonitorAction
    End If
End Sub

Private Sub PasteActions()
    Dim ActionIndex As Long
    Dim SelectedActionIndex As Long
    Dim ByteData() As Byte
    Dim ByteNumber As Long
    Dim myClipboard As New cCustomClipboard
    Dim ClipboardFormatID As Long
    Dim strTypeNameLen As Long
    Dim strTypeName As String
    Dim clsAction As Object
    Dim After As Boolean
    Dim TotalAddedActions As Long
        
    'first need to figure out where to put the pasted actions
    'find the last selected action
    SelectedActionIndex = 0
    For ActionIndex = 1 To Me.lstAction.ListItems.Count
        If Me.lstAction.ListItems(ActionIndex).Selected Then
            SelectedActionIndex = ActionIndex
            Me.lstAction.ListItems(ActionIndex).Selected = False
        End If
    Next ActionIndex
    
    If SelectedActionIndex = 0 Then
        'nothing is selected - put it at the end
        SelectedActionIndex = Me.lstAction.ListItems.Count
        After = True
    Else
        After = True
    End If
    
    'now have the location to insert, get the data
    ClipboardFormatID = myClipboard.FormatIDForName(Me.hwnd, "CCD Commander Actions")
    If ClipboardFormatID = 0 Then
        'problem, no data to get!
        Exit Sub
    End If
    
    Call myClipboard.ClipboardOpen(Me.hwnd)
    Call myClipboard.GetBinaryData(ClipboardFormatID, ByteData())
    Call myClipboard.ClipboardClose
    
    'get current size of list
    TotalAddedActions = ImportActionList.Count
    
    'got the data, now parse it!
    ByteNumber = 0
    Call MainMod.LoadActionLists(ImportActionList, ByteData, ByteNumber, UBound(ByteData), SelectedActionIndex, Me.lstAction, After)
    
    'Compute how many actions were added
    TotalAddedActions = ImportActionList.Count - TotalAddedActions
    
    'now select all the actions that were just added, starting with selectedactionindex
    If (SelectedActionIndex = 0) Then
        SelectedActionIndex = 1
    ElseIf After Then
        SelectedActionIndex = SelectedActionIndex + 1
    End If
    
    Set Me.lstAction.SelectedItem = Me.lstAction.ListItems(SelectedActionIndex)
    For ActionIndex = SelectedActionIndex To (SelectedActionIndex + TotalAddedActions - 1)
        Me.lstAction.ListItems(ActionIndex).Selected = True
    Next ActionIndex
    
    Modified = True
    'Call CheckModifiedFlag
End Sub

Private Sub CopyActions()
    Dim ActionInfo As Object
    Dim ActionString As String
    Dim SelectIndex As Long
    Dim ByteData() As Byte
    Dim ByteCount As Long
    Dim strTypeName As String
    Dim strTypeNameLen As Long
    Dim myClipboard As New cCustomClipboard
    Dim ClipboardFormatID As Long
    Dim SelectedActions As New Collection

    If Me.lstAction.SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
    
    'Put all the selected actions into a new collection
    For SelectIndex = 1 To Me.lstAction.ListItems.Count
        If Me.lstAction.ListItems(SelectIndex).Selected Then
            Set ActionInfo = ImportActionList.Item(SelectIndex)
            Call SelectedActions.Add(ActionInfo)
        End If
    Next SelectIndex
        
    'now that the selected actions are in a collection, I can use the normal save/load functions
    'Compute how many bytes I need to store the data
    ByteCount = MainMod.GetNumberOfBytesForActionList(SelectedActions)
        
    'Redimention the array
    ReDim ByteData(0 To ByteCount - 1)
    
    ByteCount = 0
    Call SaveActionLists(SelectedActions, ByteData, ByteCount)
    
    'Now just need to put the byte array on the clipboard!
    ClipboardFormatID = myClipboard.FormatIDForName(Me.hwnd, "CCD Commander Actions")
    If ClipboardFormatID = 0 Then
        'need to create format
        ClipboardFormatID = myClipboard.AddFormat("CCD Commander Actions")
        
        If ClipboardFormatID = 0 Then
            MsgBox "Clipboard error."
            Exit Sub
        End If
    End If
    
    'Now copy the data to the clipboard
    Call myClipboard.ClipboardOpen(Me.hwnd)
    Call myClipboard.ClearClipboard
    Call myClipboard.SetBinaryData(ClipboardFormatID, ByteData)
    
    'all done, close the clipboard
    Call myClipboard.ClipboardClose
End Sub

Private Sub DeleteActions()
    Dim SelectIndex As Long
    Dim ActionInfo As Object
    Dim LastDeletedIndex As Long
    Dim LastUpperActionIndex As Long
    
    If Me.lstAction.SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
    
    For SelectIndex = Me.lstAction.ListItems.Count To 1 Step -1
        If Me.lstAction.ListItems(SelectIndex).Selected Then
            Set ActionInfo = ImportActionList.Item(SelectIndex)
            
            If TypeName(ActionInfo) = "RunActionList" Then
                Call MainMod.ClearSubActions(ActionInfo)
            End If
            
            Call ImportActionList.Remove(SelectIndex)
            Call Me.lstAction.ListItems.Remove(SelectIndex)
            LastDeletedIndex = SelectIndex
        End If
    Next SelectIndex

    If Me.lstAction.ListItems.Count > 0 Then
        If LastDeletedIndex < Me.lstAction.ListItems.Count Then
            Set Me.lstAction.SelectedItem = Me.lstAction.ListItems(LastDeletedIndex)
        Else
            Set Me.lstAction.SelectedItem = Me.lstAction.ListItems(Me.lstAction.ListItems.Count)
        End If
    End If

    Modified = True
End Sub

Private Sub MoveDown()
    Dim ActionInfo As Object
    Dim ActionString As String
    Dim LastIndex As Long
    Dim LastUpperActionIndex As Long
    
    If Me.lstAction.SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
        
    'first find the last action in the list that is selected.
    For LastIndex = Me.lstAction.ListItems.Count To 1 Step -1
        If Me.lstAction.ListItems(LastIndex).Selected Then
            Exit For
        End If
    Next LastIndex
    
    If LastIndex = Me.lstAction.ListItems.Count Then
        MsgBox "Cannot move any lower in the list!"
        Exit Sub
    End If
    
    For LastIndex = LastIndex To 1 Step -1
        If Me.lstAction.ListItems(LastIndex).Selected Then
            Set ActionInfo = ImportActionList.Item(LastIndex)
            Call ImportActionList.Remove(LastIndex)
            Call ImportActionList.Add(ActionInfo, , , LastIndex)
            
            ActionString = Me.lstAction.ListItems(LastIndex).Text
            Call Me.lstAction.ListItems.Remove(LastIndex)
            Call Me.lstAction.ListItems.Add(LastIndex + 1, , ActionString)
            Set Me.lstAction.SelectedItem = Me.lstAction.ListItems(LastIndex + 1)
            Me.lstAction.ListItems(LastIndex + 1).Checked = ActionInfo.Selected
        End If
    Next LastIndex

    Modified = True
    'Call CheckModifiedFlag
End Sub

Private Sub MoveUp()
    Dim ActionInfo As Object
    Dim ActionString As String
    Dim StartIndex As Integer
    Dim LastUpperActionIndex As Long
    
    If Me.lstAction.SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
        
    'first find the first action in the list that is selected.
    For StartIndex = 1 To Me.lstAction.ListItems.Count
        If Me.lstAction.ListItems(StartIndex).Selected Then
            Exit For
        End If
    Next StartIndex
            
    If StartIndex = 1 Then
        MsgBox "Cannot move any higher in the list!"
        Exit Sub
    End If
    
    If InStr(Me.lstAction.ListItems(StartIndex - 1).Text, "- Running") Or InStr(Me.lstAction.ListItems(StartIndex - 1).Text, "- Complete") Then
        MsgBox "Cannot move any higher in the list!"
        Exit Sub
    End If
    
    For StartIndex = StartIndex To Me.lstAction.ListItems.Count
        If Me.lstAction.ListItems(StartIndex).Selected Then
            Set ActionInfo = ImportActionList.Item(StartIndex)
            Call ImportActionList.Remove(StartIndex)
            Call ImportActionList.Add(ActionInfo, , StartIndex - 1)
            
            ActionString = Me.lstAction.ListItems(StartIndex).Text
            Call Me.lstAction.ListItems.Remove(StartIndex)
            Call Me.lstAction.ListItems.Add(StartIndex - 1, , ActionString)
            Set Me.lstAction.SelectedItem = Me.lstAction.ListItems(StartIndex - 1)
            Me.lstAction.ListItems(StartIndex - 1).Checked = ActionInfo.Selected
        End If
    Next StartIndex

    Modified = True
    'Call CheckModifiedFlag
End Sub

Private Sub mnuEdit_Click()
    Dim myClip As New cCustomClipboard
    Dim FormatID As Long
    
    If Me.ActiveControl Is Nothing Then
        Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = False
    Else
        If Me.ActiveControl.Name = "lstAction" Then
            If Not (Me.ActiveControl.SelectedItem Is Nothing) Then
                Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = True
            Else
                Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = False
            End If
            
            FormatID = myClip.FormatIDForName(Me.hwnd, "CCD Commander Actions")
            If FormatID <> 0 Then
                If myClip.HasCurrentFormat(FormatID) Then
                    Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = True
                Else
                    Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
                End If
            Else
                Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
            End If
        Else
            Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = False
            Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = False
            Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
            Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = False
            Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = False
            Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = False
        End If
    End If
End Sub

Private Sub mnuEditItem_Click(Index As Integer)
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    Select Case Index
        Case EditMenuItems.CopyMenu
            'copy
            If Me.lstAction.SelectedItem Is Nothing Then
                Beep
            Else
                Call CopyActions
            End If
        Case EditMenuItems.CutMenu
            'cut
            If Me.lstAction.SelectedItem Is Nothing Then
                Beep
            Else
                Call CopyActions
                If Not (Me.lstAction.SelectedItem Is Nothing) Then Call DeleteActions
            End If
        Case EditMenuItems.PasteMenu
            'paste
            Call PasteActions
        Case EditMenuItems.DeleteMenu
            'delete
            If Me.lstAction.SelectedItem Is Nothing Then
                Beep
            Else
                Call DeleteActions
            End If
        Case EditMenuItems.MoveUpMenu
            'move up
            If Me.lstAction.SelectedItem Is Nothing Then
                Beep
            Else
                Call MoveUp
            End If
        Case EditMenuItems.MoveDownMenu
            'move down
            If Me.lstAction.SelectedItem Is Nothing Then
                Beep
            Else
                Call MoveDown
            End If
    End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Dim FileName As String
    Dim Reply As Long
    Dim actFileNum As Integer
    Dim ByteData() As Byte
    Dim NumberOfBytes As Long
    
    If Me.MousePointer = vbHourglass Then
        Beep
        Exit Sub
    End If
    
    Select Case Index
        Case FileMenuItems.NewFileMenu
'            If Modified Then
'                If CurrentFileName <> "" Then
'                    Reply = MsgBox("Do you want to save the changes you made to " & CurrentFileName & "?", vbYesNoCancel + vbExclamation)
'                Else
'                    Reply = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbExclamation)
'                End If
'
'                If Reply = vbCancel Then
'                    Exit Sub
'                ElseIf Reply = vbYes Then
'                    If CurrentFileName <> "" Then
'                        Call SaveAction(CurrentFilePath, colAction)
'                    Else
'                        Call mnuFileItem_Click(FileMenuItems.SaveAsFileMenu)
'                    End If
'                End If
'            End If
            
            Me.lstAction.ListItems.Clear
            Call MainMod.ClearCollection(ImportActionList)
                        
            Modified = False
        Case FileMenuItems.OpenFileMenu
'            If Modified Then
'                If CurrentFileName <> "" Then
'                    Reply = MsgBox("Do you want to save the changes you made to " & CurrentFileName & "?", vbYesNoCancel + vbExclamation)
'                Else
'                    Reply = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbExclamation)
'                End If
'
'                If Reply = vbCancel Then
'                    Exit Sub
'                ElseIf Reply = vbYes Then
'                    If CurrentFileName <> "" Then
'                        Call SaveAction(CurrentFilePath, colAction)
'                    Else
'                        Call mnuFileItem_Click(FileMenuItems.SaveAsFileMenu)
'                    End If
'                End If
'            End If
            
            On Error Resume Next
            Call MkDir(App.Path & "\ImportActionTemplates")
            On Error GoTo 0
            Me.ComDlg.DialogTitle = "Load Action List Template"
            Me.ComDlg.Filter = "CCDCommander Action List Template (*.alt)|*.alt"
            Me.ComDlg.FilterIndex = 1
            Me.ComDlg.CancelError = True
            Me.ComDlg.InitDir = App.Path & "\ImportActionTemplates\"
            Me.ComDlg.flags = cdlOFNFileMustExist
            On Error Resume Next
            Me.ComDlg.ShowOpen
            If Err.Number = 0 Then
                On Error GoTo 0
                actFileNum = FreeFile
                Open Me.ComDlg.FileName For Binary Access Read As #actFileNum
            
                NumberOfBytes = LOF(actFileNum)
                
                ReDim ByteData(0 To NumberOfBytes - 1)
                Get #actFileNum, , ByteData()
                Close actFileNum
                
                Me.lstAction.ListItems.Clear
                Call MainMod.ClearCollection(ImportActionList)
                
                Call LoadActionLists(ImportActionList, ByteData, 0, UBound(ByteData), , Me.lstAction)
                Modified = False
            Else
                On Error GoTo 0
            End If
        Case FileMenuItems.SaveFileMenu
'            If CurrentFileName <> "" And CurrentFilePath <> "" Then
'                Call SaveAction(CurrentFilePath, colAction)
'                Modified = False
'                Me.Caption = "CCD Commander - " & CurrentFileName
'            Else
                On Error Resume Next
                Call MkDir(App.Path & "\ImportActionTemplates")
                On Error GoTo 0
                Me.ComDlg.DialogTitle = "Save Action List Template"
                Me.ComDlg.Filter = "CCDCommander Action List Template (*.alt)|*.alt"
                Me.ComDlg.FilterIndex = 1
'                If CurrentFilePath <> "" Then
'                    Me.ComDlg.InitDir = CurrentFilePath
'                Else
                    Me.ComDlg.InitDir = App.Path & "\ImportActionTemplates\"
'                End If
'                If CurrentFileName <> "" Then
'                    Me.ComDlg.FileName = CurrentFileName
'                Else
                    Me.ComDlg.FileName = "DefaultTemplate.alt"
'                End If
                Me.ComDlg.CancelError = True
                Me.ComDlg.flags = cdlOFNOverwritePrompt
                On Error Resume Next
                Me.ComDlg.ShowSave
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Call SaveAction(Me.ComDlg.FileName, ImportActionList)
                    Modified = False
                Else
                    On Error GoTo 0
                End If
'            End If
        Case FileMenuItems.CloseFileMenu
            Unload Me
    End Select
End Sub

Private Sub mnuHelp_Click()
    Call HtmlHelp(Me.hwnd, App.HelpFile, 15, Me.HelpContextID)
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case ToolBarButtons.NewFileBtn
            'new
            Call mnuFileItem_Click(FileMenuItems.NewFileMenu)
        Case ToolBarButtons.OpenFileBtn
            'open
            Call mnuFileItem_Click(FileMenuItems.OpenFileMenu)
        Case ToolBarButtons.SaveFileBtn
            'save
            Call mnuFileItem_Click(FileMenuItems.SaveFileMenu)
        Case ToolBarButtons.CopyBtn
            'copy
            Call mnuEditItem_Click(EditMenuItems.CopyMenu)
        Case ToolBarButtons.CutBtn
            'cut
            Call mnuEditItem_Click(EditMenuItems.CutMenu)
        Case ToolBarButtons.PasteBtn
            'paste
            Call mnuEditItem_Click(EditMenuItems.PasteMenu)
        Case ToolBarButtons.DeleteBtn
            'delete
            Call mnuEditItem_Click(EditMenuItems.DeleteMenu)
        Case ToolBarButtons.MoveUpBtn
            'move up
            Call mnuEditItem_Click(EditMenuItems.MoveUpMenu)
        Case ToolBarButtons.MoveDownBtn
            'move down
            Call mnuEditItem_Click(EditMenuItems.MoveDownMenu)
    End Select
End Sub

Public Sub ImportToList(ActionCollection As Collection)
    Dim listFileNum As Long
    Dim Name As String
    Dim RA As Double
    Dim Dec As Double
    Dim clsAction As Object
    Dim CopiedAction As Object
    Dim ByteData() As Byte
    Dim NumberOfBytes As Long
    
    listFileNum = FreeFile()
    
    'Open Input target list
    On Error Resume Next
    Open Me.lblFileName.Caption For Input As #listFileNum
    If Err.Number <> 0 Then
        On Error GoTo 0
        Call MsgBox("Cannot open Import List.  Please verify the path and file name.", vbCritical)
        Exit Sub
    End If
    On Error GoTo 0
    
    Do Until EOF(listFileNum)
        DoEvents
        
        Input #listFileNum, Name
        Input #listFileNum, RA
        Input #listFileNum, Dec
    
        For Each clsAction In ImportActionList
            Select Case TypeName(clsAction)
                Case "ImagerAction"
                    Set CopiedAction = New ImagerAction
                Case "MoveRADecAction"
                    Set CopiedAction = New MoveRADecAction
                Case "WaitForAltAction"
                    Set CopiedAction = New WaitForAltAction
                Case "ImageLinkSyncAction"
                    Set CopiedAction = New ImageLinkSyncAction
                Case "FocusAction"
                    Set CopiedAction = New FocusAction
                Case "RotatorAction"
                    Set CopiedAction = New RotatorAction
                Case "IntelligentTempAction"
                    Set CopiedAction = New IntelligentTempAction
                Case "WaitForTimeAction"
                    Set CopiedAction = New WaitForTimeAction
                Case "ParkMountAction"
                    Set CopiedAction = New ParkMountAction
                Case "RunScriptAction"
                    Set CopiedAction = New RunScriptAction
                Case "AutoFlatAction"
                    Set CopiedAction = New AutoFlatAction
                Case "SkipAheadAtTimeAction"
                    Set CopiedAction = New SkipAheadAtTimeAction
                Case "SkipAheadAtAltAction"
                    Set CopiedAction = New SkipAheadAtAltAction
                Case "SkipAheadAtHAAction"
                    Set CopiedAction = New SkipAheadAtHAAction
                Case "DomeAction"
                    Set CopiedAction = New DomeAction
                Case "CloudMonitorAction"
                    Set CopiedAction = New CloudMonitorAction
            End Select
            
            NumberOfBytes = clsAction.ByteArraySize
            ReDim ByteData(0 To NumberOfBytes - 1)
            
            NumberOfBytes = 0
            Call clsAction.SaveActionByteArray(ByteData, NumberOfBytes)
            
            NumberOfBytes = 0
            Call CopiedAction.LoadActionByteArray(ByteData, NumberOfBytes)
            
            Select Case TypeName(CopiedAction)
                Case "ImagerAction"
                    'CopiedAction.FileNamePrefix = Name & clsAction.FileNamePrefix
                Case "MoveRADecAction"
                    CopiedAction.Name = Name
                    CopiedAction.RA = RA
                    CopiedAction.Dec = Dec
                Case "WaitForAltAction"
                    CopiedAction.Name = Name
                    CopiedAction.RA = RA
                    CopiedAction.Dec = Dec
                Case "SkipAheadAtAltAction"
                    CopiedAction.Name = Name
                    CopiedAction.RA = RA
                    CopiedAction.Dec = Dec
                Case "SkipAheadAtHAAction"
                    CopiedAction.Name = Name
                    CopiedAction.RA = RA
            End Select
            
            ActionCollection.Add CopiedAction
        Next clsAction
    Loop
    
    Close #listFileNum
End Sub
