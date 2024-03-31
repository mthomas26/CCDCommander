VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CCD Commander"
   ClientHeight    =   6135
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   7935
   HelpContextID   =   100
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList 
      Left            =   5580
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":106C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1200
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1314
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1428
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":153C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRunning 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   420
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7380
      TabIndex        =   0
      Top             =   60
      Width           =   315
      Begin VB.Shape shpRunning 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   300
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   7200
      Top             =   420
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      DialogTitle     =   "CCD Commander"
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6120
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InBufferSize    =   16384
      BaudRate        =   19200
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
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
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Check Selected"
            Object.ToolTipText     =   "Check Selected"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Uncheck Selected"
            Object.ToolTipText     =   "Uncheck Selected"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Check All"
            Object.ToolTipText     =   "Check All"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Uncheck All"
            Object.ToolTipText     =   "Uncheck All"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Start Action List"
            Object.ToolTipText     =   "Start Action List"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Description     =   "Pause Action List"
            Object.ToolTipText     =   "Pause Action List"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Description     =   "Stop Action List"
            Object.ToolTipText     =   "Stop Action List"
            ImageIndex      =   17
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraTabs 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   780
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtFileName 
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   22
         Top             =   0
         Width           =   5655
      End
      Begin VB.CommandButton cmdCloseSubAction 
         Height          =   195
         Index           =   0
         Left            =   7500
         MaskColor       =   &H00D8E9EC&
         Picture         =   "Form2.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton optRunAborted 
         Caption         =   "Run until Aborted"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   20
         Top             =   4320
         Width           =   2115
      End
      Begin VB.OptionButton optRunPeriod 
         Caption         =   "Run for a Period of Time"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   19
         Top             =   4920
         Width           =   2115
      End
      Begin VB.TextBox txtTime 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   4860
         Width           =   795
      End
      Begin VB.OptionButton optRunOnce 
         Caption         =   "Run Once"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   15
         Top             =   4020
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optRunMultiple 
         Caption         =   "Run Multiple Times"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   14
         Top             =   4620
         Width           =   1875
      End
      Begin VB.TextBox txtNumRepeat 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   13
         Text            =   "0"
         Top             =   4560
         Width           =   795
      End
      Begin VB.CommandButton cmdLinkUnlink 
         Caption         =   "Link To File"
         Height          =   315
         Index           =   0
         Left            =   6360
         MaskColor       =   &H00D8E9EC&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin MSComctlLib.ListView lstAction 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   5741
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
      Begin VB.Frame fraSubAction 
         Caption         =   "Sub Action List"
         Height          =   3555
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   360
         Width           =   7515
      End
      Begin VB.Label lblSubActionName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   30
         Width           =   465
      End
      Begin VB.Label lblTimes 
         AutoSize        =   -1  'True
         Caption         =   "times"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   17
         Top             =   4620
         Width           =   360
      End
      Begin VB.Label lblMinutes 
         AutoSize        =   -1  'True
         Caption         =   "minutes"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   16
         Top             =   4920
         Width           =   540
      End
   End
   Begin VB.Frame fraTabs 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   7695
      Begin VB.Frame Frame3 
         Caption         =   "Action List"
         Height          =   5055
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   7515
         Begin MSComctlLib.ListView lstAction 
            Height          =   4755
            Index           =   1
            Left            =   60
            TabIndex        =   25
            Top             =   240
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   8387
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
      End
   End
   Begin VB.Frame fraStatus 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   7695
      Begin VB.Frame Frame4 
         Caption         =   "Status"
         Height          =   4635
         Left            =   60
         TabIndex        =   5
         Top             =   480
         Width           =   7515
         Begin RichTextLib.RichTextBox txtStatus 
            Height          =   4335
            HelpContextID   =   110
            Left            =   60
            TabIndex        =   6
            ToolTipText     =   "Current Status"
            Top             =   240
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   7646
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            DisableNoScroll =   -1  'True
            TextRTF         =   $"Form2.frx":175C
         End
      End
      Begin VB.TextBox txtCurrentAction 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Current Action:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   1050
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   5655
      Left            =   60
      TabIndex        =   8
      Top             =   420
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9975
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main Action Setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Save &As..."
         Index           =   3
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Import Target List"
         Index           =   5
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   ""
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   ""
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&xit"
         Index           =   13
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy Action"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "C&ut Action"
         Index           =   1
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste Action"
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "De&lete Action"
         Index           =   4
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Move Action Up"
         Index           =   6
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Move Action &Down"
         Index           =   7
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Check Selected"
         Index           =   9
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Uncheck Selected"
         Index           =   10
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "C&heck All"
         Index           =   11
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "U&ncheck All"
         Index           =   12
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Use Check &Boxes"
         Index           =   14
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Jump to Running Action"
         Enabled         =   0   'False
         Index           =   16
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
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Skip Ahead at Time"
         Index           =   8
         Shortcut        =   {F9}
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
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Temperature Control"
         Index           =   12
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Dome Control"
         Index           =   13
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Weather Monitor Control"
         Index           =   14
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Run External Program"
         Index           =   15
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Run Sub-Action List"
         Index           =   16
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Park Mount"
         Index           =   17
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuActionItems 
         Caption         =   "Comment"
         Index           =   18
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunItem 
         Caption         =   "&Start"
         Index           =   0
      End
      Begin VB.Menu mnuRunItem 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuRunItem 
         Caption         =   "S&top"
         Enabled         =   0   'False
         Index           =   2
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Graphs"
      Begin VB.Menu mnuViewItem 
         Caption         =   "Enable Temperature Recording"
         Index           =   0
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Show CCD &Temperature Plot..."
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Enable Guider Error Recording"
         Index           =   2
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "Show CCD Guide &Error Plot..."
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowItem 
         Caption         =   "Always on Top"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Overview..."
         Index           =   0
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Requirements..."
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "&Getting Started..."
         Index           =   2
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About CCD Commander..."
      End
   End
End
Attribute VB_Name = "frmMain"
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

Public RunningAction As Boolean

Private StartTime As Double
Private StartDate As Date

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
Private Const CommentIndex = 18
    
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
    CheckSelectedBtn = 14
    UncheckSelectedBtn = 15
    CheckAllBtn = 16
    UncheckAllBTn = 17
    PlayBtn = 19
    PauseBtn = 20
    StopBtn = 21
End Enum

Private Enum FileMenuItems
    NewFileMenu = 0
    OpenFileMenu = 1
    SaveFileMenu = 2
    SaveAsFileMenu = 3
    ImportTargetListMenu = 5
    OpenPreviousFile1 = 7
    OpenPreviousFile2 = 8
    OpenPreviousFile3 = 9
    OpenPreviousFile4 = 10
    OpenPreviousFile5 = 11
    OpenPreviousFileDivider = 12
    ExitMenu = 13
End Enum

Private Enum EditMenuItems
    CopyMenu = 0
    CutMenu = 1
    PasteMenu = 2
    DeleteMenu = 4
    MoveUpMenu = 6
    MoveDownMenu = 7
    CheckSelectedMenu = 9
    UncheckSelectedMenu = 10
    CheckAllMenu = 11
    UncheckAllMenu = 12
    UseCheckBoxes = 14
    JumpToRunningAction = 16
End Enum
    
Private Enum RunMenuItems
    StartMenu = 0
    PauseMenu = 1
    StopMenu = 2
End Enum
    
Private LastCloudSensorCheck As Date
Private LastSkipToCheck As Date
Private LastSimParkCheck As Date

Public CurrentFileName As String
Public CurrentFilePath As String
Public Modified As Boolean

Private myCurrentTab As Integer

Public SubActionControlCollection As Collection

Public ActionCollections As New Collection
Public SubActionInfo As New Collection
Public UpperActionIndexes As New Collection
Public AllowSubActionEdit As New Collection

Public EditingActionNumber As Long
Public EditingActionListLevel As Long

Public DisableSkipChecking As Boolean

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
    ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Sub AddStuffForSubAction(clsAction As RunActionList)
    Dim newTab As Object
    Dim myObject As Object
    Dim NewIndex As Integer
    Dim OnStatusTab As Boolean
    
    If Me.TabStrip.SelectedItem.Index <> Me.TabStrip.Tabs.Count Then
        OnStatusTab = False
    Else
        OnStatusTab = True
    End If
        
    'make tab
    Set newTab = Me.TabStrip.Tabs.Add(Me.TabStrip.Tabs.Count, , "Sub-Action")
    
    NewIndex = Me.fraTabs.Count
    Load Me.fraTabs(NewIndex)
    Me.TabStrip.ZOrder 1
    
    For Each myObject In SubActionControlCollection
        Load myObject(NewIndex)
        Set myObject(NewIndex).Container = Me.fraTabs(NewIndex)
        myObject(NewIndex).Visible = True
    Next myObject
    
    If Not OnStatusTab Then
        Set Me.TabStrip.SelectedItem = newTab
    Else
        Set Me.TabStrip.SelectedItem = Me.TabStrip.Tabs(Me.TabStrip.Tabs.Count)
    End If
    
    If Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked Then
        Me.lstAction(NewIndex).Checkboxes = True
    Else
        Me.lstAction(NewIndex).Checkboxes = False
    End If
    
    If clsAction Is Nothing Then
        Me.optRunOnce(NewIndex).Value = True
    Else
        If clsAction.LinkToFile Then
            'clear out existing actions
            Call MainMod.SearchAndDeleteSkipActions(clsAction.ActionCollection)
            Call MainMod.ClearCollection(clsAction.ActionCollection)
            clsAction.ActionListName = CheckIfFileExists(clsAction.ActionListName)
            If clsAction.ActionListName <> "" Then
                Call MainMod.LoadActionForRunActionList(clsAction.ActionListName, clsAction.ActionCollection)
                Me.txtFileName(NewIndex).Text = clsAction.ActionListName
                Me.txtFileName(NewIndex).Enabled = False
                Me.cmdLinkUnlink(NewIndex).Caption = "Unlink"
            Else
                clsAction.LinkToFile = False
                Me.txtFileName(NewIndex).Text = ""
                Me.txtFileName(NewIndex).Enabled = True
                Me.cmdLinkUnlink(NewIndex).Caption = "Link to File"
            End If
            
            're-setup skip times
            Call MainMod.SearchActionListForSkipActions(clsAction.ActionCollection)
        Else
            Me.txtFileName(NewIndex).Text = clsAction.ActionListName
            Me.txtFileName(NewIndex).Enabled = True
            Me.cmdLinkUnlink(NewIndex).Caption = "Link to File"
        End If
        
        Me.lstAction(NewIndex).ListItems.Clear
        For Each myObject In clsAction.ActionCollection
            If RunningAction And myObject.RunTimeStatus <> "" Then
                Me.lstAction(NewIndex).ListItems.Add , , myObject.BuildActionListString & " - " & myObject.RunTimeStatus
            Else
                Me.lstAction(NewIndex).ListItems.Add , , myObject.BuildActionListString
            End If
            Me.lstAction(NewIndex).ListItems(Me.lstAction(NewIndex).ListItems.Count).Checked = myObject.Selected
        Next myObject
        
        If clsAction.RepeatMode = 0 Then
            Me.optRunOnce(NewIndex).Value = True
        ElseIf clsAction.RepeatMode = 1 Then
            Me.optRunMultiple(NewIndex).Value = True
            Me.txtNumRepeat(NewIndex).Text = clsAction.TimesToRepeat
        ElseIf clsAction.RepeatMode = 2 Then
            Me.optRunPeriod(NewIndex).Value = True
            Me.txtTime(NewIndex).Text = clsAction.RepeatTime
        ElseIf clsAction.RepeatMode = 3 Then
            Me.optRunAborted(NewIndex).Value = True
        End If
    
        If (clsAction.RunTimeStatus = "Running") Or _
            (clsAction.RunTimeStatus = "Complete") Or _
            (clsAction.RunTimeStatus = "Skipped") Then
            
            Me.cmdLinkUnlink(NewIndex).Enabled = False
        End If
    End If
    
    If Not AllowSubActionEdit(NewIndex) Then
        Me.txtFileName(NewIndex).Enabled = False
    End If
    
    If Not AllowSubActionEdit(NewIndex - 1) Then
        Me.optRunAborted(NewIndex).Enabled = False
        Me.optRunMultiple(NewIndex).Enabled = False
        Me.optRunOnce(NewIndex).Enabled = False
        Me.optRunPeriod(NewIndex).Enabled = False
        Me.txtNumRepeat(NewIndex).Enabled = False
        Me.txtTime(NewIndex).Enabled = False
    End If
    
    If Not AllowSubActionEdit(NewIndex - 1) Then
        Me.cmdLinkUnlink(NewIndex).Enabled = False
    End If
End Sub

Public Sub RemoveStuffForSubAction(Index As Integer)
    Dim myObject As Object
    Dim CurrentSelectedIndex As Long
    
    CurrentSelectedIndex = Me.TabStrip.SelectedItem.Index
        
    Set Me.TabStrip.SelectedItem = Me.TabStrip.Tabs(1)
    
    Call Me.TabStrip.Tabs.Remove(Index)
    
    If CurrentSelectedIndex >= Index Then
        CurrentSelectedIndex = CurrentSelectedIndex - 1
    End If
    
    Set Me.TabStrip.SelectedItem = Me.TabStrip.Tabs(CurrentSelectedIndex)
    
    For Each myObject In SubActionControlCollection
        Unload myObject(Index)
    Next myObject
    
    Unload Me.fraTabs(Index)
End Sub

Public Sub RemoveSubActionTabsAfter(Index As Integer)
    Dim Counter As Integer
    
    For Counter = ActionCollections.Count To (Index + 1) Step -1
        Call CloseSubActionTab(Counter)
    Next Counter
End Sub

Public Sub AddAction(WhatAction As Integer)
    Dim ActionInfo As Object
    Dim SelectedActionIndex As Long
    Dim ActionIndex As Long
        
    If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
        MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    'first need to figure out where to put the pasted actions
    'find the last selected action
    SelectedActionIndex = 0
    For ActionIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(ActionIndex).Selected Then
            SelectedActionIndex = ActionIndex
            Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(ActionIndex).Selected = False
        End If
    Next ActionIndex
    
    'selected action may be running!
    If Me.RunningAction And SelectedActionIndex > 0 And SelectedActionIndex < Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count Then
        Do While InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Text, "- Skipped") > 0 Or _
            InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Text, "- Running") > 0 Or _
            InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Text, "- Complete") > 0
            
            SelectedActionIndex = SelectedActionIndex + 1
            
            If SelectedActionIndex >= Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count Then
                SelectedActionIndex = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
                Exit Do
            End If
        Loop
    End If
    
    If SelectedActionIndex = 0 Then
        'nothing is selected - put it at the end
        SelectedActionIndex = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
    End If
    
    If WhatAction = ImagerIndex Then
        Call frmCameraAction.Show(1, Me)
        
        If CBool(frmCameraAction.Tag) Then
            Modified = True
            Set ActionInfo = New ImagerAction
            Call frmCameraAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmCameraAction
    ElseIf WhatAction = MoveToRADecIndex Then
        Call frmMoveToRADec.Show(1, Me)
        
        If CBool(frmMoveToRADec.Tag) Then
            Modified = True
            Set ActionInfo = New MoveRADecAction
            Call frmMoveToRADec.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmMoveToRADec
    ElseIf WhatAction = SyncIndex Then
        Call frmImageLinkSync.Show(1, Me)
        
        If CBool(frmImageLinkSync.Tag) Then
            Modified = True
            Set ActionInfo = New ImageLinkSyncAction
            Call frmImageLinkSync.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmImageLinkSync
    ElseIf WhatAction = FocusMaxIndex Then
        Call frmFocusAction.Show(1, Me)
        
        If CBool(frmFocusAction.Tag) Then
            Modified = True
            Set ActionInfo = New FocusAction
            Call frmFocusAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmFocusAction
    ElseIf WhatAction = RotatorIndex Then
        Call frmRotator.Show(1, Me)
        
        If CBool(frmRotator.Tag) Then
            Modified = True
            Set ActionInfo = New RotatorAction
            Call frmRotator.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmRotator
    ElseIf WhatAction = IntelligentTempIndex Then
        Call frmTempControl.Show(1, Me)
        
        If CBool(frmTempControl.Tag) Then
            Modified = True
            Set ActionInfo = New IntelligentTempAction
            Call frmTempControl.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmTempControl
    ElseIf WhatAction = WaitForAlt Then
        Call frmWaitForAlt.Show(1, Me)
        
        If CBool(frmWaitForAlt.Tag) Then
            Modified = True
            Set ActionInfo = New WaitForAltAction
            Call frmWaitForAlt.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmWaitForAlt
    ElseIf WhatAction = WaitForTime Then
        Call frmWaitForTime.Show(1, Me)
        
        If CBool(frmWaitForTime.Tag) Then
            Modified = True
            Set ActionInfo = New WaitForTimeAction
            Call frmWaitForTime.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmWaitForTime
    ElseIf WhatAction = ParkMountIndex Then
        Call frmPark.Show(1, Me)
        
        If CBool(frmPark.Tag) Then
            Modified = True
            Set ActionInfo = New ParkMountAction
            Call frmPark.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmPark
    ElseIf WhatAction = RunScript Then
        Call frmRunScript.Show(1, Me)
        
        If CBool(frmRunScript.Tag) Then
            Modified = True
            Set ActionInfo = New RunScriptAction
            Call frmRunScript.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmRunScript
    ElseIf WhatAction = RunActionList Then
        Modified = True
    
        If RunningAction Then
            MainMod.FollowRunningAction = False
            Me.mnuEditItem(EditMenuItems.JumpToRunningAction).Enabled = True
        End If
        
        Set ActionInfo = New RunActionList
                
        Call RemoveSubActionTabsAfter(Me.TabStrip.SelectedItem.Index)
        
        'Copy link to the ActionInfo object into the array so I can access the object when the user makes changes on the window
        Call SubActionInfo.Add(ActionInfo, , , Me.TabStrip.SelectedItem.Index)
        
        'Create collection for Action object
        Set SubActionInfo(Me.TabStrip.SelectedItem.Index + 1).ActionCollection = New Collection
        
        'Copy link to the Collection into the collection array
        Call ActionCollections.Add(SubActionInfo(Me.TabStrip.SelectedItem.Index + 1).ActionCollection, , , Me.TabStrip.SelectedItem.Index)
        
        ActionInfo.Selected = True
    
        'Create entry in the current collection and list
        If SelectedActionIndex = 0 Then
            ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
            Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Call UpperActionIndexes.Add(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count, , , Me.TabStrip.SelectedItem.Index)
        Else
            ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
            Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            Call UpperActionIndexes.Add(SelectedActionIndex + 1, , , Me.TabStrip.SelectedItem.Index)
        End If
            
        Call AllowSubActionEdit.Add(True, , , Me.TabStrip.SelectedItem.Index)
        
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
            MsgBox "Error!  Counts don't match!"
            End
        End If
        Set ActionInfo = Nothing
    
        'Create new tab for the sub-action
        Call AddStuffForSubAction(Nothing)
    ElseIf WhatAction = AutoFlatIndex Then
        Call frmAutoFlat.Show(1, Me)
        
        If CBool(frmAutoFlat.Tag) Then
            Modified = True
            Set ActionInfo = New AutoFlatAction
            Call frmAutoFlat.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            If RunningAction Then
                Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmAutoFlat
    ElseIf WhatAction = SkipAheadAtTimeIndex Then
        Call frmSkipAtTime.Show(1, Me)
        
        If CBool(frmSkipAtTime.Tag) Then
            Modified = True
            Set ActionInfo = New SkipAheadAtTimeAction
            Call frmSkipAtTime.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            If RunningAction Then
                Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmSkipAtTime
    ElseIf WhatAction = SkipAheadAtAltIndex Then
        Call frmSkipAtAlt.Show(1, Me)
        
        If CBool(frmSkipAtAlt.Tag) Then
            Modified = True
            Set ActionInfo = New SkipAheadAtAltAction
            Call frmSkipAtAlt.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            If RunningAction Then
                Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmSkipAtAlt
    ElseIf WhatAction = SkipAheadAtHAIndex Then
        Call frmSkipAtHA.Show(1, Me)
        
        If CBool(frmSkipAtHA.Tag) Then
            Modified = True
            Set ActionInfo = New SkipAheadAtHAAction
            Call frmSkipAtHA.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            
            If RunningAction Then
                Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
            End If
            
            Set ActionInfo = Nothing
        End If
        
        Unload frmSkipAtHA
    ElseIf WhatAction = DomeActionIndex Then
        Call frmDomeAction.Show(1, Me)
        
        If CBool(frmDomeAction.Tag) Then
            Modified = True
            Set ActionInfo = New DomeAction
            Call frmDomeAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmDomeAction
    ElseIf WhatAction = CloudMonitorActionIndex Then
        Call frmCloudMonitorAction.Show(1, Me)
        
        If CBool(frmCloudMonitorAction.Tag) Then
            Modified = True
            Set ActionInfo = New CloudMonitorAction
            Call frmCloudMonitorAction.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmCloudMonitorAction
    ElseIf WhatAction = CommentIndex Then
        Call frmComment.Show(1, Me)
        
        If CBool(frmComment.Tag) Then
            Modified = True
            Set ActionInfo = New clsCommentAction
            Call frmComment.PutFormDataIntoClass(ActionInfo)
            ActionInfo.Selected = True
            If SelectedActionIndex = 0 Then
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(1).Checked = True
            Else
                ActionCollections(Me.TabStrip.SelectedItem.Index).Add ActionInfo, , , SelectedActionIndex
                Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(SelectedActionIndex + 1, , ActionInfo.BuildActionListString())
                Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex + 1).Checked = True
            End If
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count <> ActionCollections(Me.TabStrip.SelectedItem.Index).Count Then
                MsgBox "Error!  Counts don't match!"
                End
            End If
            Set ActionInfo = Nothing
        End If
        
        Unload frmComment
    End If
    
    Call CheckModifiedFlag
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
        
    If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
        MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    'first need to figure out where to put the pasted actions
    'find the last selected action
    SelectedActionIndex = 0
    For ActionIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(ActionIndex).Selected Then
            SelectedActionIndex = ActionIndex
            Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(ActionIndex).Selected = False
        End If
    Next ActionIndex
    
    If SelectedActionIndex = 0 Then
        'nothing is selected - put it at the end
        SelectedActionIndex = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
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
    TotalAddedActions = ActionCollections(Me.TabStrip.SelectedItem.Index).Count
    
    'got the data, now parse it!
    ByteNumber = 0
    Call MainMod.LoadActionLists(ActionCollections(Me.TabStrip.SelectedItem.Index), ByteData, ByteNumber, UBound(ByteData), SelectedActionIndex, frmMain.lstAction(Me.TabStrip.SelectedItem.Index), After)
    
    'Compute how many actions were added
    TotalAddedActions = ActionCollections(Me.TabStrip.SelectedItem.Index).Count - TotalAddedActions
    
    'now select all the actions that were just added, starting with selectedactionindex
    If (SelectedActionIndex = 0) Then
        SelectedActionIndex = 1
    ElseIf After Then
        SelectedActionIndex = SelectedActionIndex + 1
    End If
    
    Set frmMain.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem = frmMain.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectedActionIndex)
    For ActionIndex = SelectedActionIndex To (SelectedActionIndex + TotalAddedActions - 1)
        frmMain.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(ActionIndex).Selected = True
    Next ActionIndex
    
    Modified = True
    Call CheckModifiedFlag
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

    If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
    
    'Don't want to allow copying completed or running actions
    'Hmmm---why not?
'    For SelectIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
'        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Selected And (InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Text, "Running") Or InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Text, "Complete")) Then
'            Beep
'            Exit Sub
'        End If
'    Next SelectIndex
        
    'Put all the selected actions into a new collection
    For SelectIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Selected Then
            Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(SelectIndex)
            Call SelectedActions.Add(ActionInfo)
        End If
    Next SelectIndex
        
    'now that the selected actions are in a collection, I can use the normal save/load functions
    'Compute how many bytes I need to store the data
    ByteCount = MainMod.GetNumberOfBytesForActionList(SelectedActions)
        
'    ByteCount = 0
'    For SelectIndex = 0 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListCount - 1
'        If Me.lstAction(Me.TabStrip.SelectedItem.Index).Selected(SelectIndex) Then
'            Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(SelectIndex + 1)
'
'            strTypeName = TypeName(ActionInfo)
'            strTypeNameLen = Len(strTypeName)
'            ByteCount = ByteCount + ActionInfo.ByteArraySize() + Len(strTypeNameLen) + Len(strTypeName)
'        End If
'    Next SelectIndex
    
    'Redimention the array
    ReDim ByteData(0 To ByteCount - 1)
    
    ByteCount = 0
    Call SaveActionLists(SelectedActions, ByteData, ByteCount)
    
'    For SelectIndex = 0 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListCount - 1
'        If Me.lstAction(Me.TabStrip.SelectedItem.Index).Selected(SelectIndex) Then
'            Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(SelectIndex + 1)
'            strTypeName = TypeName(ActionInfo)
'            strTypeNameLen = Len(strTypeName)
'            Call CopyMemory(ByteData(ByteCount), strTypeNameLen, Len(strTypeNameLen))
'            ByteCount = ByteCount + Len(strTypeNameLen)
'
'            Call CopyStringToByteArray(ByteData(), ByteCount, strTypeName)
'
'            Call ActionInfo.SaveActionByteArray(ByteData(), ByteCount)
'        End If
'    Next SelectIndex
    
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
    
    If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
    
    If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
        MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    'Don't want to allow copying completed or running actions
    For SelectIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Selected And (InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Text, "- Running") Or InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Text, "- Complete")) Then
            Beep
            Exit Sub
        End If
    Next SelectIndex
           
    For SelectIndex = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count To 1 Step -1
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(SelectIndex).Selected Then
            Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(SelectIndex)
            
            If RunningAction Then
                If TypeName(ActionInfo) = "SkipAheadAtTimeAction" Or TypeName(ActionInfo) = "SkipAheadAtAltAction" Or TypeName(ActionInfo) = "SkipAheadAtHAAction" Then
                    Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
                ElseIf TypeName(ActionInfo) = "AutoFlatAction" Then
                    If ActionInfo.FlatLocation = DuskSkyFlat Or ActionInfo.FlatLocation = DawnSkyFlat Then
                        Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
                    End If
                ElseIf TypeName(ActionInfo) = "RunActionList" Then
                    Call MainMod.SearchAndDeleteSkipActions(ActionInfo.ActionCollection)
                End If
            End If
            
            If Me.TabStrip.SelectedItem.Index < UpperActionIndexes.Count Then
                'check if I'm deleting the one with the tabs
                If UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) = SelectIndex Then
                    Call RemoveSubActionTabsAfter(Me.TabStrip.SelectedItem.Index)
                'now check if the index is below the one to be deleted
                ElseIf UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) > SelectIndex Then
                    'decrease index by one
                    LastUpperActionIndex = UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) - 1
                    
                    Call UpperActionIndexes.Remove(Me.TabStrip.SelectedItem.Index + 1)
                    Call UpperActionIndexes.Add(LastUpperActionIndex, , , Me.TabStrip.SelectedItem.Index)
                End If
            End If
            
            If TypeName(ActionInfo) = "RunActionList" Then
                Call MainMod.ClearSubActions(ActionInfo)
            End If
            
            Call ActionCollections(Me.TabStrip.SelectedItem.Index).Remove(SelectIndex)
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Remove(SelectIndex)
            LastDeletedIndex = SelectIndex
        End If
    Next SelectIndex

    If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count > 0 Then
        If LastDeletedIndex < Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count Then
            Set Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastDeletedIndex)
        Else
            Set Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count)
        End If
    End If

    Modified = True
    Call CheckModifiedFlag
End Sub

Private Sub MoveDown()
    Dim ActionInfo As Object
    Dim ActionString As String
    Dim LastIndex As Long
    Dim LastUpperActionIndex As Long
    
    If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
        
    If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
        MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    'Don't want to allow copying completed or running actions
    For LastIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex).Selected And (InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex).Text, "- Running") Or InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex).Text, "- Complete")) Then
            Beep
            Exit Sub
        End If
    Next LastIndex
        
    'first find the last action in the list that is selected.
    For LastIndex = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count To 1 Step -1
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex).Selected Then
            Exit For
        End If
    Next LastIndex
    
    If LastIndex = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count Then
        MsgBox "Cannot move any lower in the list!"
        Exit Sub
    End If
    
    For LastIndex = LastIndex To 1 Step -1
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex).Selected Then
            If Me.TabStrip.SelectedItem.Index < UpperActionIndexes.Count Then
                'check if I'm moving the one with the tabs
                If UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) = LastIndex Then
                    'increase index by one
                    LastUpperActionIndex = UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) + 1
                    
                    Call UpperActionIndexes.Remove(Me.TabStrip.SelectedItem.Index + 1)
                    Call UpperActionIndexes.Add(LastUpperActionIndex, , , Me.TabStrip.SelectedItem.Index)
                End If
            End If
            
            Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(LastIndex)
            Call ActionCollections(Me.TabStrip.SelectedItem.Index).Remove(LastIndex)
            Call ActionCollections(Me.TabStrip.SelectedItem.Index).Add(ActionInfo, , , LastIndex)
            
            ActionString = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex).Text
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Remove(LastIndex)
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(LastIndex + 1, , ActionString)
            Set Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex + 1)
            Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(LastIndex + 1).Checked = ActionInfo.Selected
        End If
    Next LastIndex

    Modified = True
    Call CheckModifiedFlag
End Sub

Private Sub MoveUp()
    Dim ActionInfo As Object
    Dim ActionString As String
    Dim StartIndex As Integer
    Dim LastUpperActionIndex As Long
    
    If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first."
        Exit Sub
    End If
        
    If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
        MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    'Don't want to allow copying completed or running actions
    For StartIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex).Selected And (InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex).Text, "- Running") Or InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex).Text, "- Complete")) Then
            Beep
            Exit Sub
        End If
    Next StartIndex
                  
    'first find the first action in the list that is selected.
    For StartIndex = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex).Selected Then
            Exit For
        End If
    Next StartIndex
            
    If StartIndex = 1 Then
        MsgBox "Cannot move any higher in the list!"
        Exit Sub
    End If
    
    If InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex - 1).Text, "- Running") Or InStr(Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex - 1).Text, "- Complete") Then
        MsgBox "Cannot move any higher in the list!"
        Exit Sub
    End If
    
    For StartIndex = StartIndex To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
        If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex).Selected Then
            If Me.TabStrip.SelectedItem.Index < UpperActionIndexes.Count Then
                'check if I'm moving the one with the tabs
                If UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) = StartIndex Then
                    'decrease index by one
                    LastUpperActionIndex = UpperActionIndexes(Me.TabStrip.SelectedItem.Index + 1) - 1
                    
                    Call UpperActionIndexes.Remove(Me.TabStrip.SelectedItem.Index + 1)
                    Call UpperActionIndexes.Add(LastUpperActionIndex, , , Me.TabStrip.SelectedItem.Index)
                End If
            End If
            
            Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(StartIndex)
            Call ActionCollections(Me.TabStrip.SelectedItem.Index).Remove(StartIndex)
            Call ActionCollections(Me.TabStrip.SelectedItem.Index).Add(ActionInfo, , StartIndex - 1)
            
            ActionString = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex).Text
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Remove(StartIndex)
            Call Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Add(StartIndex - 1, , ActionString)
            Set Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem = Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex - 1)
            Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(StartIndex - 1).Checked = ActionInfo.Selected
        End If
    Next StartIndex

    Modified = True
    Call CheckModifiedFlag
End Sub

Private Sub PauseAndResumeAction()
    If Not Paused And RunningAction Then
        Paused = True
        Call AddToStatus("Action Paused.")
        
        Me.mnuRunItem(RunMenuItems.StartMenu).Enabled = True
        Me.mnuRunItem(RunMenuItems.PauseMenu).Enabled = False
        Me.mnuRunItem(RunMenuItems.StopMenu).Enabled = True
        Me.Toolbar.Buttons(ToolBarButtons.PlayBtn).Enabled = True
        Me.Toolbar.Buttons(ToolBarButtons.PauseBtn).Enabled = False
        Me.Toolbar.Buttons(ToolBarButtons.StopBtn).Enabled = True
        
        Me.tmrRunning.Enabled = False
        Me.shpRunning.FillColor = &HFFFF&
        Me.shpRunning.Visible = True
    ElseIf Paused Then
        Me.tmrRunning.Enabled = True
        Me.shpRunning.FillColor = &HC000&
        Paused = False
        Call AddToStatus("Action Resumed.")
        
        Me.mnuRunItem(RunMenuItems.StartMenu).Enabled = False
        Me.mnuRunItem(RunMenuItems.PauseMenu).Enabled = True
        Me.mnuRunItem(RunMenuItems.StopMenu).Enabled = True
        Me.Toolbar.Buttons(ToolBarButtons.PlayBtn).Enabled = False
        Me.Toolbar.Buttons(ToolBarButtons.PauseBtn).Enabled = True
        Me.Toolbar.Buttons(ToolBarButtons.StopBtn).Enabled = True
    End If
End Sub

Public Sub StartAndAbortAction()
    If Not RunningAction Then
        Me.mnuRunItem(RunMenuItems.StartMenu).Enabled = False
        Me.mnuRunItem(RunMenuItems.PauseMenu).Enabled = True
        Me.mnuRunItem(RunMenuItems.StopMenu).Enabled = True
        Me.Toolbar.Buttons(ToolBarButtons.PlayBtn).Enabled = False
        Me.Toolbar.Buttons(ToolBarButtons.PauseBtn).Enabled = True
        Me.Toolbar.Buttons(ToolBarButtons.StopBtn).Enabled = True
        
        RunningAction = True
        RunningActionListLevel = 1
        
        LastCloudSensorCheck = 0
        LastSkipToCheck = 0
        LastSimParkCheck = 0
        Me.shpRunning.FillColor = &HC000&
        'Call StartTimer
        
        'close sub-action tabs
        Call frmMain.RemoveSubActionTabsAfter(1)
        
        'Select last tab - this will always be the status tab
        Set TabStrip.SelectedItem = TabStrip.Tabs(TabStrip.Tabs.Count)
                
        Call StartAction
        
        'Just in case it was enabled at the end of the run
        Me.mnuEditItem(EditMenuItems.JumpToRunningAction).Enabled = False
        
        If Not Mount.SimulatedPark Then
            'The timer will stay running if a simulated park is in effect.
            RunningAction = False
            
            Me.mnuRunItem(RunMenuItems.StartMenu).Enabled = True
            Me.mnuRunItem(RunMenuItems.PauseMenu).Enabled = False
            Me.mnuRunItem(RunMenuItems.StopMenu).Enabled = False
            Me.Toolbar.Buttons(ToolBarButtons.PlayBtn).Enabled = True
            Me.Toolbar.Buttons(ToolBarButtons.PauseBtn).Enabled = False
            Me.Toolbar.Buttons(ToolBarButtons.StopBtn).Enabled = False
        
            Me.tmrRunning.Enabled = False
            
            Me.shpRunning.Visible = True
            Me.shpRunning.FillColor = &HFF&
        Else
            Call AddToStatus("Simulated Park continues to reposition the mount until the Abort button is pushed.")
        End If
    Else
        Paused = False
        Call AbortAction
        Me.tmrRunning.Enabled = False
        Me.shpRunning.Visible = True
        Me.shpRunning.FillColor = &HFF&
        
        If Mount.SimulatedPark Then
            RunningAction = False
            
            Me.mnuRunItem(RunMenuItems.StartMenu).Enabled = True
            Me.mnuRunItem(RunMenuItems.PauseMenu).Enabled = False
            Me.mnuRunItem(RunMenuItems.StopMenu).Enabled = False
            Me.Toolbar.Buttons(ToolBarButtons.PlayBtn).Enabled = True
            Me.Toolbar.Buttons(ToolBarButtons.PauseBtn).Enabled = False
            Me.Toolbar.Buttons(ToolBarButtons.StopBtn).Enabled = False
        
            Mount.SimulatedPark = False
            
            If (frmOptions.chkDisconnectAtEnd = vbChecked) Then
                Call Camera.CameraUnload
                Call Mount.MountUnload
                Call Focus.FocusUnload
                Call Rotator.RotatorUnload
                Call CloudSensor.CloudSensorSetup
                Call DomeControl.DomeUnload
            End If
        End If
    End If
End Sub

Public Sub StartTimer()
    StartTime = Timer
    StartDate = Now
    Me.tmrRunning.Enabled = True
End Sub

Private Sub cmdCloseSubAction_Click(Index As Integer)
    If RunningAction Then
        MainMod.FollowRunningAction = False
        Me.mnuEditItem(EditMenuItems.JumpToRunningAction).Enabled = True
    End If
    
    Call CloseSubActionTab(Index)
End Sub

Public Sub CloseSubActionTab(Index As Integer)
    Call RemoveSubActionTabsAfter(Index)
        
    Call RemoveStuffForSubAction(Index)
            
    ActionCollections.Remove Index
    SubActionInfo.Remove Index
    UpperActionIndexes.Remove Index
    AllowSubActionEdit.Remove Index
End Sub


Private Sub cmdLinkUnlink_Click(Index As Integer)
    Dim clsAction As Object
    Dim Responce As Long
    Dim SavePrompt As Boolean
    
    Call RemoveSubActionTabsAfter(Index)
    
    If SubActionInfo(Index).LinkToFile Then
        SubActionInfo(Index).LinkToFile = False
        Call AllowSubActionEdit.Remove(Index)
        Call AllowSubActionEdit.Add(True, , , Index - 1)
        Me.txtFileName(Index).Enabled = True
        Me.cmdLinkUnlink(Index).Caption = "Link To File"
    Else
        If ActionCollections(Index).Count > 0 Then
            Responce = MsgBox("Do you want to save the existing sub-action list?", vbQuestion + vbYesNoCancel)
        Else
            Responce = vbNo
        End If
        
        If Responce = vbYes Then
            'check if current name looks like a file name
            'A file name will either start with a drive letter like "C:"
            'Or a network directory, like \\NetworkPath
            'So look for either of these.
            If Mid(SubActionInfo(Index).ActionListName, 2, 1) = ":" Or Left(SubActionInfo(Index).ActionListName, 2) = "\\" Then
                'I think we have a path & file name, try to save it
                If Not MainMod.SaveAction(SubActionInfo(Index).ActionListName, ActionCollections(Index)) Then
                    SavePrompt = True
                Else
                    SavePrompt = False
                    SubActionInfo(Index).LinkToFile = True
                    Call AllowSubActionEdit.Remove(Index)
                    Call AllowSubActionEdit.Add(False, , , Index - 1)
                    Me.txtFileName(Index).Enabled = False
                End If
            Else
                SavePrompt = True
            End If
            
            If SavePrompt Then
                'Need to prompt for the file name
                On Error Resume Next
                Call MkDir(App.Path & "\Actions")
                On Error GoTo 0
                Me.ComDlg.DialogTitle = "Save Sub-Action List"
                Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
                Me.ComDlg.FilterIndex = 1
                Me.ComDlg.InitDir = App.Path & "\Actions\"
                If SubActionInfo(Index).ActionListName <> "" Then
                    Me.ComDlg.FileName = SubActionInfo(Index).ActionListName
                Else
                    Me.ComDlg.FileName = Format(Now, "yymmdd") & ".act"
                End If
                Me.ComDlg.CancelError = True
                Me.ComDlg.flags = cdlOFNOverwritePrompt
                On Error Resume Next
                Me.ComDlg.ShowSave
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Call MainMod.SaveAction(Me.ComDlg.FileName, ActionCollections(Index))
                    SubActionInfo(Index).ActionListName = Me.ComDlg.FileName
                    SubActionInfo(Index).LinkToFile = True
                    Me.txtFileName(Index).Text = Me.ComDlg.FileName
                    Me.txtFileName(Index).Enabled = False
                    Me.cmdLinkUnlink(Index).Caption = "Unlink"
                Else
                    On Error GoTo 0
                End If
            End If
        ElseIf Responce = vbNo Then
            On Error Resume Next
            Call MkDir(App.Path & "\Actions")
            On Error GoTo 0
            Me.ComDlg.DialogTitle = "Load Sub-Action List"
            Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
            Me.ComDlg.FilterIndex = 1
            Me.ComDlg.CancelError = True
            Me.ComDlg.InitDir = App.Path & "\Actions\"
            Me.ComDlg.flags = cdlOFNFileMustExist
            On Error Resume Next
            Me.ComDlg.ShowOpen
            If Err.Number = 0 Then
                On Error GoTo 0
                
                If RunningAction Then
                    Call MainMod.SearchAndDeleteSkipActions(ActionCollections(Index))
                End If
                
                Call MainMod.ClearCollection(ActionCollections(Index))
                
                Call MainMod.LoadActionForRunActionList(Me.ComDlg.FileName, ActionCollections(Index))
                Me.txtFileName(Index).Text = Me.ComDlg.FileName
                Me.txtFileName(Index).Enabled = False
                Me.cmdLinkUnlink(Index).Caption = "Unlink"
                
                Modified = True
                Call CheckModifiedFlag
                
                Call Me.lstAction(Index).ListItems.Clear
                
                For Each clsAction In ActionCollections(Index)
                    Call Me.lstAction(Index).ListItems.Add(, , clsAction.BuildActionListString)
                    Me.lstAction(Index).ListItems(Me.lstAction(Index).ListItems.Count).Checked = clsAction.Selected
                Next clsAction
                            
                SubActionInfo(Index).ActionListName = Me.ComDlg.FileName
                SubActionInfo(Index).LinkToFile = True
                Call AllowSubActionEdit.Remove(Index)
                Call AllowSubActionEdit.Add(False, , , Index - 1)
    
                If RunningAction Then
                    Call SearchActionListForSkipActions(ActionCollections(Index))
                End If
            Else
                On Error GoTo 0
            End If
        End If
    End If
    
    Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
End Sub

Private Sub Form_Load()
    Me.Top = GetMySetting("WindowPositions", "MainTop", 200)
    Me.Left = GetMySetting("WindowPositions", "MainLeft", 200)
    
    Call frmMain.UpdatePreviousFileMenu
    
    myCurrentTab = 1
    
    RunningAction = False
    DisableSkipChecking = False
    
    Set SubActionControlCollection = New Collection
       
    With SubActionControlCollection
        .Add Me.lblSubActionName
        .Add Me.txtFileName
        .Add Me.cmdLinkUnlink
        .Add Me.fraSubAction
        .Add Me.lstAction
        .Add Me.optRunOnce
        .Add Me.optRunAborted
        .Add Me.optRunMultiple
        .Add Me.optRunPeriod
        .Add Me.txtNumRepeat
        .Add Me.txtTime
        .Add Me.lblMinutes
        .Add Me.lblTimes
        .Add Me.cmdCloseSubAction
    End With
       
    ActionCollections.Add MainMod.colAction
    
    SubActionInfo.Add Nothing   'First entry is never used
    
    UpperActionIndexes.Add Nothing 'First entry is never used
    
    AllowSubActionEdit.Add True
    
    Me.lstAction(1).Checkboxes = CBool(MainMod.GetMySetting("ProgramSettings", "Checkboxes", "False"))
    'set to the opposite state
    Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked = Not Me.lstAction(1).Checkboxes
    'now simulate a menu click so everything gets setup properly
    Call mnuEditItem_Click(EditMenuItems.UseCheckBoxes)
    
    'set the check box to the opposite state from what it stored
    Me.mnuWindowItem(0).Checked = CBool(MainMod.GetMySetting("WindowPositions", "AlwaysOnTop", "False"))
    Call MainMod.SetOnTopMode(Me)
    
    'set the check box to the opposite state from what it stored
    Me.mnuViewItem(0).Checked = Not CBool(MainMod.GetMySetting("GraphOptions", "EnableTempRecording", "True"))
    'now simulate a menu click so the proper value gets setup
    Call mnuViewItem_Click(0)
    
    'set the check box to the opposite state from what it stored
    Me.mnuViewItem(2).Checked = Not CBool(MainMod.GetMySetting("GraphOptions", "EnableAutoguiderRecording", "True"))
    'now simulate a menu click so the proper value gets setup
    Call mnuViewItem_Click(2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Reply As Long
    
    If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
        If RunningAction Then
            MsgBox "Please abort the action before exiting.", vbExclamation
            Cancel = True
            Exit Sub
        End If
        
        If Modified And Me.ActionCollections.Item(1).Count > 0 Then
            If CurrentFileName <> "" Then
                Reply = MsgBox("Do you want to save the changes you made to " & CurrentFileName & "?", vbYesNoCancel + vbExclamation)
            Else
                Reply = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbExclamation)
            End If
            
            If Reply = vbCancel Then
                Cancel = True
                Exit Sub
            ElseIf Reply = vbYes Then
                If CurrentFileName <> "" Then
                    Call SaveAction(CurrentFilePath, colAction)
                Else
                    Call mnuFileItem_Click(FileMenuItems.SaveFileMenu)
                End If
            End If
        End If
        
        Exiting = True
        If RunningAction Then Call AbortAction
        Call UnloadStuff
    End If
    
    If Me.WindowState <> vbMinimized Then
        Call SaveMySetting("WindowPositions", "MainTop", Me.Top)
        Call SaveMySetting("WindowPositions", "MainLeft", Me.Left)
    End If
End Sub

Private Sub Form_Resize()
    If mnuWindowItem(0).Checked = True Then
        'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        mnuWindowItem(0).Checked = True
        Call MainMod.SaveMySetting("WindowPositions", "AlwaysOnTop", True)
        
        Call MainMod.SetOnTopMode(Me)
        Call MainMod.SetOnTopMode(frmOptions)
    Else
        'SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        mnuWindowItem(0).Checked = False
        Call MainMod.SaveMySetting("WindowPositions", "AlwaysOnTop", False)
    
        Call MainMod.SetOnTopMode(Me)
        Call MainMod.SetOnTopMode(frmOptions)
    End If
End Sub

Private Sub lstAction_DblClick(Index As Integer)
    Dim ActionInfo As Object
    Dim AllowEdits As Boolean
    
    If Me.lstAction(Index).SelectedItem Is Nothing Then
        MsgBox "You must select an existing action first.", vbExclamation
        Exit Sub
    End If
    
    Set ActionInfo = ActionCollections(Index).Item(Me.lstAction(Index).SelectedItem.Index)
    
    If ((ActionInfo.RunTimeStatus = "Running") And (TypeName(ActionInfo) <> "RunActionList")) Or _
        (ActionInfo.RunTimeStatus = "Complete") Or _
        (ActionInfo.RunTimeStatus = "Skipped") Then
        Beep
        Exit Sub
    End If
    
    EditingActionNumber = Me.lstAction(Index).SelectedItem.Index
    EditingActionListLevel = Index
    
    AllowEdits = AllowSubActionEdit(Index)
    If Index > 1 Then
        If SubActionInfo(Index).LinkToFile Then
            MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
            'Exit Sub
            AllowEdits = False
        End If
    End If
    
    If TypeName(ActionInfo) = "ImagerAction" Then
        Call frmCameraAction.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmCameraAction.OKButton.Enabled = False
        Call frmCameraAction.Show(1, Me)
        Call frmCameraAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmCameraAction.Tag) Then Modified = True
        Unload frmCameraAction
    ElseIf TypeName(ActionInfo) = "MoveRADecAction" Then
        Call frmMoveToRADec.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmMoveToRADec.OKButton.Enabled = False
        Call frmMoveToRADec.Show(1, Me)
        Call frmMoveToRADec.PutFormDataIntoClass(ActionInfo)
        If CBool(frmMoveToRADec.Tag) Then Modified = True
        Unload frmMoveToRADec
    ElseIf TypeName(ActionInfo) = "FocusAction" Then
        Call frmFocusAction.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmFocusAction.OKButton.Enabled = False
        Call frmFocusAction.Show(1, Me)
        Call frmFocusAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmFocusAction.Tag) Then Modified = True
        Unload frmFocusAction
    ElseIf TypeName(ActionInfo) = "RotatorAction" Then
        Call frmRotator.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmRotator.OKButton.Enabled = False
        Call frmRotator.Show(1, Me)
        Call frmRotator.PutFormDataIntoClass(ActionInfo)
        If CBool(frmRotator.Tag) Then Modified = True
        Unload frmRotator
    ElseIf TypeName(ActionInfo) = "IntelligentTempAction" Then
        Call frmTempControl.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmTempControl.OKButton.Enabled = False
        Call frmTempControl.Show(1, Me)
        Call frmTempControl.PutFormDataIntoClass(ActionInfo)
        If CBool(frmTempControl.Tag) Then Modified = True
        Unload frmTempControl
    ElseIf TypeName(ActionInfo) = "ImageLinkSyncAction" Then
        Call frmImageLinkSync.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmImageLinkSync.OKButton.Enabled = False
        Call frmImageLinkSync.Show(1, Me)
        Call frmImageLinkSync.PutFormDataIntoClass(ActionInfo)
        If CBool(frmImageLinkSync.Tag) Then Modified = True
        Unload frmImageLinkSync
    ElseIf TypeName(ActionInfo) = "WaitForAltAction" Then
        Call frmWaitForAlt.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmWaitForAlt.OKButton.Enabled = False
        Call frmWaitForAlt.Show(1, Me)
        Call frmWaitForAlt.PutFormDataIntoClass(ActionInfo)
        If CBool(frmWaitForAlt.Tag) Then Modified = True
        Unload frmWaitForAlt
    ElseIf TypeName(ActionInfo) = "WaitForTimeAction" Then
        Call frmWaitForTime.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmWaitForTime.OKButton.Enabled = False
        Call frmWaitForTime.Show(1, Me)
        Call frmWaitForTime.PutFormDataIntoClass(ActionInfo)
        If CBool(frmWaitForTime.Tag) Then Modified = True
        Unload frmWaitForTime
    ElseIf TypeName(ActionInfo) = "ParkMountAction" Then
        Call frmPark.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmPark.OKButton.Enabled = False
        Call frmPark.Show(1, Me)
        Call frmPark.PutFormDataIntoClass(ActionInfo)
        If CBool(frmPark.Tag) Then Modified = True
        Unload frmPark
    ElseIf TypeName(ActionInfo) = "RunScriptAction" Then
        Call frmRunScript.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmRunScript.OKButton.Enabled = False
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
            Call RemoveSubActionTabsAfter(Me.TabStrip.SelectedItem.Index)
                    
            'Copy link to the ActionInfo object into the array so I can access the object when the user makes changes on the window
            Call SubActionInfo.Add(ActionInfo, , , Index)
            
            'Copy link to the Collection into the collection array
            Call ActionCollections.Add(SubActionInfo(Index + 1).ActionCollection, , , Index)
                
            Call UpperActionIndexes.Add(Me.lstAction(Index).SelectedItem.Index, , , Index)
            
            If AllowEdits And ActionInfo.LinkToFile Then
                AllowEdits = False
            End If
            
            Call AllowSubActionEdit.Add(AllowEdits, , , Index)
            
            'Create new tab for the sub-action
            Call AddStuffForSubAction(ActionInfo)
            
            If ActionInfo.RunTimeStatus = "Running" Then
                frmMain.optRunAborted(Index + 1).Enabled = False
                frmMain.optRunMultiple(Index + 1).Enabled = False
                frmMain.optRunOnce(Index + 1).Enabled = False
                frmMain.optRunPeriod(Index + 1).Enabled = False
            End If
                        
            If RunningAction Then
                MainMod.FollowRunningAction = False
                Me.mnuEditItem(EditMenuItems.JumpToRunningAction).Enabled = True
            End If
'        End If
    ElseIf TypeName(ActionInfo) = "AutoFlatAction" Then
        If (ActionInfo.FlatLocation = DuskSkyFlat Or ActionInfo.FlatLocation = DawnSkyFlat) And RunningAction Then
            Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
        End If
        Call frmAutoFlat.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmAutoFlat.OKButton.Enabled = False
        Call frmAutoFlat.Show(1, Me)
        Call frmAutoFlat.PutFormDataIntoClass(ActionInfo)
        If CBool(frmAutoFlat.Tag) Then Modified = True
        Unload frmAutoFlat
        If RunningAction Then
            Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
        End If
    ElseIf TypeName(ActionInfo) = "SkipAheadAtTimeAction" Then
        If RunningAction Then
            Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
        End If
        Call frmSkipAtTime.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmSkipAtTime.OKButton.Enabled = False
        Call frmSkipAtTime.Show(1, Me)
        Call frmSkipAtTime.PutFormDataIntoClass(ActionInfo)
        If CBool(frmSkipAtTime.Tag) Then Modified = True
        Unload frmSkipAtTime
        If RunningAction Then
            Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
        End If
    ElseIf TypeName(ActionInfo) = "SkipAheadAtAltAction" Then
        If RunningAction Then
            Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
        End If
        Call frmSkipAtAlt.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmSkipAtAlt.OKButton.Enabled = False
        Call frmSkipAtAlt.Show(1, Me)
        Call frmSkipAtAlt.PutFormDataIntoClass(ActionInfo)
        If CBool(frmSkipAtAlt.Tag) Then Modified = True
        Unload frmSkipAtAlt
        If RunningAction Then
            Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
        End If
    ElseIf TypeName(ActionInfo) = "SkipAheadAtHAAction" Then
        If RunningAction Then
            Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
        End If
        Call frmSkipAtHA.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmSkipAtHA.OKButton.Enabled = False
        Call frmSkipAtHA.Show(1, Me)
        Call frmSkipAtHA.PutFormDataIntoClass(ActionInfo)
        If CBool(frmSkipAtHA.Tag) Then Modified = True
        Unload frmSkipAtHA
        If RunningAction Then
            Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
        End If
    ElseIf TypeName(ActionInfo) = "DomeAction" Then
        Call frmDomeAction.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmDomeAction.OKButton.Enabled = False
        Call frmDomeAction.Show(1, Me)
        Call frmDomeAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmDomeAction.Tag) Then Modified = True
        Unload frmDomeAction
    ElseIf TypeName(ActionInfo) = "CloudMonitorAction" Then
        Call frmCloudMonitorAction.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmCloudMonitorAction.OKButton.Enabled = False
        Call frmCloudMonitorAction.Show(1, Me)
        Call frmCloudMonitorAction.PutFormDataIntoClass(ActionInfo)
        If CBool(frmCloudMonitorAction.Tag) Then Modified = True
        Unload frmCloudMonitorAction
    ElseIf TypeName(ActionInfo) = "clsCommentAction" Then
        Call frmComment.GetFormDataFromClass(ActionInfo)
        If Not AllowEdits Then frmComment.OKButton.Enabled = False
        Call frmComment.Show(1, Me)
        Call frmComment.PutFormDataIntoClass(ActionInfo)
        If CBool(frmComment.Tag) Then Modified = True
        Unload frmComment
    End If

    If TypeName(ActionInfo) <> "RunActionList" Then
        Call ActionCollections(Index).Remove(Me.lstAction(Index).SelectedItem.Index)
        If ActionCollections(Index).Count = 0 Then
            Call ActionCollections(Index).Add(ActionInfo)
        ElseIf Me.lstAction(Index).SelectedItem.Index > 1 Then
            Call ActionCollections(Index).Add(ActionInfo, , , Me.lstAction(Index).SelectedItem.Index - 1)
        Else
            Call ActionCollections(Index).Add(ActionInfo, , Me.lstAction(Index).SelectedItem.Index)
        End If
        Me.lstAction(Index).ListItems(Me.lstAction(Index).SelectedItem.Index).Text = ActionInfo.BuildActionListString()
    End If
        
    EditingActionNumber = 0
    EditingActionListLevel = 0
    
    Call CheckModifiedFlag
End Sub

Private Sub lstAction_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim ActionInfo As Object
    Dim Counter As Long
    Dim Counter2 As Long
    
    Set ActionInfo = ActionCollections(Me.TabStrip.SelectedItem.Index).Item(Item.Index)
    
    If Not AllowSubActionEdit(Index) Then
        Beep
        Item.Checked = ActionInfo.Selected
        Exit Sub
    ElseIf Index > 1 Then
        If SubActionInfo(Index).LinkToFile Then
            Beep
            Item.Checked = ActionInfo.Selected
            Exit Sub
        Else
            ActionInfo.Selected = Item.Checked
            Modified = True
        End If
    Else
        ActionInfo.Selected = Item.Checked
        Modified = True
    End If
    
    If TypeName(ActionInfo) = "RunActionList" Then
        'Un/Checking the RunActionList action should check all the subactions, unless there is a link to file
        Call SetCheckStateAllActions(ActionInfo.ActionCollection, Item.Checked)
        
        'Update the check boxes on the screen if necessary
        If ActionCollections.Count > Me.TabStrip.SelectedItem.Index Then
            If ActionCollections(Me.TabStrip.SelectedItem.Index + 1) Is ActionInfo.ActionCollection Then
                For Counter = Me.TabStrip.SelectedItem.Index + 1 To Me.TabStrip.Tabs.Count - 1
                    For Counter2 = 1 To Me.lstAction(Counter).ListItems.Count
                        Me.lstAction(Counter).ListItems(Counter2).Checked = Item.Checked
                    Next Counter2
                Next Counter
            End If
        End If
    End If
    
    If RunningAction Then
        If Item.Checked Then
            If TypeName(ActionInfo) = "SkipAheadAtTimeAction" Or TypeName(ActionInfo) = "SkipAheadAtAltAction" Or TypeName(ActionInfo) = "SkipAheadAtHAAction" Then
                Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
            ElseIf TypeName(ActionInfo) = "AutoFlatAction" Then
                Call MainMod.AddSkipActionToSkipActionList(ActionInfo)
            ElseIf TypeName(ActionInfo) = "RunActionList" Then
                Call SearchActionListForSkipActions(ActionInfo.ActionCollection)
            End If
        Else
            If TypeName(ActionInfo) = "SkipAheadAtTimeAction" Or TypeName(ActionInfo) = "SkipAheadAtAltAction" Or TypeName(ActionInfo) = "SkipAheadAtHAAction" Then
                Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
            ElseIf TypeName(ActionInfo) = "AutoFlatAction" Then
                If ActionInfo.FlatLocation = DuskSkyFlat Or ActionInfo.FlatLocation = DawnSkyFlat Then
                    Call MainMod.RemoveSkipToTime(ActionInfo.ActualSkipToTime)
                End If
            ElseIf TypeName(ActionInfo) = "RunActionList" Then
                Call MainMod.SearchAndDeleteSkipActions(ActionInfo.ActionCollection)
            End If
        End If
    End If
    
    Call CheckModifiedFlag
End Sub

Private Sub lstAction_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DeleteActions
    End If
End Sub

Private Sub lstAction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuEdit
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
    Unload frmAbout
End Sub

Private Sub mnuActionItems_Click(Index As Integer)
    If Me.TabStrip.SelectedItem.Index = Me.TabStrip.Tabs.Count Then
        'on the status tab - move to the main setup tab
        Set TabStrip.SelectedItem = TabStrip.Tabs(1)
    End If
    Call AddAction(Index)
End Sub

Private Sub mnuEdit_Click()
    Dim myClip As New cCustomClipboard
    Dim FormatID As Long
    
    If frmMain.ActiveControl Is Nothing Then
        Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.CheckSelectedMenu).Enabled = False
        Me.mnuEditItem(EditMenuItems.UncheckSelectedMenu).Enabled = False
    Else
        If frmMain.ActiveControl.Name = "lstAction" Then
            If Not (frmMain.ActiveControl.SelectedItem Is Nothing) Then
                Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = True
                If Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked Then
                    Me.mnuEditItem(EditMenuItems.CheckSelectedMenu).Enabled = True
                    Me.mnuEditItem(EditMenuItems.UncheckSelectedMenu).Enabled = True
                End If
            Else
                Me.mnuEditItem(EditMenuItems.CopyMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.CutMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.PasteMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.DeleteMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.MoveUpMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.MoveDownMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.CheckSelectedMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.UncheckSelectedMenu).Enabled = False
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
            Me.mnuEditItem(EditMenuItems.CheckSelectedMenu).Enabled = False
            Me.mnuEditItem(EditMenuItems.UncheckSelectedMenu).Enabled = False
        End If
    End If
End Sub

Private Sub mnuEditItem_Click(Index As Integer)
    Dim Counter As Long
    Dim Counter2 As Long
    Dim clsAction As Object
    
    On Error GoTo mnuEditItemError
    
    Select Case Index
        Case EditMenuItems.CopyMenu
            'copy
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                Call CopyActions
            End If
        Case EditMenuItems.CutMenu
            'cut
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                Call CopyActions
                If Not (Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing) Then Call DeleteActions
            End If
        Case EditMenuItems.PasteMenu
            'paste
            Call PasteActions
        Case EditMenuItems.DeleteMenu
            'delete
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                Call DeleteActions
            End If
        Case EditMenuItems.MoveUpMenu
            'move up
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                Call MoveUp
            End If
        Case EditMenuItems.MoveDownMenu
            'move down
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                Call MoveDown
            End If
        Case EditMenuItems.UseCheckBoxes
            Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked = Not Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked

            If Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked Then
                Me.mnuEditItem(EditMenuItems.CheckAllMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.UncheckAllMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.CheckSelectedMenu).Enabled = True
                Me.mnuEditItem(EditMenuItems.UncheckSelectedMenu).Enabled = True

                Me.Toolbar.Buttons(ToolBarButtons.CheckSelectedBtn).Enabled = True
                Me.Toolbar.Buttons(ToolBarButtons.UncheckSelectedBtn).Enabled = True
                Me.Toolbar.Buttons(ToolBarButtons.CheckAllBtn).Enabled = True
                Me.Toolbar.Buttons(ToolBarButtons.UncheckAllBTn).Enabled = True
            Else
                Me.mnuEditItem(EditMenuItems.CheckAllMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.UncheckAllMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.CheckSelectedMenu).Enabled = False
                Me.mnuEditItem(EditMenuItems.UncheckSelectedMenu).Enabled = False

                Me.Toolbar.Buttons(ToolBarButtons.CheckSelectedBtn).Enabled = False
                Me.Toolbar.Buttons(ToolBarButtons.UncheckSelectedBtn).Enabled = False
                Me.Toolbar.Buttons(ToolBarButtons.CheckAllBtn).Enabled = False
                Me.Toolbar.Buttons(ToolBarButtons.UncheckAllBTn).Enabled = False
            End If

            For Counter = 1 To Me.TabStrip.Tabs.Count - 1
                If Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked Then
                    Me.lstAction(Counter).Checkboxes = True

                    'rebuild the list
                    Me.lstAction(Counter).ListItems.Clear
                    For Each clsAction In ActionCollections(Counter)
                        Call Me.lstAction(Counter).ListItems.Add(, , clsAction.BuildActionListString)
                        If clsAction.Selected Then
                            Me.lstAction(Counter).ListItems(Me.lstAction(Counter).ListItems.Count).Checked = True
                        Else
                            Me.lstAction(Counter).ListItems(Me.lstAction(Counter).ListItems.Count).Checked = False
                        End If
                    Next clsAction

                    Call MainMod.SaveMySetting("ProgramSettings", "Checkboxes", "True")
                Else
                    Me.lstAction(Counter).Checkboxes = False

                    Call MainMod.SaveMySetting("ProgramSettings", "Checkboxes", "False")
                End If
            Next Counter
        Case EditMenuItems.CheckAllMenu
            Call SetCheckStateAllActions(ActionCollections(1), True)
            
            For Counter = 1 To Me.TabStrip.Tabs.Count - 1
                For Counter2 = 1 To Me.lstAction(Counter).ListItems.Count
                    Me.lstAction(Counter).ListItems(Counter2).Checked = True
                Next Counter2
            Next Counter
            Modified = True
            Call CheckModifiedFlag
        Case EditMenuItems.UncheckAllMenu
            Call SetCheckStateAllActions(ActionCollections(1), False)
    
            For Counter = 1 To Me.TabStrip.Tabs.Count - 1
                For Counter2 = 1 To Me.lstAction(Counter).ListItems.Count
                    Me.lstAction(Counter).ListItems(Counter2).Checked = False
                Next Counter2
            Next Counter
            Modified = True
            Call CheckModifiedFlag
        Case EditMenuItems.CheckSelectedMenu
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
                    MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
                    Exit Sub
                End If
                
                For Counter = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
                    If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(Counter).Selected Then
                        Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(Counter).Checked = True
                        ActionCollections(Me.TabStrip.SelectedItem.Index).Item(Counter).Selected = True
                    End If
                Next Counter
                Modified = True
                Call CheckModifiedFlag
            End If
        Case EditMenuItems.UncheckSelectedMenu
            If Me.lstAction(Me.TabStrip.SelectedItem.Index).SelectedItem Is Nothing Then
                Beep
            Else
                If Not AllowSubActionEdit(Me.TabStrip.SelectedItem.Index) Then
                    MsgBox "You cannot edit a sub-action linked to a file." & vbCrLf & "Unlink to allow editing.", vbInformation + vbOKOnly
                    Exit Sub
                End If
                
                For Counter = 1 To Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems.Count
                    If Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(Counter).Selected Then
                        Me.lstAction(Me.TabStrip.SelectedItem.Index).ListItems(Counter).Checked = False
                        ActionCollections(Me.TabStrip.SelectedItem.Index).Item(Counter).Selected = False
                    End If
                Next Counter
                Modified = True
                Call CheckModifiedFlag
            End If
        Case EditMenuItems.JumpToRunningAction
            frmMain.mnuEditItem(EditMenuItems.JumpToRunningAction).Enabled = False
            Call MainMod.JumpToRunningAction(MainMod.colAction, 1)
            MainMod.FollowRunningAction = True
    End Select
    
    Exit Sub
    
mnuEditItemError:
    On Error GoTo 0
    
    'just beep for now - this is likely caused by clicking one of the icons while on the status tab
    Beep
End Sub

Private Sub SetCheckStateAllActions(myActionCollection As Collection, State As Boolean)
    Dim clsAction As Object
    
    For Each clsAction In myActionCollection
        clsAction.Selected = State
        If TypeName(clsAction) = "RunActionList" Then
            Call SetCheckStateAllActions(clsAction.ActionCollection, State)
        End If
    Next clsAction
End Sub


Private Sub mnuFileItem_Click(Index As Integer)
    Dim FileName As String
    Dim Reply As Long
    
    Select Case Index
        Case FileMenuItems.NewFileMenu
            If RunningAction Then
                Beep
                Exit Sub
            End If
            
            If Modified And Me.ActionCollections.Item(1).Count > 0 Then
                If CurrentFileName <> "" Then
                    Reply = MsgBox("Do you want to save the changes you made to " & CurrentFileName & "?", vbYesNoCancel + vbExclamation)
                Else
                    Reply = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbExclamation)
                End If
                
                If Reply = vbCancel Then
                    Exit Sub
                ElseIf Reply = vbYes Then
                    If CurrentFileName <> "" Then
                        Call SaveAction(CurrentFilePath, colAction)
                    Else
                        Call mnuFileItem_Click(FileMenuItems.SaveAsFileMenu)
                    End If
                End If
            End If
            
            Call MainMod.ClearAll

            'Me.lstAction(Me.TabStrip.SelectedItem.Index).Clear
            'Call MainMod.ClearCollection(ActionCollections(Me.TabStrip.SelectedItem.Index))
            
            CurrentFileName = ""
            CurrentFilePath = ""
            Me.Caption = "CCD Commander"
            Me.mnuFileItem(FileMenuItems.SaveFileMenu).Enabled = False
            Modified = False
        Case FileMenuItems.OpenFileMenu
            If Modified And Me.ActionCollections.Item(1).Count > 0 Then
                If CurrentFileName <> "" Then
                    Reply = MsgBox("Do you want to save the changes you made to " & CurrentFileName & "?", vbYesNoCancel + vbExclamation)
                Else
                    Reply = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbExclamation)
                End If
                
                If Reply = vbCancel Then
                    Exit Sub
                ElseIf Reply = vbYes Then
                    If CurrentFileName <> "" Then
                        Call SaveAction(CurrentFilePath, colAction)
                    Else
                        Call mnuFileItem_Click(FileMenuItems.SaveAsFileMenu)
                    End If
                End If
            End If
            
            On Error Resume Next
            Call MkDir(App.Path & "\Actions")
            On Error GoTo 0
            Me.ComDlg.DialogTitle = "Load Action List"
            Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
            Me.ComDlg.FilterIndex = 1
            Me.ComDlg.CancelError = True
            Me.ComDlg.InitDir = App.Path & "\Actions\"
            Me.ComDlg.flags = cdlOFNFileMustExist
            On Error Resume Next
            Me.ComDlg.ShowOpen
            If Err.Number = 0 Then
                On Error GoTo 0
                Me.MousePointer = vbHourglass
                On Error Resume Next
                Call LoadAction(Me.ComDlg.FileName)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Me.MousePointer = vbNormal
                    Call MsgBox("Error reading file " & Mid(Me.ComDlg.FileName, InStrRev(Me.ComDlg.FileName, "\") + 1), vbOKOnly + vbCritical, "CCD Commander")
                Else
                    On Error GoTo 0
                    Me.MousePointer = vbNormal
                    CurrentFilePath = Me.ComDlg.FileName
                    CurrentFileName = Mid(Me.ComDlg.FileName, InStrRev(CurrentFilePath, "\") + 1)
                    Me.Caption = "CCD Commander - " & CurrentFileName
                    Me.mnuFileItem(FileMenuItems.SaveFileMenu).Enabled = True
                    Modified = False
                End If
            Else
                On Error GoTo 0
            End If
        Case FileMenuItems.SaveFileMenu
            If CurrentFileName <> "" And CurrentFilePath <> "" Then
                Call SaveAction(CurrentFilePath, colAction)
                Modified = False
                Me.Caption = "CCD Commander - " & CurrentFileName
            Else
                On Error Resume Next
                Call MkDir(App.Path & "\Actions")
                On Error GoTo 0
                Me.ComDlg.DialogTitle = "Save Action List"
                Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
                Me.ComDlg.FilterIndex = 1
                If CurrentFilePath <> "" Then
                    Me.ComDlg.InitDir = CurrentFilePath
                Else
                    Me.ComDlg.InitDir = App.Path & "\Actions\"
                End If
                If CurrentFileName <> "" Then
                    Me.ComDlg.FileName = CurrentFileName
                Else
                    Me.ComDlg.FileName = Format(Now, "yymmdd") & ".act"
                End If
                Me.ComDlg.CancelError = True
                Me.ComDlg.flags = cdlOFNOverwritePrompt
                On Error Resume Next
                Me.ComDlg.ShowSave
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Call SaveAction(Me.ComDlg.FileName, colAction)
                    CurrentFilePath = Me.ComDlg.FileName
                    CurrentFileName = Mid(Me.ComDlg.FileName, InStrRev(CurrentFilePath, "\") + 1)
                    Me.Caption = "CCD Commander - " & CurrentFileName
                    Me.mnuFileItem(FileMenuItems.SaveFileMenu).Enabled = True
                    Modified = False
                Else
                    On Error GoTo 0
                End If
            End If
        Case FileMenuItems.SaveAsFileMenu
            On Error Resume Next
            Call MkDir(App.Path & "\Actions")
            On Error GoTo 0
            Me.ComDlg.DialogTitle = "Save Action List"
            Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
            Me.ComDlg.FilterIndex = 1
            If CurrentFilePath <> "" Then
                Me.ComDlg.InitDir = CurrentFilePath
            Else
                Me.ComDlg.InitDir = App.Path & "\Actions\"
            End If
            If CurrentFileName <> "" Then
                Me.ComDlg.FileName = CurrentFileName
            Else
                Me.ComDlg.FileName = Format(Now, "yymmdd") & ".act"
            End If
            Me.ComDlg.CancelError = True
            Me.ComDlg.flags = cdlOFNOverwritePrompt
            On Error Resume Next
            Me.ComDlg.ShowSave
            If Err.Number = 0 Then
                On Error GoTo 0
                Call SaveAction(Me.ComDlg.FileName, colAction)
                CurrentFilePath = Me.ComDlg.FileName
                CurrentFileName = Mid(Me.ComDlg.FileName, InStrRev(CurrentFilePath, "\") + 1)
                Me.Caption = "CCD Commander - " & CurrentFileName
                Me.mnuFileItem(FileMenuItems.SaveFileMenu).Enabled = True
                Modified = False
            Else
                On Error GoTo 0
            End If
        Case FileMenuItems.ImportTargetListMenu
            frmImport.Show 1, Me
        Case FileMenuItems.ExitMenu
            Unload Me
        Case FileMenuItems.OpenPreviousFile1, FileMenuItems.OpenPreviousFile2, FileMenuItems.OpenPreviousFile3, FileMenuItems.OpenPreviousFile4, FileMenuItems.OpenPreviousFile5
            If Modified And Me.ActionCollections.Item(1).Count > 0 Then
                If CurrentFileName <> "" Then
                    Reply = MsgBox("Do you want to save the changes you made to " & CurrentFileName & "?", vbYesNoCancel + vbExclamation)
                Else
                    Reply = MsgBox("Do you want to save your changes?", vbYesNoCancel + vbExclamation)
                End If
                
                If Reply = vbCancel Then
                    Exit Sub
                ElseIf Reply = vbYes Then
                    If CurrentFileName <> "" Then
                        Call SaveAction(CurrentFilePath, colAction)
                    Else
                        Call mnuFileItem_Click(FileMenuItems.SaveAsFileMenu)
                    End If
                End If
            End If
            
            Me.MousePointer = vbHourglass
            On Error Resume Next
            Call LoadAction(Me.mnuFileItem(Index).Tag)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Me.MousePointer = vbNormal
                If MsgBox("Error opening file " & Me.mnuFileItem(Index).Caption & vbCrLf & "Do you want to remove it from the list?", vbYesNo + vbExclamation, "CCD Commander") = vbYes Then
                    Call MainMod.RemoveFileFromPreviousFileList(Index - FileMenuItems.OpenPreviousFile1 + 1)
                End If
            Else
                On Error GoTo 0
                Me.MousePointer = vbNormal
                'The LoadAction call will put the selected file at position 1, so reference that position now
                CurrentFilePath = Me.mnuFileItem(FileMenuItems.OpenPreviousFile1).Tag
                CurrentFileName = Mid(Me.mnuFileItem(FileMenuItems.OpenPreviousFile1).Tag, InStrRev(CurrentFilePath, "\") + 1)
                Me.Caption = "CCD Commander - " & CurrentFileName
                Me.mnuFileItem(FileMenuItems.SaveFileMenu).Enabled = True
                Modified = False
            End If
    End Select
End Sub

Private Sub mnuHelpItem_Click(Index As Integer)
    Select Case Index
        Case 0
            Call HtmlHelp(Me.hwnd, App.HelpFile, 15, 1)
        Case 1
            Call HtmlHelp(Me.hwnd, App.HelpFile, 15, 2)
        Case 2
            Call HtmlHelp(Me.hwnd, App.HelpFile, 15, 3)
    End Select
End Sub

Private Sub mnuRunItem_Click(Index As Integer)
    Select Case Index
        Case 0
            If Paused Then
                Call PauseAndResumeAction
            Else
                RunSelectedActionsOnly = Me.mnuEditItem(EditMenuItems.UseCheckBoxes).Checked
                Call StartAndAbortAction
            End If
        Case 1
            Call PauseAndResumeAction
        Case 2
            Call StartAndAbortAction
    End Select
End Sub

Private Sub mnuSetup_Click()
    If Me.RunningAction Then
        frmOptions.SSTab1.TabEnabled(0) = False
        If frmOptions.SSTab1.Tab = 0 Then
            frmOptions.SSTab1.Tab = 4
        End If
    Else
        frmOptions.SSTab1.TabEnabled(0) = True
    End If
    Call frmOptions.Show(1, Me)
End Sub

Private Sub mnuViewItem_Click(Index As Integer)
    Select Case Index
        Case 0
            If Me.mnuViewItem(Index).Checked = True Then
                Me.mnuViewItem(Index).Checked = False
                Me.mnuViewItem(1).Enabled = False
                Me.mnuViewItem(1).Checked = False
                frmTempGraph.Visible = False
                frmTempGraph.Timer1.Enabled = False
                Call MainMod.SaveMySetting("GraphOptions", "EnableTempRecording", False)
            Else
                Me.mnuViewItem(Index).Checked = True
                Me.mnuViewItem(1).Enabled = True
                If Me.RunningAction Then
                    frmTempGraph.Timer1.Enabled = True
                End If
                Call MainMod.SaveMySetting("GraphOptions", "EnableTempRecording", True)
            End If
        Case 1
            If Me.mnuViewItem(Index).Checked = True Then
                frmTempGraph.Visible = False
                Me.mnuViewItem(Index).Checked = False
            Else
                frmTempGraph.Visible = True
                Me.mnuViewItem(Index).Checked = True
            End If
        Case 2
            If Me.mnuViewItem(Index).Checked = True Then
                Me.mnuViewItem(Index).Checked = False
                Me.mnuViewItem(3).Enabled = False
                Me.mnuViewItem(3).Checked = False
                frmAutoguiderError.Visible = False
                frmAutoguiderError.Timer1.Enabled = False
                Call MainMod.SaveMySetting("GraphOptions", "EnableAutoguiderRecording", False)
            Else
                Me.mnuViewItem(Index).Checked = True
                Me.mnuViewItem(3).Enabled = True
                frmAutoguiderError.Timer1.Enabled = True
                Call MainMod.SaveMySetting("GraphOptions", "EnableAutoguiderRecording", True)
            End If
        Case 3
            If Me.mnuViewItem(Index).Checked = True Then
                frmAutoguiderError.Visible = False
                Me.mnuViewItem(Index).Checked = False
            Else
                frmAutoguiderError.Visible = True
                Me.mnuViewItem(Index).Checked = True
            End If
    End Select
End Sub

Private Sub mnuWindowItem_Click(Index As Integer)
    If mnuWindowItem(0).Checked = False Then
        'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        mnuWindowItem(0).Checked = True
        Call MainMod.SaveMySetting("WindowPositions", "AlwaysOnTop", True)
        
        Call MainMod.SetOnTopMode(Me)
        Call MainMod.SetOnTopMode(frmOptions)
    Else
        'SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        mnuWindowItem(0).Checked = False
        Call MainMod.SaveMySetting("WindowPositions", "AlwaysOnTop", False)
    
        Call MainMod.SetOnTopMode(Me)
        Call MainMod.SetOnTopMode(frmOptions)
    End If
End Sub

Private Sub optRunAborted_Click(Index As Integer)
    Me.txtTime(Index).Enabled = False
    Me.txtNumRepeat(Index).Enabled = False
    SubActionInfo(Index).RepeatMode = 3

    If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
    Else
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
    End If
End Sub

Private Sub optRunMultiple_Click(Index As Integer)
    Me.txtNumRepeat(Index).Enabled = True
    Me.txtTime(Index).Enabled = False
    SubActionInfo(Index).RepeatMode = 1
    
    If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
    Else
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
    End If
End Sub

Private Sub optRunOnce_Click(Index As Integer)
    Me.txtTime(Index).Enabled = False
    Me.txtNumRepeat(Index).Enabled = False
    SubActionInfo(Index).RepeatMode = 0

    If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
    Else
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
    End If
End Sub

Private Sub optRunPeriod_Click(Index As Integer)
    Me.txtNumRepeat(Index).Enabled = False
    Me.txtTime(Index).Enabled = True
    SubActionInfo(Index).RepeatMode = 2

    If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
    Else
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
    End If
End Sub

Private Sub TabStrip_Click()
    If Me.TabStrip.SelectedItem.Index = myCurrentTab Then
        Exit Sub
        'no need to change
    End If
    
    If Me.TabStrip.SelectedItem.Index = Me.TabStrip.Tabs.Count Then
        Me.fraStatus.Visible = True
        Me.HelpContextID = 110
    Else
        Me.fraTabs(Me.TabStrip.SelectedItem.Index).Visible = True
        If Me.TabStrip.SelectedItem.Index = 1 Then
            Me.HelpContextID = 100
        Else
            Me.HelpContextID = 1400
        End If
    End If
    
    If myCurrentTab = Me.TabStrip.Tabs.Count Then
        Me.fraStatus.Visible = False
    Else
        Me.fraTabs(myCurrentTab).Visible = False
    End If
    
    myCurrentTab = Me.TabStrip.SelectedItem.Index
End Sub

Private Sub tmrRunning_Timer()
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
    Dim SimParkCheckInterval As Double
    Dim MinSimParkCheckInterval As Double
    Dim TimeNow As Double
        
    tmrRunning.Enabled = False
        
    On Error GoTo tmrRunningError
    
    'Check/Update Autoguider Error if needed
    Call Camera.CheckAutoguiderError
    
    TimeNow = CDbl(Timer)
    
    If TimeNow > (StartTime + 0.5) Or _
        (DateDiff("d", Date, StartDate) <> 0 And _
        (TimeNow + 86400) > (StartTime + 0.5)) Then
        
        StartTime = Timer
        StartDate = Now
            
        Me.shpRunning.Visible = Not Me.shpRunning.Visible
        
        If DateDiff("s", LastCloudSensorCheck, Now) > (Settings.CloudMonitorQueryPeriod * 60) And frmOptions.lstCloudSensor.ListIndex <> WeatherMonitorControl.None Then
            LastCloudSensorCheck = Now
            Call CloudSensor.CheckCloudSensor
        End If
        
        If DateDiff("s", LastSkipToCheck, Now) > 5 And Not DisableSkipChecking Then
            LastSkipToCheck = Now
            Call MainMod.CheckSkipToTimes
        End If
        
        MinSimParkCheckInterval = CDbl(GetMySetting("CustomSetting", "SimulatedParkCheckInterval", "10"))
        If Settings.DelayAfterSlew > MinSimParkCheckInterval Then
            SimParkCheckInterval = Settings.DelayAfterSlew
        Else
            SimParkCheckInterval = MinSimParkCheckInterval
        End If
        
        If DateDiff("s", LastSimParkCheck, Now) > SimParkCheckInterval Then
            LastSimParkCheck = Now
            
            If Mount.SimulatedPark Then
                'goto a set alt-az coordinate, repeatedly go to the same coordinates
                Call Mount.RecenterAltAz(Mount.SimulatedParkAlt, Mount.SimulatedParkAzim)
            End If
        End If
    End If
    
tmrRunningError:
    ErrorNum = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
        
    On Error GoTo 0
    
    tmrRunning.Enabled = True
    
    If ErrorNum = &H80010005 Then
        'error is automation error - I can just ignore this and try again next time around
    ElseIf ErrorNum <> 0 Then
        Call Err.Raise(ErrorNum, ErrorSource, ErrorDescription)
    End If
End Sub

Private Sub CheckModifiedFlag()
    If Modified And CurrentFileName <> "" And InStr(Me.Caption, "*") = 0 Then
        Me.Caption = Me.Caption & "*"
    End If
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
        Case ToolBarButtons.CheckSelectedBtn
            'check selected
            Call mnuEditItem_Click(EditMenuItems.CheckSelectedMenu)
        Case ToolBarButtons.UncheckSelectedBtn
            'uncheck selected
            Call mnuEditItem_Click(EditMenuItems.UncheckSelectedMenu)
        Case ToolBarButtons.CheckAllBtn
            Call mnuEditItem_Click(EditMenuItems.CheckAllMenu)
        Case ToolBarButtons.UncheckAllBTn
            Call mnuEditItem_Click(EditMenuItems.UncheckAllMenu)
        Case ToolBarButtons.PlayBtn
            'play
            Call mnuRunItem_Click(RunMenuItems.StartMenu)
        Case ToolBarButtons.PauseBtn
            'pause
            Call mnuRunItem_Click(RunMenuItems.PauseMenu)
        Case ToolBarButtons.StopBtn
            'stop
            Call mnuRunItem_Click(RunMenuItems.StopMenu)
    End Select
End Sub

Private Sub txtFileName_Validate(Index As Integer, Cancel As Boolean)
    If SubActionInfo(Index).ActionListName <> Me.txtFileName(Index).Text Then
        SubActionInfo(Index).ActionListName = Me.txtFileName(Index).Text
        Modified = True
    End If
    
    If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
    Else
        Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
    End If
End Sub

Private Sub txtNumRepeat_GotFocus(Index As Integer)
    Me.txtNumRepeat(Index).SelStart = 0
    Me.txtNumRepeat(Index).SelLength = Len(Me.txtNumRepeat(Index).Text)
End Sub

Private Sub txtNumRepeat_Validate(Index As Integer, Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtNumRepeat(Index).Text)
    If Err.Number <> 0 Or Test < 2 Or Test <> Me.txtNumRepeat(Index).Text Then
        Beep
        Cancel = True
    Else
        SubActionInfo(Index).TimesToRepeat = Test
        If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
            Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
        Else
            Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub txtTime_GotFocus(Index As Integer)
    Me.txtTime(Index).SelStart = 0
    Me.txtTime(Index).SelLength = Len(Me.txtTime(Index).Text)
End Sub

Private Sub txtTime_Validate(Index As Integer, Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtTime(Index).Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtTime(Index).Text Then
        Beep
        Cancel = True
    Else
        Me.txtTime(Index).Text = Format(Me.txtTime(Index).Text, "0.00")
        SubActionInfo(Index).RepeatTime = Test
        If InStr(Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text, "- Running") > 0 Then
            Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString & " - Running "
        Else
            Me.lstAction(Index - 1).ListItems(UpperActionIndexes(Index)).Text = SubActionInfo(Index).BuildActionListString
        End If
    End If
    On Error GoTo 0
End Sub

Private Function CheckIfFileExists(ByVal FileName As String) As String
    Dim actFileNum As Integer
    Dim FileSize As Long
    
    actFileNum = FreeFile
    On Error Resume Next
    Open FileName For Binary Access Read As #actFileNum
    FileSize = LOF(actFileNum)
    Close actFileNum
    On Error GoTo 0
    
    If FileSize = 0 Then
        'file doesn't exist, prompt for the file
        On Error Resume Next
        Kill FileName
        On Error GoTo 0
        
        If MsgBox("Could not find:" & vbCrLf & FileName & vbCrLf & vbCrLf & "Do you want to try and find the file?" & vbCrLf & "Answering no will unlink the above file from the sub-action.", vbYesNo + vbCritical) = vbYes Then
            Me.ComDlg.Filter = "CCDCommander Action (*.act)|*.act"
            Me.ComDlg.FilterIndex = 1
            Me.ComDlg.CancelError = True
            Me.ComDlg.InitDir = FileName
            Me.ComDlg.flags = cdlOFNFileMustExist
            Me.ComDlg.FileName = FileName
            On Error Resume Next
            Me.ComDlg.ShowOpen
            If Err.Number = 0 Then
                On Error GoTo 0
                Modified = True
                Call CheckModifiedFlag
                CheckIfFileExists = Me.ComDlg.FileName
            Else
                On Error GoTo 0
                CheckIfFileExists = ""
            End If
        Else
            CheckIfFileExists = ""
        End If
    Else
        CheckIfFileExists = FileName
    End If
End Function

Public Sub UpdatePreviousFileMenu()
    Dim FileName As String
    
    FileName = GetMySetting("PreviousFiles", "PreviousFile1", "")
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile1).Caption = Mid(FileName, InStrRev(FileName, "\") + 1)
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile1).Tag = FileName
    If FileName <> "" Then
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile1).Visible = True
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFileDivider).Visible = True
    Else
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile1).Visible = False
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFileDivider).Visible = False
    End If
    
    FileName = GetMySetting("PreviousFiles", "PreviousFile2", "")
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile2).Caption = Mid(FileName, InStrRev(FileName, "\") + 1)
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile2).Tag = FileName
    If FileName <> "" Then
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile2).Visible = True
    Else
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile2).Visible = False
    End If

    FileName = GetMySetting("PreviousFiles", "PreviousFile3", "")
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile3).Caption = Mid(FileName, InStrRev(FileName, "\") + 1)
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile3).Tag = FileName
    If FileName <> "" Then
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile3).Visible = True
    Else
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile3).Visible = False
    End If
    
    FileName = GetMySetting("PreviousFiles", "PreviousFile4", "")
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile4).Caption = Mid(FileName, InStrRev(FileName, "\") + 1)
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile4).Tag = FileName
    If FileName <> "" Then
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile4).Visible = True
    Else
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile4).Visible = False
    End If
    
    FileName = GetMySetting("PreviousFiles", "PreviousFile5", "")
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile5).Caption = Mid(FileName, InStrRev(FileName, "\") + 1)
    frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile5).Tag = FileName
    If FileName <> "" Then
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile5).Visible = True
    Else
        frmMain.mnuFileItem(FileMenuItems.OpenPreviousFile5).Visible = False
    End If
End Sub

