VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAutoFlat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatic Flat Action"
   ClientHeight    =   7515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8955
   HelpContextID   =   1500
   Icon            =   "frmAutoFlat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRotations 
      Caption         =   "Rotations"
      Height          =   3015
      Left            =   3660
      TabIndex        =   76
      Top             =   3360
      Width           =   2595
      Begin VB.CheckBox chkFlipRotator 
         Caption         =   "Flip rotator 180 degrees"
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   2700
         Width           =   2355
      End
      Begin VB.TextBox txtPositionAngle 
         Height          =   252
         Left            =   1320
         TabIndex        =   29
         Text            =   "0"
         Top             =   420
         Width           =   495
      End
      Begin VB.CommandButton cmdMoveDownRotation 
         Caption         =   "Move Down"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   675
      End
      Begin VB.CommandButton cmdMoveUpRotation 
         Caption         =   "Move Up"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   1380
         Width           =   675
      End
      Begin VB.CommandButton cmdDeleteRotation 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1440
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdAddRotation 
         Caption         =   "Add"
         Height          =   435
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox lstPositionAngle 
         Height          =   1230
         Left            =   840
         TabIndex        =   34
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "degrees"
         Height          =   195
         Left            =   1860
         TabIndex        =   78
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Position Angle:"
         Height          =   195
         Left            =   225
         TabIndex        =   77
         Top             =   450
         Width           =   1050
      End
   End
   Begin VB.CheckBox chkContinuousAdjust 
      Caption         =   "Continuously adjust exposure time to match ADU target."
      Height          =   195
      Left            =   60
      TabIndex        =   41
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mount Information"
      Height          =   3015
      Left            =   60
      TabIndex        =   59
      Top             =   3360
      Width           =   3555
      Begin VB.Frame fraSunAltitude 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1155
         Left            =   1740
         TabIndex        =   80
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         Begin VB.TextBox txtDawnSunAlt 
            Height          =   285
            Left            =   120
            TabIndex        =   83
            Text            =   "-7.00"
            Top             =   840
            Width           =   675
         End
         Begin VB.TextBox txtDuskSunAlt 
            Height          =   285
            Left            =   120
            TabIndex        =   82
            Text            =   "-0.83"
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "degrees"
            Height          =   195
            Left            =   840
            TabIndex        =   85
            Top             =   900
            Width           =   570
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "degrees"
            Height          =   195
            Left            =   840
            TabIndex        =   84
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Sun Altitude:"
            Height          =   195
            Left            =   60
            TabIndex        =   81
            Top             =   60
            Width           =   900
         End
      End
      Begin VB.OptionButton optFlatPosition 
         Caption         =   "Do not connect to or move mount"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   2580
         Width           =   2835
      End
      Begin VB.CommandButton cmdComputeTime 
         Caption         =   "Compute Dusk Start Time"
         Height          =   555
         Left            =   1800
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton optFlatPosition 
         Caption         =   "Dawn Sky Flat"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optFlatPosition 
         Caption         =   "Park Mount"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2235
      End
      Begin VB.OptionButton optFlatPosition 
         Caption         =   "Dusk Sky Flat"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1515
      End
      Begin VB.OptionButton optFlatPosition 
         Caption         =   "Slew to:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1740
         Width           =   915
      End
      Begin VB.Frame fraAltAz 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   960
         TabIndex        =   60
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
         Begin VB.TextBox txtAzimS 
            Height          =   312
            Left            =   1980
            TabIndex        =   26
            Text            =   "15"
            Top             =   420
            Width           =   312
         End
         Begin VB.TextBox txtAzimM 
            Height          =   312
            Left            =   1320
            TabIndex        =   25
            Text            =   "23"
            Top             =   420
            Width           =   312
         End
         Begin VB.TextBox txtAzimD 
            Height          =   312
            Left            =   600
            TabIndex        =   24
            Text            =   "5"
            Top             =   420
            Width           =   435
         End
         Begin VB.TextBox txtAltS 
            Height          =   312
            Left            =   1980
            TabIndex        =   23
            Text            =   "58"
            Top             =   60
            Width           =   312
         End
         Begin VB.TextBox txtAltM 
            Height          =   312
            Left            =   1320
            TabIndex        =   22
            Text            =   "10"
            Top             =   60
            Width           =   312
         End
         Begin VB.TextBox txtAltD 
            Height          =   312
            Left            =   600
            TabIndex        =   21
            Text            =   "23"
            Top             =   60
            Width           =   435
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Azim:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "s"
            Height          =   195
            Left            =   2340
            TabIndex        =   68
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "m"
            Height          =   195
            Left            =   1680
            TabIndex        =   67
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "d"
            Height          =   195
            Left            =   1140
            TabIndex        =   66
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "s"
            Height          =   195
            Left            =   2340
            TabIndex        =   64
            Top             =   120
            Width           =   90
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "m"
            Height          =   195
            Left            =   1680
            TabIndex        =   63
            Top             =   120
            Width           =   135
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "d"
            Height          =   195
            Left            =   1140
            TabIndex        =   62
            Top             =   120
            Width           =   90
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Alt:"
            Height          =   195
            Left            =   240
            TabIndex        =   61
            Top             =   120
            Width           =   225
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flat Information"
      Height          =   3315
      Left            =   60
      TabIndex        =   47
      Top             =   60
      Width           =   6855
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
         Left            =   6480
         MaskColor       =   &H00D8E9EC&
         Picture         =   "frmAutoFlat.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "File Name Builder"
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   275
      End
      Begin VB.TextBox txtDarkExposureTimeTolerance 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   87
         Text            =   "1.00"
         Top             =   2940
         Width           =   552
      End
      Begin VB.TextBox txtNumberDarksPerFlat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Text            =   "10"
         Top             =   2940
         Width           =   552
      End
      Begin VB.CheckBox chkMatchingDarks 
         Alignment       =   1  'Right Justify
         Caption         =   "Take Matching Darks"
         Height          =   312
         Left            =   720
         TabIndex        =   9
         Top             =   2640
         Width           =   1875
      End
      Begin VB.Frame Frame5 
         Caption         =   "Flat Frame Size"
         Height          =   1095
         Left            =   4020
         TabIndex        =   75
         Top             =   1260
         Width           =   2775
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Half Frame"
            Height          =   252
            Index           =   1
            Left            =   780
            TabIndex        =   15
            Top             =   480
            Width           =   1152
         End
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Quarter Frame"
            Height          =   252
            Index           =   2
            Left            =   540
            TabIndex        =   16
            Top             =   720
            Width           =   1392
         End
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Frame"
            Height          =   252
            Index           =   0
            Left            =   780
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   1152
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Exposure Setup Frame Size"
         Height          =   1095
         Left            =   4020
         TabIndex        =   74
         Top             =   180
         Width           =   2775
         Begin VB.OptionButton optSetupFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Half Frame"
            Height          =   252
            Index           =   1
            Left            =   780
            TabIndex        =   12
            Top             =   480
            Width           =   1152
         End
         Begin VB.OptionButton optSetupFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Quarter Frame"
            Height          =   252
            Index           =   2
            Left            =   540
            TabIndex        =   13
            Top             =   720
            Width           =   1392
         End
         Begin VB.OptionButton optSetupFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Frame"
            Height          =   252
            Index           =   0
            Left            =   780
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1152
         End
      End
      Begin VB.TextBox txtMaxADU 
         Height          =   252
         Left            =   2400
         TabIndex        =   3
         Text            =   "20000"
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox txtNumberOfExposures 
         Height          =   240
         Left            =   2400
         TabIndex        =   4
         Text            =   "10"
         Top             =   1240
         Width           =   552
      End
      Begin VB.TextBox txtDelay 
         Enabled         =   0   'False
         Height          =   252
         Left            =   2400
         TabIndex        =   5
         Text            =   "0"
         Top             =   1500
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtAverageADU 
         Height          =   252
         Left            =   2400
         TabIndex        =   2
         Text            =   "20000"
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox txtMaxExposureTime 
         Height          =   252
         Left            =   2400
         TabIndex        =   1
         Text            =   "60"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtFileNamePrefix 
         Height          =   264
         Left            =   2400
         TabIndex        =   8
         Top             =   2400
         Width           =   4035
      End
      Begin VB.ComboBox cmbBin 
         Height          =   315
         ItemData        =   "frmAutoFlat.frx":0456
         Left            =   2400
         List            =   "frmAutoFlat.frx":0466
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1740
         Width           =   912
      End
      Begin VB.TextBox txtExposureTime 
         Height          =   252
         Left            =   2400
         TabIndex        =   0
         Text            =   "1"
         Top             =   240
         Width           =   915
      End
      Begin VB.CheckBox chkAutosave 
         Alignment       =   1  'Right Justify
         Caption         =   "Autosave Exposures"
         Height          =   312
         Left            =   780
         TabIndex        =   7
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1812
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   6480
         TabIndex        =   88
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dark Exposure Time Tolerance:"
         Height          =   195
         Left            =   3540
         TabIndex        =   86
         Top             =   3000
         Width           =   2250
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Darks per Flat:"
         Height          =   195
         Left            =   555
         TabIndex        =   79
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "ADU"
         Height          =   195
         Left            =   3360
         TabIndex        =   73
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum Average ADU Value:"
         Height          =   195
         Left            =   60
         TabIndex        =   72
         Top             =   1020
         Width           =   2295
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ADU"
         Height          =   195
         Left            =   3360
         TabIndex        =   58
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimum Average ADU Value:"
         Height          =   195
         Left            =   60
         TabIndex        =   57
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   3360
         TabIndex        =   56
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum Exposure Time:"
         Height          =   195
         Left            =   60
         TabIndex        =   55
         Top             =   540
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Filename Prefix:"
         Height          =   195
         Left            =   1260
         TabIndex        =   54
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Exposures:"
         Height          =   195
         Left            =   780
         TabIndex        =   53
         Top             =   1260
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Bin:"
         Height          =   195
         Left            =   2055
         TabIndex        =   52
         Top             =   1770
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   3360
         TabIndex        =   51
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimum Exposure Time:"
         Height          =   195
         Left            =   60
         TabIndex        =   50
         Top             =   255
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Delay Before Exposure:"
         Height          =   195
         Left            =   690
         TabIndex        =   49
         Top             =   1500
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   3360
         TabIndex        =   48
         Top             =   1500
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7380
      TabIndex        =   46
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7380
      TabIndex        =   45
      Top             =   180
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filters"
      Height          =   3015
      Left            =   6300
      TabIndex        =   69
      Top             =   3360
      Width           =   2595
      Begin VB.CommandButton cmdReverse 
         Caption         =   "Reverse Order"
         Height          =   315
         Left            =   840
         TabIndex        =   89
         Top             =   2640
         Width           =   1275
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move Down"
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   675
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move Up"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   1380
         Width           =   675
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   435
         Left            =   1440
         TabIndex        =   37
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   435
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox lstFilter 
         Height          =   1230
         Left            =   840
         TabIndex        =   40
         Top             =   1320
         Width           =   1275
      End
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         ItemData        =   "frmAutoFlat.frx":047E
         Left            =   1065
         List            =   "frmAutoFlat.frx":0491
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   420
         Width           =   912
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Filter:"
         Height          =   195
         Left            =   600
         TabIndex        =   70
         Top             =   450
         Width           =   435
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
      Left            =   7080
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmAutoFlat.frx":04B2
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7140
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.TextBox txtSaveTo 
      Height          =   315
      Left            =   1680
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "C:\"
      Top             =   7080
      Width           =   5295
   End
   Begin VB.CheckBox chkUseGlobalImageSaveLocation 
      Caption         =   "Use Global Image Save Location"
      Height          =   312
      Left            =   60
      TabIndex        =   42
      Top             =   6780
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog MSComm 
      Left            =   7680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Image Save Location:"
      Height          =   195
      Left            =   60
      TabIndex        =   71
      Top             =   7140
      Width           =   1560
   End
End
Attribute VB_Name = "frmAutoFlat"
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

Private Const RegistryName = "AutoFlatAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Public Sub GetSettings()
    Dim FilterList As String
    Dim NewFilter As Double
    
    Me.txtExposureTime.Text = GetMySetting(RegistryName, "MinExposureTime", Format(1, "0.000"))
    Me.txtMaxExposureTime.Text = GetMySetting(RegistryName, "MaxExposureTime", Format(60, "0.000"))
    Me.txtAverageADU.Text = GetMySetting(RegistryName, "AverageADU", "20000")
    Me.txtNumberOfExposures.Text = GetMySetting(RegistryName, "NumberOfExpsoures", "10")
    Me.txtDelay.Text = GetMySetting(RegistryName, "DelayTime", Format(0, "0.0"))
    Me.cmbBin.ListIndex = CInt(GetMySetting(RegistryName, "ImagerBin", "0"))
    If CInt(GetMySetting(RegistryName, "Filter", "0")) < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = CInt(GetMySetting(RegistryName, "Filter", "0"))
    End If
    Me.chkAutosave.Value = CInt(GetMySetting(RegistryName, "Autosave", "1"))
    Me.txtFileNamePrefix.Text = GetMySetting(RegistryName, "FileNamePrefix", "")
    
    Call optFlatPosition_Click(CInt(GetMySetting(RegistryName, "MountPosition", "1")))
    
    Me.txtAltD.Text = GetMySetting(RegistryName, "AltD", "0")
    Me.txtAltM.Text = GetMySetting(RegistryName, "AltM", "0")
    Me.txtAltS.Text = GetMySetting(RegistryName, "AltS", "0")
    Me.txtAzimD.Text = GetMySetting(RegistryName, "AzimD", "270")
    Me.txtAzimM.Text = GetMySetting(RegistryName, "AzimM", "0")
    Me.txtAzimS.Text = GetMySetting(RegistryName, "AzimS", "0")

    FilterList = GetMySetting(RegistryName, "FilterList", "0,1,2,3,")
    Me.lstFilter.Clear
    Do While Len(FilterList) > 0
        NewFilter = CInt(Left(FilterList, InStr(FilterList, ",") - 1))
        If NewFilter < Me.cmbFilter.ListCount Then
            Me.lstFilter.AddItem Me.cmbFilter.List(NewFilter)
            Me.lstFilter.ItemData(Me.lstFilter.NewIndex) = NewFilter
        End If
        FilterList = Mid(FilterList, InStr(FilterList, ",") + 1)
    Loop

    Me.chkContinuousAdjust.Value = CInt(GetMySetting(RegistryName, "ContinuousAdjust", "0"))
    
    Me.chkUseGlobalImageSaveLocation.Value = CInt(GetMySetting(RegistryName, "UseGlobalImageSaveLocation", "1"))
    Me.txtSaveTo.Text = GetMySetting(RegistryName, "SaveToPath", frmOptions.txtSaveTo.Text)
    
    Call chkUseGlobalImageSaveLocation_Click

    Me.optImageFrameSize(CInt(GetMySetting(RegistryName, "ImageFrameSize", "0"))).Value = True
       
    Me.optSetupFrameSize(CInt(GetMySetting(RegistryName, "SetupFrameSize", "0"))).Value = True
    Me.txtMaxADU.Text = CLng(GetMySetting(RegistryName, "MaximumADU", CLng(Me.txtAverageADU.Text) * 1.1))
    Me.chkMatchingDarks.Value = CInt(GetMySetting(RegistryName, "TakeMatchingDarks", "0"))
    Me.txtNumberDarksPerFlat.Text = CInt(GetMySetting(RegistryName, "NumberDarksPerFlat", "1"))
    If Me.txtNumberDarksPerFlat.Text = "0" Then
        Me.txtNumberDarksPerFlat.Text = "1"
    End If
    
    FilterList = GetMySetting(RegistryName, "RotationList", "")
    Me.lstPositionAngle.Clear
    Do While Len(FilterList) > 0
        NewFilter = CInt(Left(FilterList, InStr(FilterList, ",") - 1))
        If NewFilter <= 360 Then
            Me.lstPositionAngle.AddItem NewFilter
        End If
        FilterList = Mid(FilterList, InStr(FilterList, ",") + 1)
    Loop
    
    Me.txtDarkExposureTimeTolerance.Text = GetMySetting(RegistryName, "DarkExposureTimeTolerance", Format(1, "0.00"))
    Me.txtDuskSunAlt.Text = GetMySetting(RegistryName, "DuskSunAlt", Format(-0.83, "0.00"))
    Me.txtDawnSunAlt.Text = GetMySetting(RegistryName, "DawnSunAlt", Format(-7#, "0.00"))
    
    Me.chkFlipRotator.Value = CInt(GetMySetting(RegistryName, "FlipRotator", "0"))
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

Private Sub chkMatchingDarks_Click()
    If Me.chkMatchingDarks.Value = vbChecked Then
        Me.txtNumberDarksPerFlat.Enabled = True
        Me.txtDarkExposureTimeTolerance.Enabled = True
    Else
        Me.txtNumberDarksPerFlat.Enabled = False
        Me.txtDarkExposureTimeTolerance.Enabled = False
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

Private Sub cmdReverse_Click()
    Dim FilterName As New Collection
    Dim FilterNumber As New Collection
    Dim Counter As Integer
    
    For Counter = 0 To lstFilter.ListCount - 1
        FilterName.Add lstFilter.List(Counter)
        FilterNumber.Add lstFilter.ItemData(Counter)
    Next Counter
    
    lstFilter.Clear
    
    For Counter = FilterName.Count To 1 Step -1
        Me.lstFilter.AddItem FilterName(Counter)
        Me.lstFilter.ItemData(Me.lstFilter.NewIndex) = FilterNumber(Counter)
    Next Counter
End Sub

Private Sub cmdAdd_Click()
    If Me.cmbFilter.ListIndex = -1 Then
        Call MsgBox("You must select an existing filter to add it to the list.")
    ElseIf Me.lstFilter.ListCount = MaxAutoFlatFilters Then
        Call MsgBox("Maximum of " & MaxAutoFlatFilters & " reached." & vbCrLf & "Please delete a filter first.")
    Else
        Call lstFilter.AddItem(Me.cmbFilter.List(Me.cmbFilter.ListIndex))
        lstFilter.ItemData(Me.lstFilter.NewIndex) = Me.cmbFilter.ListIndex
    End If
End Sub

Private Sub cmdAddRotation_Click()
    If Me.lstPositionAngle.ListCount = Camera.MaxAutoFlatRotations Then
        Call MsgBox("Maximum of " & Camera.MaxAutoFlatRotations & " reached." & vbCrLf & "Please delete a position angle first.")
    Else
        Call Me.lstPositionAngle.AddItem(Me.txtPositionAngle.Text)
    End If
End Sub

Private Sub cmdComputeTime_Click()
    If Me.optFlatPosition(1).Value Then
        Call Mount.ComputeSunSetTime(CDbl(Me.txtDuskSunAlt.Text))
        MsgBox "Dusk begins at " & Mount.SunSetTime & ".", vbInformation
    Else
        Call Mount.ComputeTwilightStartTime(Me.txtDawnSunAlt.Text)
        
        MsgBox "Dawn begins at " & Mount.TwilightStartTime & ".", vbInformation
    End If
End Sub

Private Sub cmdDelete_Click()
    If Me.lstFilter.ListIndex = -1 Then
        Call MsgBox("You must select a filter from the list.")
    Else
        Call Me.lstFilter.RemoveItem(Me.lstFilter.ListIndex)
    End If
End Sub

Private Sub cmdDeleteRotation_Click()
    If Me.lstPositionAngle.ListIndex = -1 Then
        Call MsgBox("You must select a position angle from the list.")
    Else
        Call Me.lstPositionAngle.RemoveItem(Me.lstPositionAngle.ListIndex)
    End If
End Sub

Private Sub cmdMoveDown_Click()
    Dim FilterName As String
    Dim FilterData As Integer
    Dim NewIndex As Integer
    
    If Me.lstFilter.ListIndex = -1 Then
        Call MsgBox("You must select a filter from the list.")
    ElseIf Me.lstFilter.ListIndex = Me.lstFilter.ListCount - 1 Then
        Call MsgBox("Cannot move any lower in the list.")
    Else
        FilterName = Me.lstFilter.List(Me.lstFilter.ListIndex)
        FilterData = Me.lstFilter.ItemData(Me.lstFilter.ListIndex)
        NewIndex = Me.lstFilter.ListIndex + 1
        Call Me.lstFilter.RemoveItem(Me.lstFilter.ListIndex)
        Call Me.lstFilter.AddItem(FilterName, NewIndex)
        Me.lstFilter.ItemData(NewIndex) = FilterData
        Me.lstFilter.ListIndex = NewIndex
    End If
End Sub

Private Sub cmdMoveDownRotation_Click()
    Dim PositionAngle As String
    Dim NewIndex As Integer
    
    If Me.lstPositionAngle.ListIndex = -1 Then
        Call MsgBox("You must select a position angle from the list.")
    ElseIf Me.lstPositionAngle.ListIndex = Me.lstPositionAngle.ListCount - 1 Then
        Call MsgBox("Cannot move any lower in the list.")
    Else
        PositionAngle = Me.lstPositionAngle.List(Me.lstPositionAngle.ListIndex)
        NewIndex = Me.lstPositionAngle.ListIndex + 1
        Call Me.lstPositionAngle.RemoveItem(Me.lstPositionAngle.ListIndex)
        Call Me.lstPositionAngle.AddItem(PositionAngle, NewIndex)
        Me.lstPositionAngle.ListIndex = NewIndex
    End If
End Sub

Private Sub cmdMoveUp_Click()
    Dim FilterName As String
    Dim FilterData As Integer
    Dim NewIndex As Integer
    
    If Me.lstFilter.ListIndex = -1 Then
        Call MsgBox("You must select a filter from the list.")
    ElseIf Me.lstFilter.ListIndex = 0 Then
        Call MsgBox("Cannot move any higher in the list.")
    Else
        FilterName = Me.lstFilter.List(Me.lstFilter.ListIndex)
        FilterData = Me.lstFilter.ItemData(Me.lstFilter.ListIndex)
        NewIndex = Me.lstFilter.ListIndex - 1
        Call Me.lstFilter.RemoveItem(Me.lstFilter.ListIndex)
        Call Me.lstFilter.AddItem(FilterName, NewIndex)
        Me.lstFilter.ItemData(NewIndex) = FilterData
        Me.lstFilter.ListIndex = NewIndex
    End If
End Sub

Private Sub cmdMoveUpRotation_Click()
    Dim PositionAngle As String
    Dim NewIndex As Integer
    
    If Me.lstPositionAngle.ListIndex = -1 Then
        Call MsgBox("You must select a position angle from the list.")
    ElseIf Me.lstPositionAngle.ListIndex = 0 Then
        Call MsgBox("Cannot move any higher in the list.")
    Else
        PositionAngle = Me.lstPositionAngle.List(Me.lstPositionAngle.ListIndex)
        NewIndex = Me.lstPositionAngle.ListIndex - 1
        Call Me.lstPositionAngle.RemoveItem(Me.lstPositionAngle.ListIndex)
        Call Me.lstPositionAngle.AddItem(PositionAngle, NewIndex)
        Me.lstPositionAngle.ListIndex = NewIndex
    End If
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
    
    'setup Filter Info
    Me.cmbFilter.Clear
    For Counter = 0 To frmOptions.lstFilters.ListCount - 1
        Call Me.cmbFilter.AddItem(frmOptions.lstFilters.List(Counter))
    Next Counter
    
    Call GetSettings
    
    If frmOptions.lstRotator.ListIndex = RotatorControl.None Then
        Me.lstPositionAngle.Clear
        Me.fraRotations.Enabled = False
    Else
        Me.fraRotations.Enabled = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Tag = "False"
        Me.Hide
        Call GetSettings
    End If
End Sub

Public Sub PutFormDataIntoClass(clsAction As AutoFlatAction)
    Dim Counter As Long
    
    clsAction.AutosaveExposure = Me.chkAutosave.Value
    clsAction.AverageADU = CLng(Me.txtAverageADU.Text)
    clsAction.Bin = Me.cmbBin.ListIndex
    clsAction.DelayTime = CDbl(Me.txtDelay.Text)
    clsAction.FileNamePrefix = Me.txtFileNamePrefix.Text
    clsAction.Filter = Me.cmbFilter.ListIndex
    clsAction.MaxExpTime = CDbl(Me.txtMaxExposureTime.Text)
    clsAction.MinExpTime = CDbl(Me.txtExposureTime.Text)
    clsAction.NumExp = CInt(Me.txtNumberOfExposures.Text)
    If Me.optFlatPosition(0).Value Then
        clsAction.FlatLocation = FlatParkMount
    ElseIf Me.optFlatPosition(1).Value Then
        clsAction.FlatLocation = DuskSkyFlat
    ElseIf Me.optFlatPosition(2).Value Then
        clsAction.FlatLocation = DawnSkyFlat
    ElseIf Me.optFlatPosition(3).Value Then
        clsAction.FlatLocation = FixedLocation
    ElseIf Me.optFlatPosition(4).Value Then
        clsAction.FlatLocation = DoNotMove
    End If
    
    clsAction.AltD = CDbl(Me.txtAltD.Text)
    clsAction.AltM = CDbl(Me.txtAltM.Text)
    clsAction.AltS = CDbl(Me.txtAltS.Text)
    
    clsAction.AzimD = CDbl(Me.txtAzimD.Text)
    clsAction.AzimM = CDbl(Me.txtAzimM.Text)
    clsAction.AzimS = CDbl(Me.txtAzimS.Text)
    
    If Me.optImageFrameSize(0).Value Then
        clsAction.FrameSize = FullFrame
    ElseIf Me.optImageFrameSize(1).Value Then
        clsAction.FrameSize = HalfFrame
    ElseIf Me.optImageFrameSize(2).Value Then
        clsAction.FrameSize = QuarterFrame
    End If
    
    clsAction.NumFilters = Me.lstFilter.ListCount
    If clsAction.NumFilters > 0 Then
        For Counter = 0 To clsAction.NumFilters - 1
            clsAction.Filters(Counter) = Me.lstFilter.ItemData(Counter)
        Next Counter
    End If

    clsAction.ContinuouslyAdjust = Me.chkContinuousAdjust.Value
    
    clsAction.UseGlobalSaveToLocation = Me.chkUseGlobalImageSaveLocation.Value
    clsAction.FileSavePath = Me.txtSaveTo.Text
    
    clsAction.MaximumADU = CLng(Me.txtMaxADU.Text)
    clsAction.TakeMatchingDarks = Me.chkMatchingDarks.Value
    clsAction.NumberOfDarksPerFlat = CInt(Me.txtNumberDarksPerFlat.Text)
    
    If Me.optSetupFrameSize(0).Value Then
        clsAction.SetupFrameSize = FullFrame
    ElseIf Me.optSetupFrameSize(1).Value Then
        clsAction.SetupFrameSize = HalfFrame
    ElseIf Me.optSetupFrameSize(2).Value Then
        clsAction.SetupFrameSize = QuarterFrame
    End If
    
    clsAction.NumRotations = Me.lstPositionAngle.ListCount
    If clsAction.NumRotations > 0 Then
        For Counter = 0 To clsAction.NumRotations - 1
            clsAction.Rotations(Counter) = Me.lstPositionAngle.List(Counter)
        Next Counter
    End If
    
    clsAction.DarkFrameTolerance = CDbl(Me.txtDarkExposureTimeTolerance.Text)
    clsAction.DuskSunAltitudeStart = CDbl(Me.txtDuskSunAlt.Text)
    clsAction.DawnSunAltitudeStart = CDbl(Me.txtDawnSunAlt.Text)
    
    clsAction.FlipRotator = Me.chkFlipRotator.Value
End Sub

Public Sub GetFormDataFromClass(clsAction As AutoFlatAction)
    Dim FilterList As String
    Dim Counter As Long
    
    Me.chkAutosave.Value = clsAction.AutosaveExposure
    Me.txtAverageADU.Text = clsAction.AverageADU
    Me.cmbBin.ListIndex = clsAction.Bin
    Me.txtDelay.Text = Format(clsAction.DelayTime, "0.0")
    Me.txtFileNamePrefix.Text = clsAction.FileNamePrefix
    If (clsAction.Filter < Me.cmbFilter.ListCount) Then
        Me.cmbFilter.ListIndex = clsAction.Filter
    End If
    Me.txtMaxExposureTime.Text = Format(clsAction.MaxExpTime, "0.000")
    Me.txtExposureTime.Text = Format(clsAction.MinExpTime, "0.000")
    Me.txtNumberOfExposures.Text = clsAction.NumExp
    
    If clsAction.FlatLocation = FlatParkMount Then
        Me.optFlatPosition(0).Value = True
    ElseIf clsAction.FlatLocation = DuskSkyFlat Then
        Me.optFlatPosition(1).Value = True
    ElseIf clsAction.FlatLocation = DawnSkyFlat Then
        Me.optFlatPosition(2).Value = True
    ElseIf clsAction.FlatLocation = FixedLocation Then
        Me.optFlatPosition(3).Value = True
    ElseIf clsAction.FlatLocation = DoNotMove Then
        Me.optFlatPosition(4).Value = True
    End If
    
    Me.txtAltD.Text = clsAction.AltD
    Me.txtAltM.Text = clsAction.AltM
    Me.txtAltS.Text = clsAction.AltS
    Me.txtAzimD.Text = clsAction.AzimD
    Me.txtAzimM.Text = clsAction.AzimM
    Me.txtAzimS.Text = clsAction.AzimS
    
    If clsAction.FrameSize = FullFrame Then
        Me.optImageFrameSize(0).Value = True
    ElseIf clsAction.FrameSize = HalfFrame Then
         Me.optImageFrameSize(1).Value = True
    ElseIf clsAction.FrameSize = QuarterFrame Then
        Me.optImageFrameSize(2).Value = True
    End If
        
    Me.lstFilter.Clear
    For Counter = 0 To clsAction.NumFilters - 1
        If clsAction.Filters(Counter) < Me.cmbFilter.ListCount Then
            Call Me.lstFilter.AddItem(Me.cmbFilter.List(clsAction.Filters(Counter)))
            Me.lstFilter.ItemData(Me.lstFilter.NewIndex) = clsAction.Filters(Counter)
        End If
    Next Counter
    
    Me.chkContinuousAdjust.Value = clsAction.ContinuouslyAdjust
    
    Me.chkUseGlobalImageSaveLocation.Value = clsAction.UseGlobalSaveToLocation
    If Me.chkUseGlobalImageSaveLocation.Value = vbChecked Then
        Me.txtSaveTo.Text = frmOptions.txtSaveTo.Text
    Else
        Me.txtSaveTo.Text = clsAction.FileSavePath
    End If
    
    Call chkUseGlobalImageSaveLocation_Click

    Me.txtMaxADU.Text = clsAction.MaximumADU
    Me.chkMatchingDarks.Value = clsAction.TakeMatchingDarks
    Me.txtNumberDarksPerFlat.Text = clsAction.NumberOfDarksPerFlat
    
    If clsAction.SetupFrameSize = FullFrame Then
        Me.optSetupFrameSize(0).Value = True
    ElseIf clsAction.SetupFrameSize = HalfFrame Then
        Me.optSetupFrameSize(1).Value = True
    ElseIf clsAction.SetupFrameSize = QuarterFrame Then
        Me.optSetupFrameSize(2).Value = True
    End If
    
    Me.lstPositionAngle.Clear
    For Counter = 0 To clsAction.NumRotations - 1
        Call Me.lstPositionAngle.AddItem(clsAction.Rotations(Counter))
    Next Counter
    
    Me.txtDarkExposureTimeTolerance.Text = Format(clsAction.DarkFrameTolerance, "0.00")
    Me.txtDuskSunAlt.Text = Format(clsAction.DuskSunAltitudeStart, "0.00")
    Me.txtDawnSunAlt.Text = Format(clsAction.DawnSunAltitudeStart, "0.00")
    
    Me.chkFlipRotator.Value = clsAction.FlipRotator
    
    'save the data into the registry just in case the user cancels
    Call SaveMySetting(RegistryName, "MinExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "MaxExposureTime", Me.txtMaxExposureTime.Text)
    Call SaveMySetting(RegistryName, "AverageADU", Me.txtAverageADU.Text)
    Call SaveMySetting(RegistryName, "NumberOfExpsoures", Me.txtNumberOfExposures.Text)
    Call SaveMySetting(RegistryName, "DelayTime", Me.txtDelay.Text)
    Call SaveMySetting(RegistryName, "ImagerBin", Me.cmbBin.ListIndex)
    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    Call SaveMySetting(RegistryName, "Autosave", Me.chkAutosave.Value)
    Call SaveMySetting(RegistryName, "FileNamePrefix", Me.txtFileNamePrefix.Text)
    If Me.optFlatPosition(0).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "0")
    ElseIf Me.optFlatPosition(1).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "1")
    ElseIf Me.optFlatPosition(2).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "2")
    ElseIf Me.optFlatPosition(3).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "3")
    ElseIf Me.optFlatPosition(4).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "4")
    End If
    Call SaveMySetting(RegistryName, "AltD", Me.txtAltD.Text)
    Call SaveMySetting(RegistryName, "AltM", Me.txtAltM.Text)
    Call SaveMySetting(RegistryName, "AltS", Me.txtAltS.Text)
    Call SaveMySetting(RegistryName, "AzimD", Me.txtAzimD.Text)
    Call SaveMySetting(RegistryName, "AzimM", Me.txtAzimM.Text)
    Call SaveMySetting(RegistryName, "AzimS", Me.txtAzimS.Text)
    
    If Me.optImageFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "0")
    ElseIf Me.optImageFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "1")
    ElseIf Me.optImageFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "2")
    End If

    FilterList = ""
    For Counter = 0 To Me.lstFilter.ListCount - 1
        FilterList = FilterList & Me.lstFilter.ItemData(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "FilterList", FilterList)

    Call SaveMySetting(RegistryName, "ContinuousAdjust", Me.chkContinuousAdjust.Value)

    Call SaveMySetting(RegistryName, "UseGlobalImageSaveLocation", Me.chkUseGlobalImageSaveLocation.Value)
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)

    If Me.optSetupFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "SetupFrameSize", 0)
    ElseIf Me.optSetupFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "SetupFrameSize", 1)
    ElseIf Me.optSetupFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "SetupFrameSize", 2)
    End If
    Call SaveMySetting(RegistryName, "MaximumADU", Me.txtMaxADU.Text)
    Call SaveMySetting(RegistryName, "TakeMatchingDarks", Me.chkMatchingDarks.Value)
    Call SaveMySetting(RegistryName, "NumberDarksPerFlat", Me.txtNumberDarksPerFlat.Text)
    
    FilterList = ""
    For Counter = 0 To Me.lstPositionAngle.ListCount - 1
        FilterList = FilterList & Me.lstPositionAngle.List(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "RotationList", FilterList)

    Call SaveMySetting(RegistryName, "DarkExposureTimeTolerance", Me.txtDarkExposureTimeTolerance.Text)
    Call SaveMySetting(RegistryName, "DuskSunAlt", Me.txtDuskSunAlt.Text)
    Call SaveMySetting(RegistryName, "DawnSunAlt", Me.txtDawnSunAlt.Text)
    
    Call SaveMySetting(RegistryName, "FlipRotator", Me.chkFlipRotator.Value)
End Sub

Private Sub OKButton_Click()
    Dim Counter As Integer
    Dim FilterList As String
    Dim Cancel As Boolean
    Cancel = False
    
    If Not Cancel Then Call txtExposureTime_Validate(Cancel)
    If Not Cancel Then Call txtMaxExposureTime_Validate(Cancel)
    If Not Cancel Then Call txtAverageADU_Validate(Cancel)
    If Not Cancel Then Call txtNumberOfExposures_Validate(Cancel)
    If Not Cancel Then Call txtDelay_Validate(Cancel)
    If Not Cancel Then Call txtAltD_Validate(Cancel)
    If Not Cancel Then Call txtAltM_Validate(Cancel)
    If Not Cancel Then Call txtAltS_Validate(Cancel)
    If Not Cancel Then Call txtAzimD_Validate(Cancel)
    If Not Cancel Then Call txtAzimM_Validate(Cancel)
    If Not Cancel Then Call txtAzimS_Validate(Cancel)
    If Not Cancel Then Call txtMaxADU_Validate(Cancel)
    If Not Cancel Then Call txtNumberDarksPerFlat_Validate(Cancel)
    If Not Cancel Then Call txtDarkExposureTimeTolerance_Validate(Cancel)
    If Not Cancel Then Call txtDuskSunAlt_Validate(Cancel)
    If Not Cancel Then Call txtDawnSunAlt_Validate(Cancel)
    
    If (CInt(Me.txtAltD.Text) < 0) And (Me.optFlatPosition(3).Value) Then
        Cancel = (MsgBox("Slew To Altitude is negative!" & vbCrLf & "Are you sure this is what you want?", vbYesNoCancel + vbCritical, Me.Caption) <> vbYes)
    End If

    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "MinExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "MaxExposureTime", Me.txtMaxExposureTime.Text)
    Call SaveMySetting(RegistryName, "AverageADU", Me.txtAverageADU.Text)
    Call SaveMySetting(RegistryName, "NumberOfExpsoures", Me.txtNumberOfExposures.Text)
    Call SaveMySetting(RegistryName, "DelayTime", Me.txtDelay.Text)
    Call SaveMySetting(RegistryName, "ImagerBin", Me.cmbBin.ListIndex)
    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    Call SaveMySetting(RegistryName, "Autosave", Me.chkAutosave.Value)
    Call SaveMySetting(RegistryName, "FileNamePrefix", Me.txtFileNamePrefix.Text)
    If Me.optFlatPosition(0).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "0")
    ElseIf Me.optFlatPosition(1).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "1")
    ElseIf Me.optFlatPosition(2).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "2")
    ElseIf Me.optFlatPosition(3).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "3")
    ElseIf Me.optFlatPosition(4).Value Then
        Call SaveMySetting(RegistryName, "MountPosition", "4")
    End If
    Call SaveMySetting(RegistryName, "AltD", Me.txtAltD.Text)
    Call SaveMySetting(RegistryName, "AltM", Me.txtAltM.Text)
    Call SaveMySetting(RegistryName, "AltS", Me.txtAltS.Text)
    Call SaveMySetting(RegistryName, "AzimD", Me.txtAzimD.Text)
    Call SaveMySetting(RegistryName, "AzimM", Me.txtAzimM.Text)
    Call SaveMySetting(RegistryName, "AzimS", Me.txtAzimS.Text)
    
    If Me.optImageFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "0")
    ElseIf Me.optImageFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "1")
    ElseIf Me.optImageFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "2")
    End If
    
    FilterList = ""
    For Counter = 0 To Me.lstFilter.ListCount - 1
        FilterList = FilterList & Me.lstFilter.ItemData(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "FilterList", FilterList)
    
    Call SaveMySetting(RegistryName, "ContinuousAdjust", Me.chkContinuousAdjust.Value)
    
    Call SaveMySetting(RegistryName, "UseGlobalImageSaveLocation", Me.chkUseGlobalImageSaveLocation.Value)
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)
    
    If Me.optSetupFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "SetupFrameSize", 0)
    ElseIf Me.optSetupFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "SetupFrameSize", 1)
    ElseIf Me.optSetupFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "SetupFrameSize", 2)
    End If
    Call SaveMySetting(RegistryName, "MaximumADU", Me.txtMaxADU.Text)
    Call SaveMySetting(RegistryName, "TakeMatchingDarks", Me.chkMatchingDarks.Value)
    Call SaveMySetting(RegistryName, "NumberDarksPerFlat", Me.txtNumberDarksPerFlat.Text)
    
    FilterList = ""
    For Counter = 0 To Me.lstPositionAngle.ListCount - 1
        FilterList = FilterList & Me.lstPositionAngle.List(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "RotationList", FilterList)
    
    Call SaveMySetting(RegistryName, "DarkExposureTimeTolerance", Me.txtDarkExposureTimeTolerance.Text)
    Call SaveMySetting(RegistryName, "DuskSunAlt", Me.txtDuskSunAlt.Text)
    Call SaveMySetting(RegistryName, "DawnSunAlt", Me.txtDawnSunAlt.Text)
    
    Call SaveMySetting(RegistryName, "FlipRotator", Me.chkFlipRotator.Value)
    
    Me.Tag = "True"
    Me.Hide
End Sub

Private Sub optFlatPosition_Click(Index As Integer)
    Me.optFlatPosition(Index).Value = True

    If Me.optFlatPosition(1).Value Then
        Me.cmdComputeTime.Caption = "Compute Dusk Start Time"
        Me.cmdComputeTime.Visible = True
        Me.fraSunAltitude.Visible = True
    ElseIf Me.optFlatPosition(2).Value Then
        Me.cmdComputeTime.Caption = "Compute Dawn Start Time"
        Me.cmdComputeTime.Visible = True
        Me.fraSunAltitude.Visible = True
    Else
        Me.cmdComputeTime.Visible = False
        Me.fraSunAltitude.Visible = False
    End If
    
    If Me.optFlatPosition(3).Value Then
        Me.fraAltAz.Visible = True
    Else
        Me.fraAltAz.Visible = False
    End If
End Sub

Private Sub optImageFrameSize_Click(Index As Integer)
    Me.optImageFrameSize(Index).Value = True
    
    If Index = 0 Then
        Me.optImageFrameSize(1).Value = False
        Me.optImageFrameSize(2).Value = False
    ElseIf Index = 1 Then
        Me.optImageFrameSize(0).Value = False
        Me.optImageFrameSize(2).Value = False
    ElseIf Index = 2 Then
        Me.optImageFrameSize(0).Value = False
        Me.optImageFrameSize(1).Value = False
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
    Else
        Me.txtDelay.Text = Format(Me.txtDelay.Text, "0.0")
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
    If Err.Number <> 0 Or Test < 0.001 Or Test <> Me.txtExposureTime.Text Then
        Beep
        Cancel = True
    Else
        Me.txtExposureTime.Text = Format(Me.txtExposureTime.Text, "0.000")
    End If
    On Error GoTo 0
End Sub

Private Sub txtFileNamePrefix_GotFocus()
    Me.txtFileNamePrefix.SelStart = 0
    Me.txtFileNamePrefix.SelLength = Len(Me.txtFileNamePrefix.Text)
End Sub

Private Sub txtMaxADU_GotFocus()
    Me.txtMaxADU.SelStart = 0
    Me.txtMaxADU.SelLength = Len(Me.txtMaxADU.Text)
End Sub

Private Sub txtMaxADU_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CLng(Me.txtMaxADU.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxADU.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxExposureTime_GotFocus()
    Me.txtMaxExposureTime.SelStart = 0
    Me.txtMaxExposureTime.SelLength = Len(Me.txtMaxExposureTime.Text)
End Sub

Private Sub txtMaxExposureTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMaxExposureTime.Text)
    If Err.Number <> 0 Or Test < 0.001 Or Test <> Me.txtMaxExposureTime.Text Then
        Beep
        Cancel = True
    Else
        Me.txtMaxExposureTime.Text = Format(Me.txtMaxExposureTime.Text, "0.000")
    End If
    On Error GoTo 0
End Sub

Private Sub txtAverageADU_GotFocus()
    Me.txtAverageADU.SelStart = 0
    Me.txtAverageADU.SelLength = Len(Me.txtAverageADU.Text)
End Sub

Private Sub txtAverageADU_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CLng(Me.txtAverageADU.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtAverageADU.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtNumberDarksPerFlat_GotFocus()
    Me.txtNumberDarksPerFlat.SelStart = 0
    Me.txtNumberDarksPerFlat.SelLength = Len(Me.txtNumberDarksPerFlat.Text)
End Sub

Private Sub txtNumberDarksPerFlat_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtNumberDarksPerFlat.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtNumberDarksPerFlat.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtNumberOfExposures_GotFocus()
    Me.txtNumberOfExposures.SelStart = 0
    Me.txtNumberOfExposures.SelLength = Len(Me.txtNumberOfExposures.Text)
End Sub

Private Sub txtNumberOfExposures_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtNumberOfExposures.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtNumberOfExposures.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAltD_GotFocus()
    Me.txtAltD.SelStart = 0
    Me.txtAltD.SelLength = Len(Me.txtAltD.Text)
End Sub

Private Sub txtAltD_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAltD.Text)
    If Err.Number <> 0 Or Test < -90 Or Test >= 90 Or Test <> Me.txtAltD.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAltM_GotFocus()
    Me.txtAltM.SelStart = 0
    Me.txtAltM.SelLength = Len(Me.txtAltM.Text)
End Sub

Private Sub txtAltM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAltM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAltM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAltS_GotFocus()
    Me.txtAltS.SelStart = 0
    Me.txtAltS.SelLength = Len(Me.txtAltS.Text)
End Sub

Private Sub txtAltS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAltS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAltS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAzimD_GotFocus()
    Me.txtAzimD.SelStart = 0
    Me.txtAzimD.SelLength = Len(Me.txtAzimD.Text)
End Sub

Private Sub txtAzimD_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimD.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 360 Or Test <> Me.txtAzimD.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAzimM_GotFocus()
    Me.txtAzimM.SelStart = 0
    Me.txtAzimM.SelLength = Len(Me.txtAzimM.Text)
End Sub

Private Sub txtAzimM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimM.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAzimM.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtAzimS_GotFocus()
    Me.txtAzimS.SelStart = 0
    Me.txtAzimS.SelLength = Len(Me.txtAzimS.Text)
End Sub

Private Sub txtAzimS_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtAzimS.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 60 Or Test <> Me.txtAzimS.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtPositionAngle_GotFocus()
    Me.txtPositionAngle.SelStart = 0
    Me.txtPositionAngle.SelLength = Len(Me.txtPositionAngle.Text)
End Sub

Private Sub txtPositionAngle_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtPositionAngle.Text)
    If Err.Number <> 0 Or Test < 0 Or Test >= 360 Or Test <> Me.txtPositionAngle.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtDarkExposureTimeTolerance_GotFocus()
    Me.txtDarkExposureTimeTolerance.SelStart = 0
    Me.txtDarkExposureTimeTolerance.SelLength = Len(Me.txtDarkExposureTimeTolerance.Text)
End Sub

Private Sub txtDarkExposureTimeTolerance_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDarkExposureTimeTolerance.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtDarkExposureTimeTolerance.Text Then
        Beep
        Cancel = True
    Else
        Me.txtDarkExposureTimeTolerance.Text = Format(Me.txtDarkExposureTimeTolerance.Text, "0.0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtDuskSunAlt_GotFocus()
    Me.txtDuskSunAlt.SelStart = 0
    Me.txtDuskSunAlt.SelLength = Len(Me.txtDuskSunAlt.Text)
End Sub

Private Sub txtDuskSunAlt_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDuskSunAlt.Text)
    If Err.Number <> 0 Or Test <> Me.txtDuskSunAlt.Text Then
        Beep
        Cancel = True
    Else
        Me.txtDuskSunAlt.Text = Format(Me.txtDuskSunAlt.Text, "0.00")
    End If
    On Error GoTo 0
End Sub

Private Sub txtDawnSunAlt_GotFocus()
    Me.txtDawnSunAlt.SelStart = 0
    Me.txtDawnSunAlt.SelLength = Len(Me.txtDawnSunAlt.Text)
End Sub

Private Sub txtDawnSunAlt_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDawnSunAlt.Text)
    If Err.Number <> 0 Or Test <> Me.txtDawnSunAlt.Text Then
        Beep
        Cancel = True
    Else
        Me.txtDawnSunAlt.Text = Format(Me.txtDawnSunAlt.Text, "0.00")
    End If
    On Error GoTo 0
End Sub


