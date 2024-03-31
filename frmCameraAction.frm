VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCameraAction 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Take Image Action"
   ClientHeight    =   10125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7410
   HelpContextID   =   1000
   Icon            =   "frmCameraAction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSyncToCurrent 
      Caption         =   "Sync to Current RA/Dec When Complete"
      Height          =   255
      Left            =   3300
      TabIndex        =   95
      Top             =   4740
      Width           =   4035
   End
   Begin VB.CheckBox chkDoubleImageLink 
      Caption         =   "Perform two Plate Solves after Meridian Flip"
      Enabled         =   0   'False
      Height          =   312
      Left            =   3420
      TabIndex        =   89
      Top             =   6300
      Width           =   3675
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
      Left            =   7020
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmCameraAction.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   86
      ToolTipText     =   "File Name Builder"
      Top             =   5970
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.TextBox txtFileNamePrefix 
      Height          =   285
      Left            =   1800
      TabIndex        =   83
      Top             =   5940
      Width           =   5175
   End
   Begin VB.CheckBox chkAutosave 
      Caption         =   "Autosave Exposures"
      Height          =   312
      Left            =   60
      TabIndex        =   84
      Top             =   5700
      Value           =   1  'Checked
      Width           =   1812
   End
   Begin VB.Frame fraDither 
      Caption         =   "Dither Information"
      Height          =   1515
      Left            =   3240
      TabIndex        =   72
      Top             =   3180
      Width           =   4095
      Begin VB.CheckBox chkYAxisDither 
         Caption         =   "Y Axis Dither"
         Height          =   252
         Left            =   2640
         TabIndex        =   94
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkXAxisDither 
         Caption         =   "X Axis Dither"
         Height          =   252
         Left            =   2640
         TabIndex        =   93
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkMaintainDitherOnFilterChange 
         Caption         =   "Maintain Dither Position With Filter Change"
         Height          =   255
         Left            =   300
         TabIndex        =   82
         Top             =   1200
         Width           =   3315
      End
      Begin VB.TextBox txtDitherFrequency 
         Height          =   285
         Left            =   1410
         TabIndex        =   75
         Text            =   "1"
         Top             =   300
         Width           =   372
      End
      Begin VB.TextBox txtDitherStep 
         Height          =   285
         Left            =   1410
         TabIndex        =   74
         Text            =   "4"
         Top             =   600
         Width           =   500
      End
      Begin VB.TextBox txtDitherAmount 
         Height          =   285
         Left            =   1410
         TabIndex        =   73
         Text            =   "4"
         Top             =   870
         Width           =   500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "exposures"
         Height          =   195
         Left            =   1815
         TabIndex        =   81
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dither Frequency:"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label lblDitherStepTerm 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   1980
         TabIndex        =   79
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dither Step:"
         Height          =   195
         Left            =   525
         TabIndex        =   78
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum Dither:"
         Height          =   195
         Left            =   195
         TabIndex        =   77
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label lblMaxDitherTerm 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   1980
         TabIndex        =   76
         Top             =   900
         Width           =   420
      End
   End
   Begin VB.CheckBox chkUnguidedDither 
      Caption         =   "Unguided Dither"
      Height          =   252
      Left            =   3240
      TabIndex        =   71
      Top             =   2940
      Width           =   1752
   End
   Begin VB.CheckBox chkAutoguiderEnabled 
      Caption         =   "Autoguider Enabled"
      Height          =   252
      Left            =   3240
      TabIndex        =   70
      Top             =   120
      Value           =   1  'Checked
      Width           =   1752
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
      Left            =   7020
      MaskColor       =   &H00D8E9EC&
      Picture         =   "frmCameraAction.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   275
   End
   Begin VB.CheckBox chkCenterAO 
      Caption         =   "Center AO When Complete"
      Height          =   255
      Left            =   60
      TabIndex        =   38
      Top             =   4740
      Width           =   2355
   End
   Begin VB.TextBox txtSaveTo 
      Height          =   315
      Left            =   1800
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "C:\"
      Top             =   5340
      Width           =   5175
   End
   Begin VB.CheckBox chkRotate 
      Caption         =   "Rotate 180 degrees after Meridian Flip"
      Height          =   312
      Left            =   3300
      TabIndex        =   16
      Top             =   5040
      Width           =   3135
   End
   Begin VB.CheckBox chkImageLinkAfterFlip 
      Caption         =   "Plate Solve and Sync after Meridian Flip"
      Height          =   312
      Left            =   60
      TabIndex        =   15
      Top             =   6300
      Width           =   3135
   End
   Begin VB.Frame fraAutoguider 
      Caption         =   "Autoguider Information"
      Height          =   2475
      Left            =   3240
      TabIndex        =   26
      Top             =   420
      Width           =   2712
      Begin VB.TextBox txtGuiderDelay 
         Height          =   285
         Left            =   1404
         TabIndex        =   90
         Text            =   "0"
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox txtMaxGuideCycles 
         Height          =   285
         Left            =   2070
         TabIndex        =   87
         Text            =   "20"
         Top             =   1800
         Width           =   435
      End
      Begin VB.CommandButton cmdGetGuideStarPos 
         Caption         =   "Get Position"
         Height          =   555
         Left            =   1920
         TabIndex        =   39
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtGuideMinError 
         Height          =   285
         Left            =   1410
         TabIndex        =   13
         Text            =   "0.3"
         Top             =   1500
         Width           =   372
      End
      Begin VB.TextBox txtGuideStarYPos 
         Height          =   285
         Left            =   1404
         TabIndex        =   12
         Text            =   "50"
         Top             =   1170
         Width           =   500
      End
      Begin VB.TextBox txtGuideStarXPos 
         Height          =   285
         Left            =   1404
         TabIndex        =   11
         Text            =   "50"
         Top             =   870
         Width           =   500
      End
      Begin VB.TextBox txtGuiderExposureTime 
         Height          =   285
         Left            =   1404
         TabIndex        =   9
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbGuiderBin 
         Height          =   315
         ItemData        =   "frmCameraAction.frx":05A2
         Left            =   1404
         List            =   "frmCameraAction.frx":05B2
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   540
         Width           =   912
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   1980
         TabIndex        =   92
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guider Init Delay:"
         Height          =   195
         Left            =   150
         TabIndex        =   91
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Max Guide Cycles to Wait:"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   1830
         Width           =   1875
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   1830
         TabIndex        =   33
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Max Error to Start:"
         Height          =   195
         Left            =   75
         TabIndex        =   32
         Top             =   1530
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guide Star Y pos:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guide Star X pos:"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   900
         Width           =   1230
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exposure Time:"
         Height          =   195
         Left            =   210
         TabIndex        =   29
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   1980
         TabIndex        =   28
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bin:"
         Height          =   195
         Left            =   1050
         TabIndex        =   27
         Top             =   600
         Width           =   315
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6060
      TabIndex        =   18
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6060
      TabIndex        =   17
      Top             =   180
      Width           =   1215
   End
   Begin VB.CheckBox chkUseGlobalImageSaveLocation 
      Caption         =   "Use Global Image Save Location"
      Height          =   312
      Left            =   60
      TabIndex        =   14
      Top             =   5040
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog MSComm 
      Left            =   6420
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraPlateSolve 
      Enabled         =   0   'False
      Height          =   3435
      Left            =   60
      TabIndex        =   48
      Top             =   6600
      Width           =   7275
      Begin VB.CheckBox chkAbort 
         Alignment       =   1  'Right Justify
         Caption         =   "Abort List/Sub-List if Solve Fails"
         Height          =   195
         Left            =   3360
         TabIndex        =   69
         Top             =   1680
         Width           =   3075
      End
      Begin VB.Frame Frame4 
         Caption         =   "Retry"
         Height          =   1215
         Left            =   60
         TabIndex        =   63
         Top             =   2160
         Width           =   6015
         Begin VB.CheckBox chkSkip 
            Caption         =   "Skip to Next Target if Second Solve Succeeds"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   840
            Width           =   5175
         End
         Begin VB.TextBox txtSlewOnFailAmount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2760
            TabIndex        =   66
            Text            =   "5"
            Top             =   510
            Width           =   495
         End
         Begin VB.CheckBox chkSlewOnFail 
            Caption         =   "Slew Mount if First Solve Fails"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   540
            Width           =   2535
         End
         Begin VB.CheckBox chkRetry 
            Caption         =   "Retry Plate Solve on Failure"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   2355
         End
         Begin VB.Label Label30 
            Caption         =   "arcminutes toward the zenith."
            Height          =   255
            Left            =   3300
            TabIndex        =   68
            Top             =   540
            Width           =   2475
         End
      End
      Begin VB.Frame fraLink 
         Caption         =   "Plate Solve Exposure Information"
         Height          =   2055
         Left            =   60
         TabIndex        =   52
         Top             =   120
         Width           =   3132
         Begin VB.OptionButton optFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Half Frame"
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   58
            Top             =   1500
            Value           =   -1  'True
            Width           =   1152
         End
         Begin VB.OptionButton optFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Quarter Frame"
            Height          =   252
            Index           =   2
            Left            =   600
            TabIndex        =   57
            Top             =   1740
            Width           =   1392
         End
         Begin VB.OptionButton optFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Frame"
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   56
            Top             =   1260
            Width           =   1152
         End
         Begin VB.ComboBox cmbLinkBin 
            Height          =   315
            ItemData        =   "frmCameraAction.frx":05CA
            Left            =   1800
            List            =   "frmCameraAction.frx":05D7
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   600
            Width           =   912
         End
         Begin VB.TextBox txtLinkExpTime 
            Height          =   252
            Left            =   1800
            TabIndex        =   54
            Text            =   "5"
            Top             =   300
            Width           =   372
         End
         Begin VB.ComboBox cmbLinkFilter 
            Height          =   315
            ItemData        =   "frmCameraAction.frx":05EA
            Left            =   1800
            List            =   "frmCameraAction.frx":05FD
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   900
            Width           =   912
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Bin:"
            Height          =   192
            Left            =   1440
            TabIndex        =   62
            Top             =   660
            Width           =   312
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "seconds"
            Height          =   192
            Left            =   2220
            TabIndex        =   61
            Top             =   360
            Width           =   624
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Exposure Time:"
            Height          =   192
            Left            =   552
            TabIndex        =   60
            Top             =   330
            Width           =   1212
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Filter:"
            Height          =   195
            Left            =   1320
            TabIndex        =   59
            Top             =   960
            Width           =   435
         End
      End
      Begin VB.Frame fraSync 
         Caption         =   "Sync Selection"
         Height          =   1212
         Left            =   3240
         TabIndex        =   49
         Top             =   120
         Width           =   1812
         Begin VB.OptionButton optSync 
            Caption         =   "Offset Position"
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   660
            Width           =   1512
         End
         Begin VB.OptionButton optSync 
            Caption         =   "Sync Mount"
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Value           =   -1  'True
            Width           =   1212
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Imager Information"
      Height          =   4635
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   3132
      Begin VB.Frame Frame3 
         Height          =   1155
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Width           =   2895
         Begin VB.TextBox txtCustomHeight 
            Enabled         =   0   'False
            Height          =   252
            Left            =   2160
            TabIndex        =   7
            Text            =   "1024"
            Top             =   600
            Width           =   552
         End
         Begin VB.TextBox txtCustomWidth 
            Enabled         =   0   'False
            Height          =   252
            Left            =   2160
            TabIndex        =   6
            Text            =   "1024"
            Top             =   240
            Width           =   552
         End
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Custom Size"
            Height          =   252
            Index           =   3
            Left            =   180
            TabIndex        =   44
            Top             =   840
            Width           =   1275
         End
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Half Frame"
            Height          =   252
            Index           =   1
            Left            =   300
            TabIndex        =   43
            Top             =   360
            Width           =   1152
         End
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Quarter Frame"
            Height          =   252
            Index           =   2
            Left            =   60
            TabIndex        =   42
            Top             =   600
            Width           =   1392
         End
         Begin VB.OptionButton optImageFrameSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Full Frame"
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   41
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Height:"
            Height          =   195
            Left            =   1620
            TabIndex        =   46
            Top             =   630
            Width           =   510
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Width:"
            Height          =   195
            Left            =   1620
            TabIndex        =   45
            Top             =   270
            Width           =   465
         End
      End
      Begin VB.CheckBox chkCalibrateImages 
         Alignment       =   1  'Right Justify
         Caption         =   "Calibrate Images"
         Height          =   312
         Left            =   480
         TabIndex        =   8
         Top             =   3480
         Width           =   1515
      End
      Begin VB.TextBox txtDelay 
         Height          =   252
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   780
         Width           =   915
      End
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         ItemData        =   "frmCameraAction.frx":061E
         Left            =   1800
         List            =   "frmCameraAction.frx":0631
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1380
         Width           =   912
      End
      Begin VB.TextBox txtNumberOfExposures 
         Height          =   252
         Left            =   1800
         TabIndex        =   1
         Text            =   "10"
         Top             =   480
         Width           =   552
      End
      Begin VB.TextBox txtExposureTime 
         Height          =   252
         Left            =   1800
         TabIndex        =   0
         Text            =   "300"
         Top             =   180
         Width           =   915
      End
      Begin VB.ComboBox cmbBin 
         Height          =   315
         ItemData        =   "frmCameraAction.frx":0652
         Left            =   1800
         List            =   "frmCameraAction.frx":0662
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   912
      End
      Begin VB.ComboBox cmbImageType 
         Height          =   315
         ItemData        =   "frmCameraAction.frx":067A
         Left            =   1800
         List            =   "frmCameraAction.frx":068A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   912
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   2760
         TabIndex        =   35
         Top             =   780
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Delay Before Exposure:"
         Height          =   195
         Left            =   90
         TabIndex        =   34
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Filter:"
         Height          =   195
         Left            =   1335
         TabIndex        =   25
         Top             =   1410
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Exposure Time:"
         Height          =   195
         Left            =   555
         TabIndex        =   24
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   2760
         TabIndex        =   23
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Bin:"
         Height          =   195
         Left            =   1455
         TabIndex        =   22
         Top             =   1110
         Width           =   315
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Image Type:"
         Height          =   195
         Left            =   855
         TabIndex        =   21
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Exposures:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   510
         Width           =   1590
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Filename Prefix"
      Height          =   195
      Left            =   540
      TabIndex        =   85
      Top             =   6000
      Width           =   1110
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Image Save Location:"
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   5400
      Width           =   1560
   End
End
Attribute VB_Name = "frmCameraAction"
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

Private Const RegistryName = "CameraAction"

Private Sub CancelButton_Click()
    Me.Tag = "False"
    Me.Hide
    Call GetSettings
End Sub

Private Sub GetSettings()
    Me.cmbBin.ListIndex = CInt(GetMySetting(RegistryName, "ImagerBin", "0"))
    If CInt(GetMySetting(RegistryName, "Filter", "0")) < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = CInt(GetMySetting(RegistryName, "Filter", "0"))
    End If
    Me.cmbImageType.ListIndex = CInt(GetMySetting(RegistryName, "ImageType", "0"))
    Me.cmbGuiderBin.ListIndex = CInt(GetMySetting(RegistryName, "GuiderBin", "0"))
    Me.chkAutoguiderEnabled.Value = CInt(GetMySetting(RegistryName, "AutoguiderEnabled", "1"))
    Me.chkCalibrateImages.Value = CInt(GetMySetting(RegistryName, "CalibrateImages", "0"))
    Me.chkAutosave.Value = CInt(GetMySetting(RegistryName, "Autosave", "1"))
    Me.txtDitherStep.Text = GetMySetting(RegistryName, "DitherStep", Format(1, "0.0"))
    Me.txtDitherAmount.Text = GetMySetting(RegistryName, "DitherAmount", Format(4, "0.0"))
    Me.txtDitherFrequency.Text = GetMySetting(RegistryName, "DitherFrequency", "1")
    Me.txtExposureTime.Text = GetMySetting(RegistryName, "ExposureTime", Format(300, "0.000"))
    Me.txtDelay.Text = GetMySetting(RegistryName, "DelayTime", Format(0, "0.0"))
    Me.txtFileNamePrefix.Text = GetMySetting(RegistryName, "FileNamePrefix", "")
    Me.txtGuiderExposureTime.Text = GetMySetting(RegistryName, "GuiderExposureTime", Format(1, "0.00"))
    Me.txtGuideStarXPos.Text = GetMySetting(RegistryName, "GuideStarXPos", "50")
    Me.txtGuideStarYPos.Text = GetMySetting(RegistryName, "GuideStarYPos", "50")
    Me.txtGuideMinError.Text = GetMySetting(RegistryName, "GuideMinError", Format(0.3, "0.0"))
    Me.txtMaxGuideCycles.Text = GetMySetting(RegistryName, "GuideMaxCycles", "20")
    Me.txtGuiderDelay.Text = GetMySetting(RegistryName, "GuiderDelay", "0")
    Me.txtNumberOfExposures.Text = GetMySetting(RegistryName, "NumberOfExposures", "10")
    
    Me.chkImageLinkAfterFlip.Value = CInt(GetMySetting(RegistryName, "ImageLinkAfterFlip", "0"))
    Me.chkDoubleImageLink.Value = CInt(GetMySetting(RegistryName, "DoubleImageLink", "0"))
    Me.cmbLinkBin.ListIndex = CInt(GetMySetting(RegistryName, "ImageLinkBin", "0"))
    If CInt(GetMySetting(RegistryName, "ImageLinkFilter", GetMySetting(RegistryName, "Filter", "0"))) < Me.cmbLinkFilter.ListCount Then
        Me.cmbLinkFilter.ListIndex = CInt(GetMySetting(RegistryName, "ImageLinkFilter", GetMySetting(RegistryName, "Filter", "0")))
    End If
    Me.txtLinkExpTime.Text = GetMySetting(RegistryName, "ImageLinkExposureTime", "5")
    Me.optFrameSize(CInt(GetMySetting(RegistryName, "FrameSize", "0"))).Value = True
    Me.optSync(CInt(GetMySetting(RegistryName, "SyncMode", "0"))).Value = True
    
    Me.chkRotate.Value = CInt(GetMySetting(RegistryName, "RotateAfterFlip", "0"))
    
    Me.optImageFrameSize(CInt(GetMySetting(RegistryName, "ImageFrameSize", "0"))).Value = True
    Me.txtCustomHeight.Text = GetMySetting(RegistryName, "CustomFrameHeight", "1024")
    Me.txtCustomWidth.Text = GetMySetting(RegistryName, "CustomFrameWidth", "1024")
        
    Me.chkUseGlobalImageSaveLocation.Value = CInt(GetMySetting(RegistryName, "UseGlobalImageSaveLocation", "1"))
    Me.txtSaveTo.Text = GetMySetting(RegistryName, "SaveToPath", frmOptions.txtSaveTo.Text)
    
    Call chkUseGlobalImageSaveLocation_Click
    
    Me.chkCenterAO.Value = CInt(GetMySetting(RegistryName, "CenterAO", "0"))
    Me.chkSyncToCurrent.Value = CInt(GetMySetting(RegistryName, "SyncToCurrent", "0"))

    Me.chkUnguidedDither.Value = CInt(GetMySetting(RegistryName, "UnguidedDither", "0"))
    Me.chkMaintainDitherOnFilterChange.Value = CInt(GetMySetting(RegistryName, "MaintainDitherOnFilterChange", "0"))
    Me.chkXAxisDither.Value = CInt(GetMySetting(RegistryName, "XAxisDither", "1"))
    Me.chkYAxisDither.Value = CInt(GetMySetting(RegistryName, "YAxisDither", "1"))
    
    If Me.chkAutoguiderEnabled.Value = vbChecked Then
        Call chkAutoguiderEnabled_Click
    Else
        Call chkUnguidedDither_Click
    End If
    
    Me.chkAbort.Value = GetMySetting(RegistryName, "AbortIfFail", "0")
    Me.chkRetry.Value = GetMySetting(RegistryName, "Retry", "0")
    Me.chkSlewOnFail.Value = GetMySetting(RegistryName, "SlewOnFail", "0")
    Me.txtSlewOnFailAmount.Text = GetMySetting(RegistryName, "SlewOnFailAmount", "1")
    Me.chkSkip.Value = GetMySetting(RegistryName, "SkipOnFail", "0")
End Sub

Private Sub chkAutoguiderEnabled_Click()
    If Me.chkAutoguiderEnabled.Value = vbChecked Then
        Me.fraAutoguider.Enabled = True
        
        If frmOptions.chkContinuousAutoguide.Value = vbChecked And frmOptions.chkEnable.Value = vbChecked Then
            Me.fraDither.Enabled = False
        Else
            Me.fraDither.Enabled = True
        End If
        
        Me.chkUnguidedDither.Enabled = False
        
        Me.lblDitherStepTerm.Caption = "pixels"
        Me.lblMaxDitherTerm.Caption = "pixels"
    Else
        Me.fraAutoguider.Enabled = False
        Me.fraDither.Enabled = False
        Me.chkUnguidedDither.Enabled = True
    End If
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

Private Sub chkImageLinkAfterFlip_Click()
    If Me.chkImageLinkAfterFlip.Value = vbChecked Then
        Me.fraPlateSolve.Enabled = True
        Me.chkDoubleImageLink.Enabled = True
    Else
        Me.fraPlateSolve.Enabled = False
        Me.chkDoubleImageLink.Enabled = False
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

Private Sub chkUnguidedDither_Click()
    If Me.chkUnguidedDither.Value = vbChecked Then
        Me.fraDither.Enabled = True
        Me.lblDitherStepTerm.Caption = "arcsec"
        Me.lblMaxDitherTerm.Caption = "steps"
        Me.chkAutoguiderEnabled.Enabled = False
    Else
        Me.fraDither.Enabled = False
        Me.chkAutoguiderEnabled.Enabled = True
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

Private Sub cmbImageType_Click()
    If Me.cmbImageType.ListIndex = 1 Then
        'Bias frame
        Me.txtExposureTime.Enabled = False
    Else
        Me.txtExposureTime.Enabled = True
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

Private Sub cmdGetGuideStarPos_Click()
    Call Camera.CameraSetup
    
    Me.txtGuideStarXPos.Text = CLng(Camera.objCameraControl.GuideStarX)
    Me.txtGuideStarYPos.Text = CLng(Camera.objCameraControl.GuideStarY)
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
    Me.cmbLinkFilter.Clear
    
    For Counter = 0 To frmOptions.lstFilters.ListCount - 1
        Call Me.cmbFilter.AddItem(frmOptions.lstFilters.List(Counter))
        Call Me.cmbLinkFilter.AddItem(frmOptions.lstFilters.List(Counter))
    Next Counter
    
    Me.cmbGuiderBin.Enabled = True
    
    If frmOptions.lstRotator.ListIndex = RotatorControl.None Then
        Me.chkRotate.Enabled = False
    Else
        Me.chkRotate.Enabled = True
    End If
    
    If frmOptions.chkEnable.Value = vbChecked Then
        Me.txtGuiderExposureTime.Enabled = False
        Me.txtGuideStarXPos.Enabled = False
        Me.txtGuideStarYPos.Enabled = False
        Me.cmdGetGuideStarPos.Visible = False
    Else
        Me.txtGuiderExposureTime.Enabled = True
        Me.txtGuideStarXPos.Enabled = True
        Me.txtGuideStarYPos.Enabled = True
        Me.cmdGetGuideStarPos.Visible = True
    End If
    
    If frmOptions.optMountType(0).Value Then
        Me.chkImageLinkAfterFlip.Enabled = True
    Else
        Me.chkImageLinkAfterFlip.Value = 0
        Me.chkImageLinkAfterFlip.Enabled = False
    End If
    
    Call GetSettings

    If frmOptions.chkContinuousAutoguide.Value = vbChecked And frmOptions.chkEnable.Value = vbChecked Then
        Me.fraDither.Enabled = False
        Me.txtDitherFrequency.Text = "0"
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

Public Sub PutFormDataIntoClass(clsAction As ImagerAction)
    clsAction.ImagerExpTime = CDbl(Me.txtExposureTime.Text)
    clsAction.ImagerNumExp = CLng(Me.txtNumberOfExposures.Text)
    clsAction.ImagerDelayTime = CDbl(Me.txtDelay.Text)
    clsAction.ImagerBin = Me.cmbBin.ListIndex
    clsAction.ImagerFilter = Me.cmbFilter.ListIndex
    clsAction.ImagerType = Me.cmbImageType.ListIndex + 1
    clsAction.CalibrateImages = Me.chkCalibrateImages.Value
    clsAction.AutosaveExposure = Me.chkAutosave.Value
    clsAction.FileNamePrefix = Me.txtFileNamePrefix.Text
    clsAction.AutoguiderExpTime = CDbl(Me.txtGuiderExposureTime.Text)
    clsAction.AutoguiderBin = Me.cmbGuiderBin.ListIndex
    clsAction.AutoguiderDitherFreq = Me.txtDitherFrequency.Text
    clsAction.AutoguiderDitherAmount = Me.txtDitherAmount.Text
    clsAction.AutoguiderDitherStep = Me.txtDitherStep.Text
    clsAction.AutoguiderXPos = Me.txtGuideStarXPos.Text
    clsAction.AutoguiderYPos = Me.txtGuideStarYPos.Text
    clsAction.AutoguiderEnabled = Me.chkAutoguiderEnabled.Value
    clsAction.AutoguiderMinError = CDbl(Me.txtGuideMinError.Text)
    clsAction.AutoguiderMaxGuideCycles = Me.txtMaxGuideCycles.Text
    clsAction.AutoguiderDelay = Me.txtGuiderDelay.Text
    clsAction.ImageLinkAfterMeridianFlip = Me.chkImageLinkAfterFlip.Value
    clsAction.DoubleImageLink = Me.chkDoubleImageLink.Value
    clsAction.clsImageLinkAction.Bin = Me.cmbLinkBin.ListIndex
    clsAction.clsImageLinkAction.ExpTime = CDbl(Me.txtLinkExpTime.Text)
    clsAction.clsImageLinkAction.DelayTime = CDbl(Me.txtDelay.Text)
    clsAction.clsImageLinkAction.Filter = Me.cmbLinkFilter.ListIndex
    If Me.optFrameSize(0).Value Then
        clsAction.clsImageLinkAction.FrameSize = FullFrame
    ElseIf Me.optFrameSize(1).Value Then
        clsAction.clsImageLinkAction.FrameSize = HalfFrame
    ElseIf Me.optFrameSize(2).Value Then
        clsAction.clsImageLinkAction.FrameSize = QuarterFrame
    End If
    If Me.optSync(0).Value Then
        clsAction.clsImageLinkAction.SyncMode = MountSync
    ElseIf Me.optSync(1).Value Then
        clsAction.clsImageLinkAction.SyncMode = Offset
    End If
    clsAction.clsImageLinkAction.SlewToOriginalLocation = vbChecked
    
    clsAction.clsImageLinkAction.RetryPlateSolveOnFailure = Me.chkRetry.Value
    clsAction.clsImageLinkAction.SlewMountForRetry = Me.chkSlewOnFail.Value
    clsAction.clsImageLinkAction.ArcminutesToSlew = CLng(Me.txtSlewOnFailAmount.Text)
    clsAction.clsImageLinkAction.SkipIfRetrySucceeds = Me.chkSkip.Value
    clsAction.clsImageLinkAction.AbortListOnFailure = Me.chkAbort.Value
    clsAction.clsImageLinkAction.AutosaveExposure = vbUnchecked
    
    clsAction.RotateAfterFlip = Me.chkRotate.Value
    
    If Me.optImageFrameSize(0).Value Then
        clsAction.FrameSize = FullFrame
    ElseIf Me.optImageFrameSize(1).Value Then
        clsAction.FrameSize = HalfFrame
    ElseIf Me.optImageFrameSize(2).Value Then
        clsAction.FrameSize = QuarterFrame
    ElseIf Me.optImageFrameSize(3).Value Then
        clsAction.FrameSize = CustomFrame
    End If

    clsAction.CustomFrameHeight = Me.txtCustomHeight.Text
    clsAction.CustomFrameWidth = Me.txtCustomWidth.Text

    clsAction.UseGlobalSaveToLocation = Me.chkUseGlobalImageSaveLocation.Value
    clsAction.FileSavePath = Me.txtSaveTo.Text
    
    clsAction.CenterAO = Me.chkCenterAO.Value
    clsAction.SyncToCurrentAtEnd = Me.chkSyncToCurrent.Value
    
    clsAction.UnguidedDither = Me.chkUnguidedDither.Value
    clsAction.MaintainDitherOnFilterChange = Me.chkMaintainDitherOnFilterChange.Value
    
    clsAction.XAxisDither = Me.chkXAxisDither.Value
    clsAction.YAxisDither = Me.chkYAxisDither.Value
End Sub

Public Sub GetFormDataFromClass(clsAction As ImagerAction)
    Me.txtExposureTime.Text = Format(clsAction.ImagerExpTime, "0.000")
    Me.txtNumberOfExposures.Text = clsAction.ImagerNumExp
    Me.txtDelay.Text = Format(clsAction.ImagerDelayTime, "0.0")
    Me.cmbBin.ListIndex = clsAction.ImagerBin
    If clsAction.ImagerFilter < Me.cmbFilter.ListCount Then
        Me.cmbFilter.ListIndex = clsAction.ImagerFilter
    End If
    Me.cmbImageType.ListIndex = clsAction.ImagerType - 1
    Me.chkCalibrateImages.Value = clsAction.CalibrateImages
    Me.chkAutosave.Value = clsAction.AutosaveExposure
    Me.txtFileNamePrefix.Text = clsAction.FileNamePrefix
    Me.txtGuiderExposureTime.Text = Format(clsAction.AutoguiderExpTime, "0.00")
    Me.cmbGuiderBin.ListIndex = clsAction.AutoguiderBin
    Me.txtDitherFrequency.Text = clsAction.AutoguiderDitherFreq
    Me.txtDitherAmount.Text = Format(clsAction.AutoguiderDitherAmount, "0.0")
    Me.txtDitherStep.Text = Format(clsAction.AutoguiderDitherStep, "0.0")
    Me.txtGuideStarXPos.Text = clsAction.AutoguiderXPos
    Me.txtGuideStarYPos.Text = clsAction.AutoguiderYPos
    Me.chkAutoguiderEnabled.Value = clsAction.AutoguiderEnabled
    Me.txtGuideMinError.Text = clsAction.AutoguiderMinError
    Me.txtMaxGuideCycles.Text = clsAction.AutoguiderMaxGuideCycles
    Me.txtGuiderDelay.Text = clsAction.AutoguiderDelay
    Me.chkImageLinkAfterFlip.Value = -clsAction.ImageLinkAfterMeridianFlip
    Me.chkDoubleImageLink.Value = -clsAction.DoubleImageLink
    Me.cmbLinkBin.ListIndex = clsAction.clsImageLinkAction.Bin
    If clsAction.clsImageLinkAction.Filter < Me.cmbLinkFilter.ListCount Then
        Me.cmbLinkFilter.ListIndex = clsAction.clsImageLinkAction.Filter
    End If
    Me.txtLinkExpTime.Text = clsAction.clsImageLinkAction.ExpTime
    If clsAction.clsImageLinkAction.FrameSize = FullFrame Then
        Me.optFrameSize(0).Value = True
    ElseIf clsAction.clsImageLinkAction.FrameSize = HalfFrame Then
         Me.optFrameSize(1).Value = True
    ElseIf clsAction.clsImageLinkAction.FrameSize = QuarterFrame Then
        Me.optFrameSize(2).Value = True
    End If
    If clsAction.clsImageLinkAction.SyncMode = MountSync Then
        Me.optSync(0).Value = True
    ElseIf clsAction.clsImageLinkAction.SyncMode = Offset Then
        Me.optSync(1).Value = True
    End If
    
    Me.chkRetry.Value = -clsAction.clsImageLinkAction.RetryPlateSolveOnFailure
    Me.chkSlewOnFail.Value = -clsAction.clsImageLinkAction.SlewMountForRetry
    Me.txtSlewOnFailAmount.Text = clsAction.clsImageLinkAction.ArcminutesToSlew
    Me.chkSkip.Value = -clsAction.clsImageLinkAction.SkipIfRetrySucceeds
    Me.chkAbort.Value = -clsAction.clsImageLinkAction.AbortListOnFailure
    
    Me.chkRotate.Value = clsAction.RotateAfterFlip
    
    If clsAction.FrameSize = FullFrame Then
        Me.optImageFrameSize(0).Value = True
    ElseIf clsAction.FrameSize = HalfFrame Then
         Me.optImageFrameSize(1).Value = True
    ElseIf clsAction.FrameSize = QuarterFrame Then
        Me.optImageFrameSize(2).Value = True
    ElseIf clsAction.FrameSize = CustomFrame Then
        Me.optImageFrameSize(3).Value = True
    End If
    
    Me.txtCustomHeight.Text = clsAction.CustomFrameHeight
    Me.txtCustomWidth.Text = clsAction.CustomFrameWidth
    
    Me.chkUseGlobalImageSaveLocation.Value = clsAction.UseGlobalSaveToLocation
    If Me.chkUseGlobalImageSaveLocation.Value = vbChecked Then
        Me.txtSaveTo.Text = frmOptions.txtSaveTo.Text
    Else
        Me.txtSaveTo.Text = clsAction.FileSavePath
    End If
    
    Call chkUseGlobalImageSaveLocation_Click
    
    Me.chkCenterAO.Value = clsAction.CenterAO
    Me.chkSyncToCurrent.Value = clsAction.SyncToCurrentAtEnd
    Me.chkUnguidedDither.Value = clsAction.UnguidedDither
    
    If Me.chkAutoguiderEnabled.Value = vbChecked Then
        Call chkAutoguiderEnabled_Click
    Else
        Call chkUnguidedDither_Click
    End If
    
    Me.chkMaintainDitherOnFilterChange.Value = clsAction.MaintainDitherOnFilterChange
    
    Me.chkXAxisDither = -clsAction.XAxisDither
    Me.chkYAxisDither = -clsAction.YAxisDither
    
    'save the data into the registry just in case the user cancels
    Call SaveMySetting(RegistryName, "ImagerBin", Me.cmbBin.ListIndex)
    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    Call SaveMySetting(RegistryName, "ImageType", Me.cmbImageType.ListIndex)
    Call SaveMySetting(RegistryName, "GuiderBin", Me.cmbGuiderBin.ListIndex)
    Call SaveMySetting(RegistryName, "AutoguiderEnabled", Me.chkAutoguiderEnabled.Value)
    Call SaveMySetting(RegistryName, "CalibrateImages", Me.chkCalibrateImages.Value)
    Call SaveMySetting(RegistryName, "Autosave", Me.chkAutosave.Value)
    Call SaveMySetting(RegistryName, "DitherAmount", Me.txtDitherAmount.Text)
    Call SaveMySetting(RegistryName, "DitherFrequency", Me.txtDitherFrequency.Text)
    Call SaveMySetting(RegistryName, "DitherStep", Me.txtDitherStep.Text)
    Call SaveMySetting(RegistryName, "ExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "DelayTime", Me.txtDelay.Text)
    Call SaveMySetting(RegistryName, "FileNamePrefix", Me.txtFileNamePrefix.Text)
    Call SaveMySetting(RegistryName, "GuiderExposureTime", Me.txtGuiderExposureTime.Text)
    Call SaveMySetting(RegistryName, "GuideStarXPos", Me.txtGuideStarXPos.Text)
    Call SaveMySetting(RegistryName, "GuideStarYPos", Me.txtGuideStarYPos.Text)
    Call SaveMySetting(RegistryName, "GuideMinError", Me.txtGuideMinError.Text)
    Call SaveMySetting(RegistryName, "GuideMaxCycles", Me.txtMaxGuideCycles.Text)
    Call SaveMySetting(RegistryName, "GuiderDelay", Me.txtGuiderDelay.Text)
    Call SaveMySetting(RegistryName, "NumberOfExposures", Me.txtNumberOfExposures.Text)
    
    Call SaveMySetting(RegistryName, "ImageLinkAfterFlip", Me.chkImageLinkAfterFlip.Value)
    Call SaveMySetting(RegistryName, "DoubleImageLink", Me.chkDoubleImageLink.Value)
    Call SaveMySetting(RegistryName, "ImageLinkBin", Me.cmbLinkBin.ListIndex)
    Call SaveMySetting(RegistryName, "ImageLinkFilter", Me.cmbLinkFilter.ListIndex)
    Call SaveMySetting(RegistryName, "ImageLinkExposureTime", Me.txtLinkExpTime.Text)
    If Me.optFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "0")
    ElseIf Me.optFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "1")
    ElseIf Me.optFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "2")
    End If
    If Me.optSync(0).Value = True Then
        Call SaveMySetting(RegistryName, "SyncMode", "0")
    Else
        Call SaveMySetting(RegistryName, "SyncMode", "1")
    End If
    
    Call SaveMySetting(RegistryName, "AbortIfFail", Me.chkAbort.Value)
    Call SaveMySetting(RegistryName, "Retry", Me.chkRetry.Value)
    Call SaveMySetting(RegistryName, "SlewOnFail", Me.chkSlewOnFail.Value)
    Call SaveMySetting(RegistryName, "SlewOnFailAmount", Me.txtSlewOnFailAmount.Text)
    Call SaveMySetting(RegistryName, "SkipOnFail", Me.chkSkip.Value)
    
    Call SaveMySetting(RegistryName, "RotateAfterFlip", Me.chkRotate)

    If Me.optImageFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "0")
    ElseIf Me.optImageFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "1")
    ElseIf Me.optImageFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "2")
    ElseIf Me.optImageFrameSize(3).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "3")
    End If

    Call SaveMySetting(RegistryName, "CustomFrameHeight", Me.txtCustomHeight.Text)
    Call SaveMySetting(RegistryName, "CustomFrameWidth", Me.txtCustomWidth.Text)

    Call SaveMySetting(RegistryName, "UseGlobalImageSaveLocation", Me.chkUseGlobalImageSaveLocation.Value)
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)
   
    Call SaveMySetting(RegistryName, "CenterAO", Me.chkCenterAO.Value)
    Call SaveMySetting(RegistryName, "SyncToCurrent", Me.chkSyncToCurrent.Value)

    Call SaveMySetting(RegistryName, "UnguidedDither", Me.chkUnguidedDither.Value)
    
    Call SaveMySetting(RegistryName, "MaintainDitherOnFilterChange", Me.chkMaintainDitherOnFilterChange.Value)
    Call SaveMySetting(RegistryName, "XAxisDither", Me.chkXAxisDither.Value)
    Call SaveMySetting(RegistryName, "YAxisDither", Me.chkYAxisDither.Value)

End Sub

Private Sub OKButton_Click()
    Dim Cancel As Boolean
    Cancel = False
    
    If Not Cancel Then Call txtDitherStep_Validate(Cancel)
    If Not Cancel Then Call txtDitherAmount_Validate(Cancel)
    If Not Cancel Then Call txtDitherFrequency_Validate(Cancel)
    If Not Cancel Then Call txtExposureTime_Validate(Cancel)
    If Not Cancel Then Call txtGuideMinError_Validate(Cancel)
    If Not Cancel Then Call txtMaxGuideCycles_Validate(Cancel)
    If Not Cancel Then Call txtGuiderDelay_Validate(Cancel)
    If Not Cancel Then Call txtGuiderExposureTime_Validate(Cancel)
    If Not Cancel Then Call txtGuideStarXPos_Validate(Cancel)
    If Not Cancel Then Call txtGuideStarYPos_Validate(Cancel)
    If Not Cancel Then Call txtLinkExpTime_Validate(Cancel)
    If Not Cancel Then Call txtNumberOfExposures_Validate(Cancel)
    If Not Cancel Then Call txtDelay_Validate(Cancel)
    If Not Cancel Then Call txtCustomWidth_Validate(Cancel)
    If Not Cancel Then Call txtCustomHeight_Validate(Cancel)
    If Not Cancel Then Call txtSlewOnFailAmount_Validate(Cancel)
    
    If Cancel Then
        Beep
        Exit Sub
    End If
    
    Call SaveMySetting(RegistryName, "ImagerBin", Me.cmbBin.ListIndex)
    Call SaveMySetting(RegistryName, "Filter", Me.cmbFilter.ListIndex)
    Call SaveMySetting(RegistryName, "ImageType", Me.cmbImageType.ListIndex)
    Call SaveMySetting(RegistryName, "GuiderBin", Me.cmbGuiderBin.ListIndex)
    Call SaveMySetting(RegistryName, "AutoguiderEnabled", Me.chkAutoguiderEnabled.Value)
    Call SaveMySetting(RegistryName, "CalibrateImages", Me.chkCalibrateImages.Value)
    Call SaveMySetting(RegistryName, "Autosave", Me.chkAutosave.Value)
    Call SaveMySetting(RegistryName, "DitherAmount", Me.txtDitherAmount.Text)
    Call SaveMySetting(RegistryName, "DitherStep", Me.txtDitherStep.Text)
    Call SaveMySetting(RegistryName, "DitherFrequency", Me.txtDitherFrequency.Text)
    Call SaveMySetting(RegistryName, "ExposureTime", Me.txtExposureTime.Text)
    Call SaveMySetting(RegistryName, "DelayTime", Me.txtDelay.Text)
    Call SaveMySetting(RegistryName, "FileNamePrefix", Me.txtFileNamePrefix.Text)
    Call SaveMySetting(RegistryName, "GuiderExposureTime", Me.txtGuiderExposureTime.Text)
    Call SaveMySetting(RegistryName, "GuideStarXPos", Me.txtGuideStarXPos.Text)
    Call SaveMySetting(RegistryName, "GuideStarYPos", Me.txtGuideStarYPos.Text)
    Call SaveMySetting(RegistryName, "GuideMinError", Me.txtGuideMinError.Text)
    Call SaveMySetting(RegistryName, "GuideMaxCycles", Me.txtMaxGuideCycles.Text)
    Call SaveMySetting(RegistryName, "GuiderDelay", Me.txtGuiderDelay.Text)
    Call SaveMySetting(RegistryName, "NumberOfExposures", Me.txtNumberOfExposures.Text)
    
    Call SaveMySetting(RegistryName, "ImageLinkAfterFlip", Me.chkImageLinkAfterFlip.Value)
    Call SaveMySetting(RegistryName, "DoubleImageLink", Me.chkDoubleImageLink.Value)
    Call SaveMySetting(RegistryName, "ImageLinkBin", Me.cmbLinkBin.ListIndex)
    Call SaveMySetting(RegistryName, "ImageLinkFilter", Me.cmbLinkFilter.ListIndex)
    Call SaveMySetting(RegistryName, "ImageLinkExposureTime", Me.txtLinkExpTime.Text)
    If Me.optFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "0")
    ElseIf Me.optFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "1")
    ElseIf Me.optFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "FrameSize", "2")
    End If
    If Me.optSync(0).Value = True Then
        Call SaveMySetting(RegistryName, "SyncMode", "0")
    Else
        Call SaveMySetting(RegistryName, "SyncMode", "1")
    End If
    
    Call SaveMySetting(RegistryName, "AbortIfFail", Me.chkAbort.Value)
    Call SaveMySetting(RegistryName, "Retry", Me.chkRetry.Value)
    Call SaveMySetting(RegistryName, "SlewOnFail", Me.chkSlewOnFail.Value)
    Call SaveMySetting(RegistryName, "SlewOnFailAmount", Me.txtSlewOnFailAmount.Text)
    Call SaveMySetting(RegistryName, "SkipOnFail", Me.chkSkip.Value)
    
    Call SaveMySetting(RegistryName, "RotateAfterFlip", Me.chkRotate)
    
    If Me.optImageFrameSize(0).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "0")
    ElseIf Me.optImageFrameSize(1).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "1")
    ElseIf Me.optImageFrameSize(2).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "2")
    ElseIf Me.optImageFrameSize(3).Value Then
        Call SaveMySetting(RegistryName, "ImageFrameSize", "3")
    End If

    Call SaveMySetting(RegistryName, "CustomFrameHeight", Me.txtCustomHeight.Text)
    Call SaveMySetting(RegistryName, "CustomFrameWidth", Me.txtCustomWidth.Text)
    
    Call SaveMySetting(RegistryName, "UseGlobalImageSaveLocation", Me.chkUseGlobalImageSaveLocation.Value)
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)
    
    Call SaveMySetting(RegistryName, "CenterAO", Me.chkCenterAO.Value)
    Call SaveMySetting(RegistryName, "SyncToCurrent", Me.chkSyncToCurrent.Value)
    
    Call SaveMySetting(RegistryName, "UnguidedDither", Me.chkUnguidedDither.Value)
    
    Call SaveMySetting(RegistryName, "MaintainDitherOnFilterChange", Me.chkMaintainDitherOnFilterChange.Value)
    Call SaveMySetting(RegistryName, "XAxisDither", Me.chkXAxisDither.Value)
    Call SaveMySetting(RegistryName, "YAxisDither", Me.chkYAxisDither.Value)
    
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

Private Sub optImageFrameSize_Click(Index As Integer)
    Me.optImageFrameSize(Index).Value = True
    
    If Index = 0 Then
        Me.optImageFrameSize(1).Value = False
        Me.optImageFrameSize(2).Value = False
        Me.optImageFrameSize(3).Value = False
    ElseIf Index = 1 Then
        Me.optImageFrameSize(0).Value = False
        Me.optImageFrameSize(2).Value = False
        Me.optImageFrameSize(3).Value = False
    ElseIf Index = 2 Then
        Me.optImageFrameSize(0).Value = False
        Me.optImageFrameSize(1).Value = False
        Me.optImageFrameSize(3).Value = False
    ElseIf Index = 3 Then
        Me.optImageFrameSize(0).Value = False
        Me.optImageFrameSize(1).Value = False
        Me.optImageFrameSize(2).Value = False
    End If
    
    If Index = 3 Then
        Me.txtCustomHeight.Enabled = True
        Me.txtCustomWidth.Enabled = True
    Else
        Me.txtCustomHeight.Enabled = False
        Me.txtCustomWidth.Enabled = False
    End If
End Sub

Private Sub optSync_Click(Index As Integer)
    Me.optSync(Index).Value = True
    
    If Index = 0 Then
        Me.optSync(1).Value = False
    Else
        Me.optSync(0).Value = False
    End If
End Sub

Private Sub txtCustomWidth_GotFocus()
    Me.txtCustomWidth.SelStart = 0
    Me.txtCustomWidth.SelLength = Len(Me.txtCustomWidth.Text)
End Sub

Private Sub txtCustomWidth_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtCustomWidth.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtCustomWidth.Text Then
        Beep
        Cancel = True
    Else
        Me.txtCustomWidth.Text = Format(Me.txtCustomWidth.Text, "0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtCustomHeight_GotFocus()
    Me.txtCustomHeight.SelStart = 0
    Me.txtCustomHeight.SelLength = Len(Me.txtCustomHeight.Text)
End Sub

Private Sub txtCustomHeight_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtCustomHeight.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtCustomHeight.Text Then
        Beep
        Cancel = True
    Else
        Me.txtCustomHeight.Text = Format(Me.txtCustomHeight.Text, "0")
    End If
    On Error GoTo 0
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

Private Sub txtDitherAmount_GotFocus()
    Me.txtDitherAmount.SelStart = 0
    Me.txtDitherAmount.SelLength = Len(Me.txtDitherAmount.Text)
End Sub

Private Sub txtDitherAmount_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDitherAmount.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtDitherAmount.Text Then
        Beep
        Cancel = True
    Else
        Me.txtDitherAmount.Text = Format(Me.txtDitherAmount.Text, "0.0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtDitherStep_GotFocus()
    Me.txtDitherStep.SelStart = 0
    Me.txtDitherStep.SelLength = Len(Me.txtDitherStep.Text)
End Sub

Private Sub txtDitherStep_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDitherStep.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtDitherStep.Text Then
        Beep
        Cancel = True
    Else
        Me.txtDitherStep.Text = Format(Me.txtDitherStep.Text, "0.0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtDitherFrequency_GotFocus()
    Me.txtDitherFrequency.SelStart = 0
    Me.txtDitherFrequency.SelLength = Len(Me.txtDitherFrequency.Text)
End Sub

Private Sub txtDitherFrequency_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtDitherFrequency.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtDitherFrequency.Text Then
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
    Else
        Me.txtExposureTime.Text = Format(Me.txtExposureTime.Text, "0.000")
    End If
    On Error GoTo 0
End Sub

Private Sub txtFileNamePrefix_GotFocus()
    Me.txtFileNamePrefix.SelStart = 0
    Me.txtFileNamePrefix.SelLength = Len(Me.txtFileNamePrefix.Text)
End Sub

Private Sub txtGuideMinError_GotFocus()
    Me.txtGuideMinError.SelStart = 0
    Me.txtGuideMinError.SelLength = Len(Me.txtGuideMinError.Text)
End Sub

Private Sub txtGuideMinError_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtGuideMinError.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtGuideMinError.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuiderDelay_GotFocus()
    Me.txtGuiderDelay.SelStart = 0
    Me.txtGuiderDelay.SelLength = Len(Me.txtGuiderDelay.Text)
End Sub

Private Sub txtGuiderDelay_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtGuiderDelay.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtGuiderDelay.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuiderExposureTime_GotFocus()
    Me.txtGuiderExposureTime.SelStart = 0
    Me.txtGuiderExposureTime.SelLength = Len(Me.txtGuiderExposureTime.Text)
End Sub

Private Sub txtGuiderExposureTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtGuiderExposureTime.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtGuiderExposureTime.Text Then
        Beep
        Cancel = True
    Else
        Me.txtGuiderExposureTime.Text = Format(Me.txtGuiderExposureTime.Text, "0.00")
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuideStarXPos_GotFocus()
    Me.txtGuideStarXPos.SelStart = 0
    Me.txtGuideStarXPos.SelLength = Len(Me.txtGuideStarXPos.Text)
End Sub

Private Sub txtGuideStarXPos_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtGuideStarXPos.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtGuideStarXPos.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuideStarYPos_GotFocus()
    Me.txtGuideStarYPos.SelStart = 0
    Me.txtGuideStarYPos.SelLength = Len(Me.txtGuideStarYPos.Text)
End Sub

Private Sub txtGuideStarYPos_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtGuideStarYPos.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtGuideStarYPos.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtLinkExpTime_GotFocus()
    Me.txtLinkExpTime.SelStart = 0
    Me.txtLinkExpTime.SelLength = Len(Me.txtLinkExpTime.Text)
End Sub

Private Sub txtLinkExpTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtLinkExpTime.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtLinkExpTime.Text Then
        Beep
        Cancel = True
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxGuideCycles_GotFocus()
    Me.txtMaxGuideCycles.SelStart = 0
    Me.txtMaxGuideCycles.SelLength = Len(Me.txtMaxGuideCycles.Text)
End Sub

Private Sub txtMaxGuideCycles_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtMaxGuideCycles.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxGuideCycles.Text Then
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

