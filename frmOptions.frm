VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CCD Commander Options & Settings"
   ClientHeight    =   8205
   ClientLeft      =   2565
   ClientTop       =   1800
   ClientWidth     =   8130
   HelpContextID   =   200
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   7380
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      HelpContextID   =   200
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   14314
      _Version        =   393216
      Tabs            =   11
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Control/Device"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(7)=   "Label55"
      Tab(0).Control(8)=   "lblMaximumStarFaded"
      Tab(0).Control(9)=   "lstCloudSensor"
      Tab(0).Control(10)=   "lstDomeControl"
      Tab(0).Control(11)=   "cmdDeleteFilter"
      Tab(0).Control(12)=   "cmdAddFilter"
      Tab(0).Control(13)=   "cmdGetFilters"
      Tab(0).Control(14)=   "lstFilters"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lstRotator"
      Tab(0).Control(16)=   "lstFocuserControl"
      Tab(0).Control(17)=   "lstMountControl"
      Tab(0).Control(18)=   "lstCameraControl"
      Tab(0).Control(19)=   "chkInternalGuider"
      Tab(0).Control(20)=   "lstPlanetarium"
      Tab(0).Control(21)=   "cmdASCOMMountConfigurre"
      Tab(0).Control(22)=   "cmdASCOMDomeConfigurre"
      Tab(0).Control(23)=   "cmdRotatorConfigure"
      Tab(0).Control(24)=   "cmdWeatherMonitorRemoteFileOpen"
      Tab(0).Control(25)=   "txtWeatherMonitorRemoteFile"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkDisableForceFilterChange"
      Tab(0).Control(27)=   "chkIgnoreStarFaded"
      Tab(0).Control(28)=   "chkIgnoreExposureAborted"
      Tab(0).Control(29)=   "chkDisconnectAtEnd"
      Tab(0).Control(30)=   "txtMaximumStarFaded"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Plate Solve"
      TabPicture(1)   =   "frmOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(1)=   "Label18"
      Tab(1).Control(2)=   "Label19"
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(4)=   "fraPinPointSetup"
      Tab(1).Control(5)=   "lstPlateSolve"
      Tab(1).Control(6)=   "txtNorthAngle"
      Tab(1).Control(7)=   "txtPixelScale"
      Tab(1).Control(8)=   "chkDarkSubtractPlateSolveImage"
      Tab(1).Control(9)=   "fraPinPointLESetup"
      Tab(1).Control(10)=   "chkIgnoreNorthAngle"
      Tab(1).Control(11)=   "cmdGet"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Rotator"
      TabPicture(2)   =   "frmOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(3)=   "Label8"
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(5)=   "lblAngles"
      Tab(2).Control(6)=   "chkGuiderRotates"
      Tab(2).Control(7)=   "cmdGetAngleFromTheSky"
      Tab(2).Control(8)=   "chkRotFromSky(1)"
      Tab(2).Control(9)=   "chkRotFromSky(0)"
      Tab(2).Control(10)=   "txtGuiderCalAngle"
      Tab(2).Control(11)=   "txtCOMNum"
      Tab(2).Control(12)=   "txtHomeRotationAngle"
      Tab(2).Control(13)=   "chkReverseRotatorDirection"
      Tab(2).Control(14)=   "chkRotateTheSkyFOVI"
      Tab(2).Control(15)=   "fraRotatorFlip"
      Tab(2).Control(16)=   "chkGuiderMirrorImage"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Weather Monitor"
      TabPicture(3)   =   "frmOptions.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label25"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label24"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label23"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label22"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame5"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "chkParkMountWhenCloudy"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtClearTime"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtQuerySensorPeriod"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Frame6"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Frame7"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "chkAutoDomeClose"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "chkWeatherMonitorRepeatAction"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "fraWeatherMonitorScripts"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "chkEnableWeatherMonitorScripts"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Mount Parameters"
      TabPicture(4)   =   "frmOptions.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label33"
      Tab(4).Control(1)=   "Label34"
      Tab(4).Control(2)=   "Label52"
      Tab(4).Control(3)=   "Label53"
      Tab(4).Control(4)=   "chkEnableScripts"
      Tab(4).Control(5)=   "fraScripts"
      Tab(4).Control(6)=   "fraGEMSetup"
      Tab(4).Control(7)=   "txtDelayAfterSlew"
      Tab(4).Control(8)=   "Frame1"
      Tab(4).Control(9)=   "chkVerifyTeleCoords"
      Tab(4).Control(10)=   "txtMaxPointingError"
      Tab(4).Control(11)=   "chkDisableDecComp"
      Tab(4).Control(12)=   "chkOnlyDelayAfterMeridianFlip"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "Auto Guide Star"
      TabPicture(5)   =   "frmOptions.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label35"
      Tab(5).Control(1)=   "Label36"
      Tab(5).Control(2)=   "Label37"
      Tab(5).Control(3)=   "Label38"
      Tab(5).Control(4)=   "Label39"
      Tab(5).Control(5)=   "Label40"
      Tab(5).Control(6)=   "Label41"
      Tab(5).Control(7)=   "Label42"
      Tab(5).Control(8)=   "Label43"
      Tab(5).Control(9)=   "Label44"
      Tab(5).Control(10)=   "Label45"
      Tab(5).Control(11)=   "Label46"
      Tab(5).Control(12)=   "Label47"
      Tab(5).Control(13)=   "Label48"
      Tab(5).Control(14)=   "Label49"
      Tab(5).Control(15)=   "Label50"
      Tab(5).Control(16)=   "Label68"
      Tab(5).Control(17)=   "Label69"
      Tab(5).Control(18)=   "Label72"
      Tab(5).Control(19)=   "Label73"
      Tab(5).Control(20)=   "Label51"
      Tab(5).Control(21)=   "Label84"
      Tab(5).Control(22)=   "Label85"
      Tab(5).Control(23)=   "Label86"
      Tab(5).Control(24)=   "chkContinuousAutoguide"
      Tab(5).Control(25)=   "txtMaxStarMovement"
      Tab(5).Control(26)=   "txtGuideBoxY"
      Tab(5).Control(27)=   "txtGuideBoxX"
      Tab(5).Control(28)=   "txtMaxBright"
      Tab(5).Control(29)=   "txtMinBright"
      Tab(5).Control(30)=   "txtMaxExp"
      Tab(5).Control(31)=   "txtMinExp"
      Tab(5).Control(32)=   "chkEnable"
      Tab(5).Control(33)=   "txtGuideExposureIncrement"
      Tab(5).Control(34)=   "chkDisableGuideStarRecovery"
      Tab(5).Control(35)=   "txtGuideStarFWHM"
      Tab(5).Control(36)=   "chkIgnore1PixelStars"
      Tab(5).Control(37)=   "chkRestartGuidingWhenLargeError"
      Tab(5).Control(38)=   "txtRestartError"
      Tab(5).Control(39)=   "txtRestartCycles"
      Tab(5).ControlCount=   40
      TabCaption(6)   =   "File Options"
      TabPicture(6)   =   "frmOptions.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label9"
      Tab(6).Control(1)=   "txtSaveTo"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "cmdOpen"
      Tab(6).Control(3)=   "chkMaxImCompression"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Dome/Roof"
      TabPicture(7)   =   "frmOptions.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label54"
      Tab(7).Control(1)=   "chkUncoupleDomeDuringSlews"
      Tab(7).Control(2)=   "fraDDWRetry"
      Tab(7).Control(3)=   "chkHaltOnDomeError"
      Tab(7).Control(4)=   "chkParkMountFirst"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "E-Mail Alerts"
      TabPicture(8)   =   "frmOptions.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame2"
      Tab(8).Control(1)=   "Frame3"
      Tab(8).Control(2)=   "Frame4"
      Tab(8).Control(3)=   "Frame8"
      Tab(8).ControlCount=   4
      TabCaption(9)   =   "Focuser Options"
      TabPicture(9)   =   "frmOptions.frx":0108
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label16"
      Tab(9).Control(1)=   "Label21"
      Tab(9).Control(2)=   "Label74"
      Tab(9).Control(3)=   "Label77"
      Tab(9).Control(4)=   "Label79"
      Tab(9).Control(5)=   "chkEnableFilterOffsets"
      Tab(9).Control(6)=   "lstFiltersFocusOffset"
      Tab(9).Control(6).Enabled=   0   'False
      Tab(9).Control(7)=   "chkMeasureAverageHFD"
      Tab(9).Control(8)=   "chkRetryFocusRunOnFailure"
      Tab(9).Control(9)=   "txtFocusRetryCount"
      Tab(9).Control(10)=   "txtFocusTimeOut"
      Tab(9).ControlCount=   11
      TabCaption(10)  =   "Program Errors"
      TabPicture(10)  =   "frmOptions.frx":0124
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Label80"
      Tab(10).Control(1)=   "chkParkMount"
      Tab(10).Control(2)=   "chkCloseDomeOnError"
      Tab(10).Control(3)=   "chkEnableWatchdog"
      Tab(10).ControlCount=   4
      Begin VB.CheckBox chkOnlyDelayAfterMeridianFlip 
         Caption         =   "Only Delay after Meridian Flip"
         Height          =   315
         Left            =   -72000
         TabIndex        =   273
         Top             =   1360
         Width           =   3375
      End
      Begin VB.TextBox txtMaximumStarFaded 
         Height          =   312
         Left            =   -70770
         TabIndex        =   271
         Text            =   "5"
         ToolTipText     =   "Maximum Star Faded Errors in MaxIm/DL"
         Top             =   2940
         Width           =   390
      End
      Begin VB.TextBox txtRestartCycles 
         Height          =   312
         Left            =   -71100
         TabIndex        =   269
         Text            =   "1"
         Top             =   5160
         Width           =   315
      End
      Begin VB.TextBox txtRestartError 
         Height          =   312
         Left            =   -72300
         TabIndex        =   267
         Text            =   "1"
         Top             =   5160
         Width           =   315
      End
      Begin VB.CheckBox chkRestartGuidingWhenLargeError 
         Caption         =   "Restart autoguiding when the guide error"
         Height          =   255
         Left            =   -73500
         TabIndex        =   265
         Top             =   4920
         Width           =   3315
      End
      Begin VB.CheckBox chkDisconnectAtEnd 
         Caption         =   "Disconnect from all programs/devices at end of action list."
         Height          =   435
         Left            =   -72840
         TabIndex        =   264
         Top             =   6960
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame Frame8 
         Caption         =   "E-Mail Script"
         Height          =   735
         Left            =   -74880
         TabIndex        =   260
         Top             =   6180
         Width           =   6315
         Begin VB.TextBox txtEMailScript 
            BackColor       =   &H8000000F&
            Height          =   288
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   262
            Top             =   300
            Width           =   5595
         End
         Begin VB.CommandButton cmdEMailScript 
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
            Picture         =   "frmOptions.frx":0140
            Style           =   1  'Graphical
            TabIndex        =   261
            Top             =   300
            UseMaskColor    =   -1  'True
            Width           =   275
         End
      End
      Begin VB.CheckBox chkIgnoreExposureAborted 
         Caption         =   "Ignore ""Exposure Aborted"" messages from MaxIm/DL"
         Height          =   195
         Left            =   -72840
         TabIndex        =   259
         Top             =   3300
         Width           =   4095
      End
      Begin VB.CheckBox chkEnableWatchdog 
         Caption         =   "Enable Software Watchdog"
         Height          =   255
         Left            =   -73140
         TabIndex        =   258
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkIgnoreStarFaded 
         Caption         =   "Ignore ""Star Faded"" messages from MaxIm/DL"
         Height          =   195
         Left            =   -72840
         TabIndex        =   257
         Top             =   2700
         Width           =   4095
      End
      Begin VB.CheckBox chkParkMountFirst 
         Caption         =   "Always Park Mount before Closing Dome/Roof"
         Height          =   255
         Left            =   -73200
         TabIndex        =   256
         Top             =   3240
         Width           =   3735
      End
      Begin VB.CheckBox chkEnableWeatherMonitorScripts 
         Caption         =   "Enable Weather Monitor Scripts"
         Height          =   255
         Left            =   120
         TabIndex        =   252
         Top             =   6240
         Width           =   2835
      End
      Begin VB.Frame fraWeatherMonitorScripts 
         Caption         =   "Weather Monitor Scripts"
         Enabled         =   0   'False
         Height          =   1275
         Left            =   120
         TabIndex        =   245
         Top             =   6480
         Width           =   6255
         Begin VB.TextBox txtAfterGoodScript 
            BackColor       =   &H8000000F&
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   254
            Top             =   900
            Width           =   4575
         End
         Begin VB.CommandButton cmdAfterGoodScript 
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
            Left            =   5880
            MaskColor       =   &H00D8E9EC&
            Picture         =   "frmOptions.frx":028C
            Style           =   1  'Graphical
            TabIndex        =   253
            Top             =   900
            UseMaskColor    =   -1  'True
            Width           =   275
         End
         Begin VB.TextBox txtAfterPauseScript 
            BackColor       =   &H8000000F&
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   249
            Top             =   240
            Width           =   4575
         End
         Begin VB.TextBox txtAfterCloseScript 
            BackColor       =   &H8000000F&
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   248
            Top             =   570
            Width           =   4575
         End
         Begin VB.CommandButton cmdAfterCloseScript 
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
            Left            =   5880
            MaskColor       =   &H00D8E9EC&
            Picture         =   "frmOptions.frx":03D8
            Style           =   1  'Graphical
            TabIndex        =   247
            Top             =   580
            UseMaskColor    =   -1  'True
            Width           =   275
         End
         Begin VB.CommandButton cmdAfterPauseScript 
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
            Left            =   5880
            MaskColor       =   &H00D8E9EC&
            Picture         =   "frmOptions.frx":0524
            Style           =   1  'Graphical
            TabIndex        =   246
            Top             =   255
            UseMaskColor    =   -1  'True
            Width           =   275
         End
         Begin VB.Label Label83 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "After Good:"
            Height          =   195
            Left            =   240
            TabIndex        =   255
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label82 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "After Pause:"
            Height          =   195
            Left            =   225
            TabIndex        =   251
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "After Close:"
            Height          =   195
            Left            =   240
            TabIndex        =   250
            Top             =   600
            Width           =   810
         End
      End
      Begin VB.CheckBox chkHaltOnDomeError 
         Caption         =   "Halt on Dome Error"
         Height          =   375
         Left            =   -73200
         TabIndex        =   244
         Top             =   3960
         Value           =   1  'Checked
         Width           =   2715
      End
      Begin VB.CheckBox chkCloseDomeOnError 
         Caption         =   "Close Dome on Error"
         Height          =   255
         Left            =   -73140
         TabIndex        =   242
         Top             =   3420
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkParkMount 
         Caption         =   "Park Mount on Error"
         Height          =   255
         Left            =   -73140
         TabIndex        =   241
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkDisableForceFilterChange 
         Caption         =   "Disable dummy image on filter change"
         Height          =   195
         Left            =   -72840
         TabIndex        =   240
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CheckBox chkGuiderMirrorImage 
         Caption         =   "Guider sees a Mirror Image of the sky"
         Height          =   195
         Left            =   -73200
         TabIndex        =   239
         Top             =   3180
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtFocusTimeOut 
         Height          =   285
         Left            =   -71880
         TabIndex        =   235
         Text            =   "10"
         Top             =   5400
         Width           =   615
      End
      Begin VB.CheckBox chkIgnore1PixelStars 
         Caption         =   "Ignore 1-pixel Stars"
         Height          =   195
         Left            =   -70260
         TabIndex        =   233
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Frame fraDDWRetry 
         Caption         =   "DDW Retry Settings"
         Height          =   1635
         Left            =   -74040
         TabIndex        =   226
         Top             =   4440
         Width           =   3615
         Begin VB.TextBox txtDDWRetryCount 
            Height          =   312
            Left            =   1875
            TabIndex        =   230
            Text            =   "5"
            ToolTipText     =   "Minimum guide exposure time"
            Top             =   960
            Width           =   750
         End
         Begin VB.TextBox txtDDWTimeout 
            Height          =   312
            Left            =   1860
            TabIndex        =   227
            Text            =   "5"
            ToolTipText     =   "Minimum guide exposure time"
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Retry Attempts:"
            Height          =   195
            Left            =   690
            TabIndex        =   231
            Top             =   1020
            Width           =   1080
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Operation Timeout:"
            Height          =   195
            Left            =   405
            TabIndex        =   229
            Top             =   540
            Width           =   1350
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            Caption         =   "minutes"
            Height          =   195
            Left            =   2715
            TabIndex        =   228
            Top             =   555
            Width           =   540
         End
      End
      Begin VB.TextBox txtFocusRetryCount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71880
         TabIndex        =   224
         Text            =   "2"
         Top             =   4920
         Width           =   615
      End
      Begin VB.CheckBox chkRetryFocusRunOnFailure 
         Caption         =   "Retry Focus Run on Failure"
         Height          =   372
         Left            =   -73380
         TabIndex        =   223
         Top             =   4560
         Width           =   3075
      End
      Begin VB.TextBox txtGuideStarFWHM 
         Height          =   312
         Left            =   -71760
         TabIndex        =   220
         Text            =   "4"
         ToolTipText     =   "Guide box ""Y"" pixel size"
         Top             =   3900
         Width           =   750
      End
      Begin VB.CheckBox chkDisableGuideStarRecovery 
         Caption         =   "Disable Automatic Guide Star Recovery"
         Height          =   372
         Left            =   -73500
         TabIndex        =   219
         Top             =   6900
         Width           =   3252
      End
      Begin VB.CheckBox chkWeatherMonitorRepeatAction 
         Caption         =   "Repeat last action after resuming"
         Height          =   255
         Left            =   300
         TabIndex        =   218
         Top             =   4450
         Width           =   2715
      End
      Begin VB.TextBox txtGuideExposureIncrement 
         Height          =   312
         Left            =   -71760
         TabIndex        =   215
         Text            =   "2"
         ToolTipText     =   "Minimum guide exposure time"
         Top             =   1740
         Width           =   750
      End
      Begin VB.TextBox txtWeatherMonitorRemoteFile 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   214
         TabStop         =   0   'False
         Text            =   "C:\"
         Top             =   6420
         Width           =   3915
      End
      Begin VB.CommandButton cmdWeatherMonitorRemoteFileOpen 
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
         Left            =   -68880
         MaskColor       =   &H00D8E9EC&
         Picture         =   "frmOptions.frx":0670
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   6480
         UseMaskColor    =   -1  'True
         Width           =   275
      End
      Begin VB.CheckBox chkMeasureAverageHFD 
         Caption         =   "Measure average HFD after focus run"
         Height          =   372
         Left            =   -73380
         TabIndex        =   212
         Top             =   4020
         Width           =   3075
      End
      Begin VB.ListBox lstFiltersFocusOffset 
         Enabled         =   0   'False
         Height          =   1620
         ItemData        =   "frmOptions.frx":07BC
         Left            =   -73740
         List            =   "frmOptions.frx":07CF
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   1860
         Width           =   3675
      End
      Begin VB.CheckBox chkEnableFilterOffsets 
         Caption         =   "Enable Filter Offsets"
         Height          =   195
         Left            =   -72720
         TabIndex        =   208
         Top             =   1260
         Width           =   1935
      End
      Begin VB.Frame fraRotatorFlip 
         Caption         =   "During ""Move To"" Actions that Flip the Mount"
         Height          =   1395
         Left            =   -73620
         TabIndex        =   205
         Top             =   3900
         Width           =   4035
         Begin VB.OptionButton optRotatorFlip 
            Caption         =   "Maintain Rotator Angle (Rotator does not move)"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   207
            Top             =   780
            Width           =   3735
         End
         Begin VB.OptionButton optRotatorFlip 
            Caption         =   "Maintain Position Angle by Rotating 180 degrees"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   206
            Top             =   360
            Value           =   -1  'True
            Width           =   3735
         End
      End
      Begin VB.CheckBox chkRotateTheSkyFOVI 
         Caption         =   "Rotate TheSky FOVI to Match Camera PA"
         Enabled         =   0   'False
         Height          =   195
         Left            =   -73380
         TabIndex        =   204
         Top             =   5580
         Width           =   3375
      End
      Begin VB.CheckBox chkAutoDomeClose 
         Caption         =   "Roof/Dome will close autonomously via the Emergency Contact Closure"
         Height          =   495
         Left            =   3360
         TabIndex        =   203
         Top             =   3480
         Width           =   2955
      End
      Begin VB.Frame Frame7 
         Caption         =   "Monitor all these conditions for a ""good"" reading:"
         Height          =   1335
         Left            =   120
         TabIndex        =   194
         Top             =   4800
         Width           =   6255
         Begin VB.ListBox lstLightSensorResumeActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0822
            Left            =   4740
            List            =   "frmOptions.frx":082F
            Style           =   1  'Checkbox
            TabIndex        =   198
            Top             =   480
            Width           =   1395
         End
         Begin VB.ListBox lstRainSensorResumeActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0849
            Left            =   3240
            List            =   "frmOptions.frx":0853
            Style           =   1  'Checkbox
            TabIndex        =   197
            Top             =   480
            Width           =   1395
         End
         Begin VB.ListBox lstWindSensorResumeActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0865
            Left            =   1680
            List            =   "frmOptions.frx":0872
            Style           =   1  'Checkbox
            TabIndex        =   196
            Top             =   480
            Width           =   1395
         End
         Begin VB.ListBox lstCloudSensorResumeActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":088C
            Left            =   120
            List            =   "frmOptions.frx":0899
            Style           =   1  'Checkbox
            TabIndex        =   195
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label lblLightConditions 
            AutoSize        =   -1  'True
            Caption         =   "Light Conditions:"
            Height          =   195
            Index           =   2
            Left            =   4740
            TabIndex        =   202
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label lblRainConditions 
            AutoSize        =   -1  'True
            Caption         =   "Rain Conditions:"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   201
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblWindConditions 
            AutoSize        =   -1  'True
            Caption         =   "Wind Conditions:"
            Height          =   195
            Index           =   2
            Left            =   1680
            TabIndex        =   200
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblCloudCondition 
            AutoSize        =   -1  'True
            Caption         =   "Cloud Conditions:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   199
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Close Roof/Dome When:"
         Height          =   2415
         Left            =   3300
         TabIndex        =   181
         Top             =   1020
         Width           =   3075
         Begin VB.ListBox lstLightSensorCloseDomeWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":08B5
            Left            =   1620
            List            =   "frmOptions.frx":08C2
            Style           =   1  'Checkbox
            TabIndex        =   189
            Top             =   1560
            Width           =   1395
         End
         Begin VB.ListBox lstRainSensorCloseDomeWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":08E2
            Left            =   60
            List            =   "frmOptions.frx":08EF
            Style           =   1  'Checkbox
            TabIndex        =   188
            Top             =   1560
            Width           =   1395
         End
         Begin VB.ListBox lstWindSensorCloseDomeWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0907
            Left            =   1620
            List            =   "frmOptions.frx":0914
            Style           =   1  'Checkbox
            TabIndex        =   187
            Top             =   540
            Width           =   1395
         End
         Begin VB.ListBox lstCloudSensorCloseDomeWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0934
            Left            =   60
            List            =   "frmOptions.frx":0941
            Style           =   1  'Checkbox
            TabIndex        =   186
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label lblLightConditions 
            AutoSize        =   -1  'True
            Caption         =   "Light Conditions:"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   193
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label lblRainConditions 
            AutoSize        =   -1  'True
            Caption         =   "Rain Conditions:"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   192
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label lblWindConditions 
            AutoSize        =   -1  'True
            Caption         =   "Wind Conditions:"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   191
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblCloudCondition 
            AutoSize        =   -1  'True
            Caption         =   "Cloud Conditions:"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   190
            Top             =   300
            Width           =   1230
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Send E-Mail Alert When:"
         Height          =   1875
         Left            =   -74880
         TabIndex        =   175
         Top             =   4260
         Width           =   6315
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Guide Star Faded"
            Height          =   195
            Index           =   10
            Left            =   3360
            TabIndex        =   263
            Top             =   780
            Width           =   2415
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Comment Actions"
            Height          =   195
            Index           =   9
            Left            =   3360
            TabIndex        =   238
            Top             =   1500
            Width           =   2775
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Dome Operation Failed"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   232
            Top             =   1260
            Width           =   2415
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Action List Complete"
            Height          =   195
            Index           =   7
            Left            =   3360
            TabIndex        =   86
            Top             =   1260
            Width           =   2775
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Generic Error - Action List Aborted"
            Height          =   195
            Index           =   6
            Left            =   3360
            TabIndex        =   85
            Top             =   1020
            Width           =   2775
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Guide Star Failed to Center"
            Height          =   195
            Index           =   5
            Left            =   3360
            TabIndex        =   84
            Top             =   540
            Width           =   2415
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Guide Star Acquisition Failed"
            Height          =   195
            Index           =   4
            Left            =   3360
            TabIndex        =   83
            Top             =   300
            Width           =   2415
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Plate Solve Failed"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   82
            Top             =   1020
            Width           =   2415
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Weather Monitor - Resuming"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   81
            Top             =   780
            Width           =   2415
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Weather Monitor - Dome Closed"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   80
            Top             =   540
            Width           =   3015
         End
         Begin VB.CheckBox chkEMailAlert 
            Caption         =   "Weather Monitor - Action List Paused"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   79
            Top             =   300
            Width           =   3135
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "E-Mail Addresses"
         Height          =   1755
         Left            =   -74880
         TabIndex        =   172
         Top             =   2400
         Width           =   6315
         Begin VB.CommandButton cmdSendTestEMail 
            Caption         =   "Send Test E-Mail"
            Height          =   315
            Left            =   3180
            TabIndex        =   78
            Top             =   1260
            Width           =   1575
         End
         Begin VB.CommandButton cmdRemoveToAddress 
            Caption         =   "Remove"
            Height          =   315
            Left            =   420
            TabIndex        =   77
            Top             =   1260
            Width           =   975
         End
         Begin VB.CommandButton cmdAddToAddress 
            Caption         =   "Add"
            Height          =   315
            Left            =   420
            TabIndex        =   76
            Top             =   900
            Width           =   975
         End
         Begin VB.ListBox lstToAddresses 
            Height          =   645
            Left            =   1860
            TabIndex        =   75
            Top             =   540
            Width           =   4275
         End
         Begin VB.TextBox txtFromAddress 
            Height          =   285
            Left            =   1860
            TabIndex        =   74
            Top             =   240
            Width           =   4275
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E-Mail ""To"" Addresses:"
            Height          =   195
            Left            =   150
            TabIndex        =   174
            Top             =   600
            Width           =   1650
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E-Mail Return Address:"
            Height          =   195
            Left            =   180
            TabIndex        =   173
            Top             =   300
            Width           =   1620
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SMTP Server Setup"
         Height          =   1275
         Left            =   -74880
         TabIndex        =   167
         Top             =   1020
         Width           =   6315
         Begin VB.TextBox txtSMTPPassword 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            TabIndex        =   73
            Top             =   900
            Width           =   1995
         End
         Begin VB.TextBox txtSMTPUsername 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   72
            Top             =   900
            Width           =   1995
         End
         Begin VB.CheckBox chkAuthentication 
            Caption         =   "Use SMTP Authentication"
            Height          =   255
            Left            =   180
            TabIndex        =   71
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox txtSMTPPort 
            Height          =   285
            Left            =   5700
            TabIndex        =   70
            Text            =   "25"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtSMTPServer 
            Height          =   285
            Left            =   2700
            TabIndex        =   69
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblPassword 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3270
            TabIndex        =   171
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblUsername 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Username:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   170
            Top             =   960
            Width           =   765
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Port #:"
            Height          =   195
            Left            =   5160
            TabIndex        =   169
            Top             =   300
            Width           =   480
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SMTP Server Name or IP Address:"
            Height          =   195
            Left            =   180
            TabIndex        =   168
            Top             =   300
            Width           =   2460
         End
      End
      Begin VB.CheckBox chkMaxImCompression 
         Caption         =   "Use MaxIm DL FITS Compression"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   1920
         Width           =   3555
      End
      Begin VB.CheckBox chkReverseRotatorDirection 
         Caption         =   "Reverse rotator direction"
         Height          =   195
         Left            =   -72780
         TabIndex        =   39
         Top             =   3540
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.CheckBox chkDisableDecComp 
         Caption         =   "Disable Guider Declination Compensation"
         Height          =   315
         Left            =   -72000
         TabIndex        =   46
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Get"
         Height          =   495
         Left            =   -70680
         TabIndex        =   19
         Top             =   1680
         Width           =   735
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
         Left            =   -69060
         MaskColor       =   &H00D8E9EC&
         Picture         =   "frmOptions.frx":0963
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1410
         UseMaskColor    =   -1  'True
         Width           =   275
      End
      Begin VB.CheckBox chkIgnoreNorthAngle 
         Caption         =   "Disable North Angle Verification"
         Height          =   195
         Left            =   -73260
         TabIndex        =   20
         Top             =   2340
         Width           =   2595
      End
      Begin VB.Frame fraPinPointLESetup 
         Caption         =   "PinPoint Full && LE Setup"
         Enabled         =   0   'False
         Height          =   855
         Left            =   -73260
         TabIndex        =   156
         Top             =   4620
         Width           =   2655
         Begin VB.TextBox txtPinPointLETimeOut 
            Height          =   285
            Left            =   1380
            TabIndex        =   29
            Text            =   "5"
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkPinPointLERetry 
            Caption         =   "Retry timed out attempts"
            Height          =   255
            Left            =   300
            TabIndex        =   30
            Top             =   540
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            Caption         =   "Maximum Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   158
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "min"
            Height          =   195
            Left            =   2040
            TabIndex        =   157
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdRotatorConfigure 
         Caption         =   "Configure"
         Height          =   315
         Left            =   -70740
         TabIndex        =   12
         Top             =   5340
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkDarkSubtractPlateSolveImage 
         Caption         =   "Dark subtract all plate solve images."
         Height          =   375
         Left            =   -73320
         TabIndex        =   31
         Top             =   5700
         Width           =   2895
      End
      Begin VB.CommandButton cmdASCOMDomeConfigurre 
         Caption         =   "Configure"
         Height          =   315
         Left            =   -70740
         TabIndex        =   14
         Top             =   5700
         Width           =   1335
      End
      Begin VB.CommandButton cmdASCOMMountConfigurre 
         Caption         =   "Configure"
         Height          =   315
         Left            =   -70740
         TabIndex        =   8
         Top             =   4260
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox lstPlanetarium 
         Height          =   315
         ItemData        =   "frmOptions.frx":0AAF
         Left            =   -72840
         List            =   "frmOptions.frx":0AB1
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4620
         Width           =   2055
      End
      Begin VB.CheckBox chkInternalGuider 
         Caption         =   "Guider is Internal Guider (SBIG Cameras)"
         Height          =   195
         Left            =   -72840
         TabIndex        =   6
         Top             =   3900
         Width           =   3315
      End
      Begin VB.CheckBox chkUncoupleDomeDuringSlews 
         Caption         =   "Uncouple the Dome during slews"
         Height          =   375
         Left            =   -73200
         TabIndex        =   68
         Top             =   1380
         Width           =   2715
      End
      Begin VB.TextBox txtMaxPointingError 
         Height          =   315
         Left            =   -70740
         TabIndex        =   56
         Text            =   "3"
         Top             =   5760
         Width           =   375
      End
      Begin VB.CheckBox chkVerifyTeleCoords 
         Caption         =   "Verify telescope coordinates after Move To action."
         Height          =   255
         Left            =   -73680
         TabIndex        =   55
         Top             =   5520
         Value           =   1  'Checked
         Width           =   3915
      End
      Begin VB.CheckBox chkEnable 
         Caption         =   "Enable Automatic Guide Exposure Mode"
         Height          =   372
         Left            =   -73440
         TabIndex        =   57
         Top             =   1020
         Width           =   3252
      End
      Begin VB.TextBox txtMinExp 
         Height          =   312
         Left            =   -71760
         TabIndex        =   58
         Text            =   "1"
         ToolTipText     =   "Minimum guide exposure time"
         Top             =   1380
         Width           =   750
      End
      Begin VB.TextBox txtMaxExp 
         Height          =   312
         Left            =   -71760
         TabIndex        =   59
         Text            =   "30"
         ToolTipText     =   "Maximum guide exposure time"
         Top             =   2100
         Width           =   750
      End
      Begin VB.TextBox txtMinBright 
         Height          =   285
         Left            =   -71760
         TabIndex        =   60
         Text            =   "3000"
         ToolTipText     =   "Minimum guide star brightness value"
         Top             =   2460
         Width           =   750
      End
      Begin VB.TextBox txtMaxBright 
         Height          =   312
         Left            =   -71760
         TabIndex        =   61
         Text            =   "20000"
         ToolTipText     =   "Maximum guide star brightness value"
         Top             =   2820
         Width           =   750
      End
      Begin VB.TextBox txtGuideBoxX 
         Height          =   312
         Left            =   -71760
         TabIndex        =   62
         Text            =   "32"
         ToolTipText     =   "Guide box ""X"" pixel size"
         Top             =   3180
         Width           =   750
      End
      Begin VB.TextBox txtGuideBoxY 
         Height          =   312
         Left            =   -71760
         TabIndex        =   63
         Text            =   "32"
         ToolTipText     =   "Guide box ""Y"" pixel size"
         Top             =   3540
         Width           =   750
      End
      Begin VB.TextBox txtMaxStarMovement 
         Height          =   312
         Left            =   -71460
         TabIndex        =   64
         Text            =   "1"
         ToolTipText     =   "Maximum star movement between exposures"
         Top             =   4305
         Width           =   315
      End
      Begin VB.CheckBox chkContinuousAutoguide 
         Caption         =   "Continuous Autoguiding"
         Height          =   255
         Left            =   -73500
         TabIndex        =   65
         Top             =   5940
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mount Type"
         Height          =   915
         Left            =   -74820
         TabIndex        =   132
         Top             =   1020
         Width           =   2715
         Begin VB.OptionButton optMountType 
            Caption         =   "German Equatorial Mount"
            Height          =   372
            Index           =   0
            Left            =   300
            TabIndex        =   43
            ToolTipText     =   "Select your mount type"
            Top             =   180
            Value           =   -1  'True
            Width           =   2172
         End
         Begin VB.OptionButton optMountType 
            Caption         =   "Fork Mount"
            Height          =   375
            Index           =   1
            Left            =   300
            TabIndex        =   44
            Top             =   480
            Width           =   2172
         End
      End
      Begin VB.TextBox txtDelayAfterSlew 
         Height          =   288
         Left            =   -70560
         TabIndex        =   45
         Text            =   "0"
         ToolTipText     =   "GEM Eastern Limit"
         Top             =   1020
         Width           =   432
      End
      Begin VB.Frame fraGEMSetup 
         Caption         =   "GEM Setup"
         Height          =   2115
         Left            =   -74700
         TabIndex        =   126
         Top             =   1980
         Width           =   5835
         Begin VB.TextBox txtEasternLimit 
            Height          =   288
            Left            =   2580
            TabIndex        =   47
            Text            =   "30"
            ToolTipText     =   "GEM Eastern Limit"
            Top             =   300
            Width           =   432
         End
         Begin VB.TextBox txtWesternLimit 
            Height          =   288
            Left            =   2580
            TabIndex        =   48
            Text            =   "30"
            ToolTipText     =   "GEM Western Limit"
            Top             =   720
            Width           =   432
         End
         Begin VB.OptionButton optGuiderCal 
            Caption         =   "Eastern Sky"
            Height          =   252
            Index           =   0
            Left            =   4200
            TabIndex        =   50
            ToolTipText     =   "Side of the meridian where the autoguider was calibrated"
            Top             =   1560
            Value           =   -1  'True
            Width           =   1392
         End
         Begin VB.OptionButton optGuiderCal 
            Caption         =   "Western Sky"
            Height          =   252
            Index           =   1
            Left            =   4200
            TabIndex        =   51
            Top             =   1800
            Width           =   1212
         End
         Begin VB.CheckBox chkAutoDetermineMountSide 
            Caption         =   "Automatically determine mount side at start-up"
            Height          =   375
            Left            =   1020
            TabIndex        =   49
            Top             =   1140
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Eastern Limit (east side):"
            Height          =   195
            Left            =   540
            TabIndex        =   131
            Top             =   360
            Width           =   1920
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "minutes past the meridian"
            Height          =   195
            Left            =   3120
            TabIndex        =   130
            Top             =   360
            Width           =   2145
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Western Limit (west side):"
            Height          =   195
            Left            =   600
            TabIndex        =   129
            Top             =   780
            Width           =   1860
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "minutes past the meridian"
            Height          =   195
            Left            =   3120
            TabIndex        =   128
            Top             =   780
            Width           =   1845
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Autoguider calibrated in (Rotator Angles computed in):"
            Height          =   195
            Left            =   240
            TabIndex        =   127
            Top             =   1680
            Width           =   3900
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraScripts 
         Caption         =   "Slew Scripts"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   -74700
         TabIndex        =   121
         Top             =   4380
         Width           =   5835
         Begin VB.CommandButton cmdOpenBefore 
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
            Left            =   5460
            MaskColor       =   &H00D8E9EC&
            Picture         =   "frmOptions.frx":0AB3
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   310
            UseMaskColor    =   -1  'True
            Width           =   275
         End
         Begin VB.CommandButton cmdOpenAfter 
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
            Left            =   5460
            MaskColor       =   &H00D8E9EC&
            Picture         =   "frmOptions.frx":0BFF
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   275
         End
         Begin VB.TextBox txtAfterScript 
            BackColor       =   &H8000000F&
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   123
            Top             =   720
            Width           =   4275
         End
         Begin VB.TextBox txtBeforeScript 
            BackColor       =   &H8000000F&
            Height          =   288
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   122
            Top             =   300
            Width           =   4275
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "After Slew:"
            Height          =   195
            Left            =   330
            TabIndex        =   125
            Top             =   750
            Width           =   765
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Before Slew:"
            Height          =   195
            Left            =   195
            TabIndex        =   124
            Top             =   325
            Width           =   900
         End
      End
      Begin VB.CheckBox chkEnableScripts 
         Caption         =   "Enable Slew Scripts"
         Height          =   255
         Left            =   -74700
         TabIndex        =   52
         Top             =   4140
         Width           =   1815
      End
      Begin VB.TextBox txtSaveTo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   -74820
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Text            =   "C:\"
         Top             =   1380
         Width           =   5715
      End
      Begin VB.TextBox txtQuerySensorPeriod 
         Height          =   315
         Left            =   1800
         TabIndex        =   41
         Text            =   "5"
         Top             =   4080
         Width           =   435
      End
      Begin VB.TextBox txtClearTime 
         Height          =   315
         Left            =   5340
         TabIndex        =   42
         Text            =   "30"
         Top             =   4140
         Width           =   435
      End
      Begin VB.CheckBox chkParkMountWhenCloudy 
         Caption         =   "Park Mount when Pausing Action"
         Height          =   255
         Left            =   300
         TabIndex        =   40
         Top             =   3600
         Width           =   2715
      End
      Begin VB.TextBox txtHomeRotationAngle 
         Height          =   315
         Left            =   -72840
         TabIndex        =   33
         Text            =   "320.45"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtCOMNum 
         Height          =   315
         Left            =   -72840
         TabIndex        =   32
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtGuiderCalAngle 
         Height          =   315
         Left            =   -72840
         TabIndex        =   34
         Text            =   "320.45"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkRotFromSky 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   -70620
         TabIndex        =   36
         Top             =   1680
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkRotFromSky 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   -70620
         TabIndex        =   37
         Top             =   2040
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton cmdGetAngleFromTheSky 
         Caption         =   "Get Angle from TheSky"
         Enabled         =   0   'False
         Height          =   795
         Left            =   -70380
         TabIndex        =   38
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox chkGuiderRotates 
         Caption         =   "Guider rotates with Imager"
         Height          =   195
         Left            =   -72780
         TabIndex        =   35
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.TextBox txtPixelScale 
         Height          =   285
         Left            =   -72240
         TabIndex        =   17
         Text            =   "4.01"
         Top             =   1620
         Width           =   615
      End
      Begin VB.TextBox txtNorthAngle 
         Height          =   285
         Left            =   -72240
         TabIndex        =   18
         Text            =   "354"
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox lstPlateSolve 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D4B
         Left            =   -73020
         List            =   "frmOptions.frx":0D4D
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Frame fraPinPointSetup 
         Caption         =   "Pin Point (Full) Setup"
         Enabled         =   0   'False
         Height          =   1935
         Left            =   -74880
         TabIndex        =   102
         Top             =   2580
         Width           =   6195
         Begin VB.TextBox txtMaxNumStars 
            Height          =   285
            Left            =   4920
            TabIndex        =   28
            Text            =   "200"
            Top             =   1020
            Width           =   555
         End
         Begin VB.TextBox txtStandardDeviation 
            Height          =   285
            Left            =   3000
            TabIndex        =   24
            Text            =   "4.00"
            Top             =   1500
            Width           =   555
         End
         Begin VB.TextBox txtMinStarBrightness 
            Height          =   285
            Left            =   1980
            TabIndex        =   23
            Text            =   "200"
            Top             =   1140
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtSearchArea 
            Height          =   285
            Left            =   4920
            TabIndex        =   27
            Text            =   "256.00"
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtCatalogMagMin 
            Height          =   285
            Left            =   5640
            TabIndex        =   26
            Text            =   "20.0"
            Top             =   300
            Width           =   435
         End
         Begin VB.TextBox txtCatalogMagMax 
            Height          =   285
            Left            =   4920
            TabIndex        =   25
            Text            =   "-2.0"
            Top             =   300
            Width           =   435
         End
         Begin VB.ComboBox lstCatalog 
            Height          =   315
            ItemData        =   "frmOptions.frx":0D4F
            Left            =   720
            List            =   "frmOptions.frx":0D51
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   300
            Width           =   1755
         End
         Begin VB.CommandButton cmdOpenCatalogPath 
            Caption         =   "..."
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
            Left            =   2280
            TabIndex        =   22
            Top             =   675
            Width           =   275
         End
         Begin VB.TextBox txtCatalogPath 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   103
            TabStop         =   0   'False
            Text            =   "C:\"
            Top             =   660
            Width           =   1515
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Maximum # of stars used:"
            Height          =   195
            Left            =   3045
            TabIndex        =   166
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Standard deviation above background:"
            Height          =   195
            Left            =   180
            TabIndex        =   165
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "ADU"
            Height          =   195
            Left            =   2580
            TabIndex        =   164
            Top             =   1200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Minimum star brightness:"
            Height          =   195
            Left            =   180
            TabIndex        =   163
            Top             =   1200
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   5580
            TabIndex        =   162
            Top             =   720
            Width           =   120
         End
         Begin VB.Label Label60 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Search Area (% of image):"
            Height          =   195
            Left            =   3030
            TabIndex        =   161
            Top             =   720
            Width           =   1830
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "to"
            Height          =   195
            Left            =   5400
            TabIndex        =   160
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Catalog Stellar Magnitudes:"
            Height          =   195
            Left            =   2925
            TabIndex        =   159
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Catalog:"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Path:"
            Height          =   195
            Left            =   300
            TabIndex        =   104
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.ComboBox lstCameraControl 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D53
         Left            =   -72840
         List            =   "frmOptions.frx":0D55
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1140
         Width           =   2055
      End
      Begin VB.ComboBox lstMountControl 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D57
         Left            =   -72840
         List            =   "frmOptions.frx":0D59
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4260
         Width           =   2055
      End
      Begin VB.ComboBox lstFocuserControl 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D5B
         Left            =   -72840
         List            =   "frmOptions.frx":0D5D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4980
         Width           =   2055
      End
      Begin VB.ComboBox lstRotator 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D5F
         Left            =   -72840
         List            =   "frmOptions.frx":0D61
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   5340
         Width           =   2055
      End
      Begin VB.ListBox lstFilters 
         Height          =   1035
         ItemData        =   "frmOptions.frx":0D63
         Left            =   -72840
         List            =   "frmOptions.frx":0D76
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1500
         Width           =   2055
      End
      Begin VB.CommandButton cmdGetFilters 
         Caption         =   "Get Filters from CCDSoft"
         Height          =   555
         Left            =   -70740
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddFilter 
         Caption         =   "Add Filter"
         Height          =   555
         Left            =   -70740
         TabIndex        =   4
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdDeleteFilter 
         Caption         =   "Delete Filter"
         Height          =   555
         Left            =   -70020
         TabIndex        =   5
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox lstDomeControl 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D97
         Left            =   -72840
         List            =   "frmOptions.frx":0D99
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   5700
         Width           =   2055
      End
      Begin VB.ComboBox lstCloudSensor 
         Height          =   315
         ItemData        =   "frmOptions.frx":0D9B
         Left            =   -72840
         List            =   "frmOptions.frx":0D9D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   6060
         Width           =   2055
      End
      Begin VB.Frame Frame5 
         Caption         =   "Pause Action When:"
         Height          =   2415
         Left            =   120
         TabIndex        =   176
         Top             =   1020
         Width           =   3075
         Begin VB.ListBox lstLightSensorPauseActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0D9F
            Left            =   1620
            List            =   "frmOptions.frx":0DAC
            Style           =   1  'Checkbox
            TabIndex        =   180
            Top             =   1560
            Width           =   1395
         End
         Begin VB.ListBox lstRainSensorPauseActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0DCC
            Left            =   60
            List            =   "frmOptions.frx":0DD9
            Style           =   1  'Checkbox
            TabIndex        =   179
            Top             =   1560
            Width           =   1395
         End
         Begin VB.ListBox lstWindSensorPauseActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0DF1
            Left            =   1620
            List            =   "frmOptions.frx":0DFE
            Style           =   1  'Checkbox
            TabIndex        =   178
            Top             =   540
            Width           =   1395
         End
         Begin VB.ListBox lstCloudSensorPauseActionWhen 
            Height          =   735
            ItemData        =   "frmOptions.frx":0E1E
            Left            =   60
            List            =   "frmOptions.frx":0E2B
            Style           =   1  'Checkbox
            TabIndex        =   177
            Top             =   540
            Width           =   1395
         End
         Begin VB.Label lblLightConditions 
            AutoSize        =   -1  'True
            Caption         =   "Light Conditions:"
            Height          =   195
            Index           =   0
            Left            =   1620
            TabIndex        =   185
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label lblRainConditions 
            AutoSize        =   -1  'True
            Caption         =   "Rain Conditions:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   184
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label lblWindConditions 
            AutoSize        =   -1  'True
            Caption         =   "Wind Conditions:"
            Height          =   195
            Index           =   0
            Left            =   1620
            TabIndex        =   183
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblCloudCondition 
            AutoSize        =   -1  'True
            Caption         =   "Cloud Conditions:"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   182
            Top             =   300
            Width           =   1230
         End
      End
      Begin VB.Label lblMaximumStarFaded 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum Star Faded Errors:"
         Height          =   195
         Left            =   -72855
         TabIndex        =   272
         Top             =   3000
         Width           =   1980
      End
      Begin VB.Label Label86 
         AutoSize        =   -1  'True
         Caption         =   "guide cycles."
         Height          =   195
         Left            =   -70740
         TabIndex        =   270
         Top             =   5220
         Width           =   930
      End
      Begin VB.Label Label85 
         AutoSize        =   -1  'True
         Caption         =   "pixels, for "
         Height          =   195
         Left            =   -71880
         TabIndex        =   268
         Top             =   5220
         Width           =   705
      End
      Begin VB.Label Label84 
         Caption         =   "is more than +/-"
         Height          =   195
         Left            =   -73500
         TabIndex        =   266
         Top             =   5220
         Width           =   1155
      End
      Begin VB.Label Label80 
         Caption         =   $"frmOptions.frx":0E4D
         Height          =   1455
         Left            =   -73980
         TabIndex        =   243
         Top             =   1260
         Width           =   4695
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "minutes"
         Height          =   195
         Left            =   -71100
         TabIndex        =   237
         Top             =   5460
         Width           =   540
      End
      Begin VB.Label Label77 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Focus Time Out:"
         Height          =   195
         Left            =   -73140
         TabIndex        =   236
         Top             =   5460
         Width           =   1170
      End
      Begin VB.Label Label51 
         Caption         =   "(Useful for guiders with many hot pixels)"
         Height          =   435
         Left            =   -70020
         TabIndex        =   234
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Retries:"
         Height          =   195
         Left            =   -73275
         TabIndex        =   225
         Top             =   4950
         Width           =   1320
      End
      Begin VB.Label Label73 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Typical Guide Star FWHM:"
         Height          =   195
         Left            =   -73770
         TabIndex        =   222
         Top             =   3960
         Width           =   1905
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   -70905
         TabIndex        =   221
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label69 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum Guide Exposure Increment:"
         Height          =   195
         Left            =   -74490
         TabIndex        =   217
         Top             =   1800
         Width           =   2625
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   -70905
         TabIndex        =   216
         Top             =   1815
         Width           =   630
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Offset:"
         Height          =   195
         Left            =   -70860
         TabIndex        =   211
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Filter Name:"
         Height          =   195
         Left            =   -73740
         TabIndex        =   210
         Top             =   1620
         Width           =   840
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Planetarium:"
         Height          =   195
         Left            =   -73755
         TabIndex        =   155
         Top             =   4680
         Width           =   870
      End
      Begin VB.Label Label54 
         Caption         =   $"frmOptions.frx":0FC8
         Height          =   855
         Left            =   -74160
         TabIndex        =   154
         Top             =   1860
         Width           =   4695
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "arcminutes"
         Height          =   195
         Left            =   -70260
         TabIndex        =   153
         Top             =   5820
         Width           =   765
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Maximum acceptable telescope pointing error:"
         Height          =   195
         Left            =   -74040
         TabIndex        =   152
         Top             =   5820
         Width           =   3240
      End
      Begin VB.Label lblAngles 
         AutoSize        =   -1  'True
         Caption         =   "Angles determined when the telescope was pointing to the eastern sky."
         Height          =   315
         Left            =   -74280
         TabIndex        =   151
         Top             =   2820
         Visible         =   0   'False
         Width           =   5025
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   -70905
         TabIndex        =   150
         Top             =   1455
         Width           =   630
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   -70905
         TabIndex        =   149
         Top             =   2145
         Width           =   630
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "ADU"
         Height          =   195
         Left            =   -70905
         TabIndex        =   148
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "ADU"
         Height          =   195
         Left            =   -70905
         TabIndex        =   147
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   -70905
         TabIndex        =   146
         Top             =   3240
         Width           =   420
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   -70905
         TabIndex        =   145
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum Guide Star Brightness:"
         Height          =   195
         Left            =   -74220
         TabIndex        =   144
         Top             =   2880
         Width           =   2355
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Minimum Guide Star Brightness:"
         Height          =   195
         Left            =   -74160
         TabIndex        =   143
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum Guide Exposure:"
         Height          =   195
         Left            =   -73800
         TabIndex        =   142
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Minimum Guide Exposure:"
         Height          =   195
         Left            =   -73740
         TabIndex        =   141
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guide Box X Size:"
         Height          =   195
         Left            =   -73140
         TabIndex        =   140
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guide Box Y Size:"
         Height          =   195
         Left            =   -73140
         TabIndex        =   139
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum Star Movement between exposures:"
         Height          =   390
         Left            =   -73860
         TabIndex        =   138
         Top             =   4260
         Width           =   2070
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   -71040
         TabIndex        =   137
         Top             =   4365
         Width           =   420
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "+/-"
         Height          =   195
         Left            =   -71700
         TabIndex        =   136
         Top             =   4365
         Width           =   210
      End
      Begin VB.Label Label35 
         Caption         =   $"frmOptions.frx":1098
         Height          =   675
         Left            =   -73920
         TabIndex        =   135
         Top             =   6240
         Width           =   4395
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "seconds"
         Height          =   195
         Left            =   -70020
         TabIndex        =   134
         Top             =   1080
         Width           =   1125
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Delay after slew:"
         Height          =   195
         Left            =   -71940
         TabIndex        =   133
         Top             =   1080
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Global Image Save Location:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   120
         Top             =   1140
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Query Sensor every:"
         Height          =   195
         Left            =   300
         TabIndex        =   118
         Top             =   4140
         Width           =   1440
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "minutes"
         Height          =   195
         Left            =   2280
         TabIndex        =   117
         Top             =   4140
         Width           =   540
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Before resuming action, sensors must read ""good"" for:"
         Height          =   435
         Left            =   3180
         TabIndex        =   116
         Top             =   4020
         Width           =   2130
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "minutes"
         Height          =   195
         Left            =   5820
         TabIndex        =   115
         Top             =   4200
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "degrees from North"
         Height          =   195
         Left            =   -72180
         TabIndex        =   114
         Top             =   1740
         Width           =   1350
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Home Rotation Angle:"
         Height          =   195
         Left            =   -74460
         TabIndex        =   113
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rotator COM:"
         Height          =   195
         Left            =   -73860
         TabIndex        =   112
         Top             =   1380
         Width           =   975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "degrees from North"
         Height          =   195
         Left            =   -72180
         TabIndex        =   111
         Top             =   2100
         Width           =   1350
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Guider Calibrated at:"
         Height          =   195
         Left            =   -74325
         TabIndex        =   110
         Top             =   2100
         Width           =   1440
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Unbinned Pixel Scale:"
         Height          =   375
         Left            =   -73200
         TabIndex        =   109
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "arcsec/pix"
         Height          =   195
         Left            =   -71580
         TabIndex        =   108
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "North Angle:"
         Height          =   195
         Left            =   -73200
         TabIndex        =   107
         Top             =   1950
         Width           =   885
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "degrees"
         Height          =   195
         Left            =   -71580
         TabIndex        =   106
         Top             =   1950
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Camera Control:"
         Height          =   195
         Left            =   -74040
         TabIndex        =   101
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mount Control:"
         Height          =   195
         Left            =   -73920
         TabIndex        =   100
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Focuser Control:"
         Height          =   195
         Left            =   -74040
         TabIndex        =   99
         Top             =   5040
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Instrument Rotator:"
         Height          =   195
         Left            =   -74235
         TabIndex        =   98
         Top             =   5400
         Width           =   1350
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Filters:"
         Height          =   195
         Left            =   -73440
         TabIndex        =   97
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dome/Roof Control:"
         Height          =   195
         Left            =   -74310
         TabIndex        =   96
         Top             =   5760
         Width           =   1425
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Weather Monitor"
         Height          =   195
         Left            =   -74070
         TabIndex        =   95
         Top             =   6120
         Width           =   1185
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   94
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   93
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   92
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   88
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6960
      TabIndex        =   87
      Top             =   120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog MSComm 
      Left            =   6900
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuSaveSettings 
      Caption         =   "Save Settings"
   End
   Begin VB.Menu mnuLoadSettings 
      Caption         =   "Load Settings"
   End
End
Attribute VB_Name = "frmOptions"
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
Private Const RegistryName = "ProgramSettings"
Private Const MountRegistryName = "MountParameters"
Private Const GuiderRegistryName = "AutoGuideExposure"

Public Enum EMailAlertIndexes
    WeatherMonitorActionListPaused = 0
    WeatherMonitorDomeClosed = 1
    WeatherMonitorResuming = 2
    PlateSolveFailed = 3
    GuideStarAcquisitionFailed = 4
    GuideStarFailedToCenter = 5
    GenericError = 6
    ActionListComplete = 7
    DomeOpFailed = 8
    CommentActions = 9
    GuideStarFaded = 10
End Enum

Private Sub CancelButton_Click()
    Me.Hide
    Call Form_Load
End Sub

Private Sub chkAuthentication_Click()
    If Me.chkAuthentication.Value = vbChecked Then
        Me.lblUsername.Enabled = True
        Me.txtSMTPUsername.Enabled = True
        Me.lblPassword.Enabled = True
        Me.txtSMTPPassword.Enabled = True
    Else
        Me.lblUsername.Enabled = False
        Me.txtSMTPUsername.Enabled = False
        Me.lblPassword.Enabled = False
        Me.txtSMTPPassword.Enabled = False
    End If
End Sub

Private Sub chkAutoDomeClose_Click()
    If Me.chkAutoDomeClose.Value = vbChecked Then
        If Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityII Or _
        Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityIIRemote Or _
        Me.lstCloudSensor.ListIndex = WeatherMonitorControl.AAG Or _
        Me.lstCloudSensor.ListIndex = WeatherMonitorControl.AAGRemote Then
            'Clarity I has a hard coded "Very Cloudy" level that may be worse than the software programmed level
            'So allow the Very Cloudy condition to be enabled for CCD Commander to close the dome
            Me.lstCloudSensorCloseDomeWhen.Selected(2) = False
        End If
        Me.lstRainSensorCloseDomeWhen.Selected(1) = False
        Me.lstRainSensorCloseDomeWhen.Selected(2) = False
        Me.lstWindSensorCloseDomeWhen.Selected(2) = False
        Me.lstLightSensorCloseDomeWhen.Selected(2) = False
    End If
End Sub

Private Sub chkEnableFilterOffsets_Click()
    If Me.chkEnableFilterOffsets.Value = vbChecked Then
        Me.lstFiltersFocusOffset.Enabled = True
    Else
        Me.lstFiltersFocusOffset.Enabled = False
    End If
End Sub

Private Sub chkEnableWeatherMonitorScripts_Click()
    If Me.chkEnableWeatherMonitorScripts.Value = vbChecked Then
        Me.fraWeatherMonitorScripts.Enabled = True
    Else
        Me.fraWeatherMonitorScripts.Enabled = False
    End If
End Sub

Private Sub chkGuiderRotates_Click()
    If Me.chkGuiderRotates.Value = vbUnchecked Then
        Me.txtGuiderCalAngle.Enabled = False
        Me.txtGuiderCalAngle.Text = "0"
        Me.chkRotFromSky(1).Enabled = False
        Me.chkRotFromSky(1).Value = vbUnchecked
    Else
        Me.txtGuiderCalAngle.Enabled = True
        Me.chkRotFromSky(1).Enabled = True
        Me.chkRotFromSky(1).Value = vbChecked
    End If
End Sub

Private Sub chkParkMountFirst_Click()
    Dim i As Integer
    
    If Me.chkParkMountFirst.Value = vbChecked Then
        Me.chkAutoDomeClose.Enabled = False
        Me.chkAutoDomeClose.Value = vbUnchecked
    
        If Me.lstCloudSensor.ListIndex > 0 Then
            For i = 0 To Me.lstCloudSensorCloseDomeWhen.ListCount - 1
                Me.lstCloudSensorPauseActionWhen.Selected(i) = Me.lstCloudSensorPauseActionWhen.Selected(i) Or Me.lstCloudSensorCloseDomeWhen.Selected(i)
            Next i
        
            For i = 0 To Me.lstRainSensorCloseDomeWhen.ListCount - 1
                Me.lstRainSensorPauseActionWhen.Selected(i) = Me.lstRainSensorPauseActionWhen.Selected(i) Or Me.lstRainSensorCloseDomeWhen.Selected(i)
            Next i
        
            For i = 0 To Me.lstWindSensorCloseDomeWhen.ListCount - 1
                Me.lstWindSensorPauseActionWhen.Selected(i) = Me.lstWindSensorPauseActionWhen.Selected(i) Or Me.lstWindSensorCloseDomeWhen.Selected(i)
            Next i
        
            For i = 0 To Me.lstLightSensorCloseDomeWhen.ListCount - 1
                Me.lstLightSensorPauseActionWhen.Selected(i) = Me.lstLightSensorPauseActionWhen.Selected(i) Or Me.lstLightSensorCloseDomeWhen.Selected(i)
            Next i
        End If
    Else
        Me.chkAutoDomeClose.Enabled = True
    End If
End Sub

Private Sub chkRetryFocusRunOnFailure_Click()
    If Me.chkRetryFocusRunOnFailure.Value = vbChecked Then
        Me.txtFocusRetryCount.Enabled = True
    Else
        Me.txtFocusRetryCount.Enabled = False
    End If
End Sub

Private Sub chkVerifyTeleCoords_Click()
    If Me.chkVerifyTeleCoords.Value = vbChecked Then
        Me.txtMaxPointingError.Enabled = True
    Else
        Me.txtMaxPointingError.Enabled = False
    End If
End Sub

Private Sub cmdAddFilter_Click()
    Dim FilterDescription As String
    
    FilterDescription = InputBox("Enter the description for filter #" & Me.lstFilters.ListCount + 1, "Adding Filter #" & Me.lstFilters.ListCount + 1, "Filter " & Me.lstFilters.ListCount + 1)

    If FilterDescription <> "" Then
        Me.lstFilters.AddItem FilterDescription
        FilterDescription = Left(FilterDescription, 55)
        Me.lstFiltersFocusOffset.AddItem FilterDescription & Space(55 - Len(FilterDescription)) & vbTab & "0"
    End If
End Sub

Private Sub cmdAddToAddress_Click()
    Dim EmailAddress As String
    
    EmailAddress = InputBox("Enter the e-mail address to add.", "Adding 'To' E-mail Address")

    If EmailAddress <> "" Then
        Me.lstToAddresses.AddItem EmailAddress
    End If
End Sub

Private Sub cmdAfterCloseScript_Click()
    Me.MSComm.DialogTitle = "Weather Monitor After Close Slew Script"
    Me.MSComm.FileName = Me.txtAfterCloseScript.Text
    Me.MSComm.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.CancelError = True
    Me.MSComm.InitDir = Me.txtAfterCloseScript.Text
    Me.MSComm.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtAfterCloseScript.Text = Me.MSComm.FileName
    Else
        Me.txtAfterCloseScript.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub cmdAfterGoodScript_Click()
    Me.MSComm.DialogTitle = "Weather Monitor After Good Slew Script"
    Me.MSComm.FileName = Me.txtAfterGoodScript.Text
    Me.MSComm.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.CancelError = True
    Me.MSComm.InitDir = Me.txtAfterGoodScript.Text
    Me.MSComm.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtAfterGoodScript.Text = Me.MSComm.FileName
    Else
        Me.txtAfterGoodScript.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub cmdAfterPauseScript_Click()
    Me.MSComm.DialogTitle = "Weather Monitor After Pause Slew Script"
    Me.MSComm.FileName = Me.txtAfterPauseScript.Text
    Me.MSComm.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.CancelError = True
    Me.MSComm.InitDir = Me.txtAfterPauseScript.Text
    Me.MSComm.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtAfterPauseScript.Text = Me.MSComm.FileName
    Else
        Me.txtAfterPauseScript.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub cmdASCOMDomeConfigurre_Click()
    Dim chsr As DriverHelper.Chooser
    Dim scopeProgID As String
    
    On Error GoTo ASCOMDomeError
    
    Set chsr = New DriverHelper.Chooser
    
    scopeProgID = GetMySetting(RegistryName, "ASCOMDomeProgID", "")
   
    ' This will be a Telescope chooser
    chsr.DeviceType = "Dome"
    ' Retrieve the ProgID of the previously chosen
    ' device, or set it to ""
    scopeProgID = chsr.Choose(scopeProgID)

    If scopeProgID = "" Then
        scopeProgID = GetMySetting(RegistryName, "ASCOMDomeProgID", scopeProgID)
    Else
        Call SaveMySetting(RegistryName, "ASCOMDomeProgID", scopeProgID)
    End If
    
    On Error GoTo 0
    Exit Sub
    
ASCOMDomeError:
    On Error GoTo 0

    Call MsgBox("Error accessing the ASCOM components." & vbCrLf & "Please reinstall ASCOM from http://www.ascom-standards.org" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical, , Err.HelpFile, Err.HelpContext)
End Sub

Private Sub cmdASCOMMountConfigurre_Click()
    Dim chsr As DriverHelper.Chooser
    Dim scopeProgID As String
    
    Call Mount.MountUnload
    
    On Error GoTo ASCOMMountError
    
    Set chsr = New DriverHelper.Chooser
    
    scopeProgID = GetMySetting(RegistryName, "ASCOMScopeProgID", "")
   
    ' This will be a Telescope chooser
    chsr.DeviceType = "Telescope"
    ' Retrieve the ProgID of the previously chosen
    ' device, or set it to ""
    scopeProgID = chsr.Choose(scopeProgID)
    
    If scopeProgID = "" Then
        scopeProgID = GetMySetting(RegistryName, "ASCOMScopeProgID", scopeProgID)
    Else
        Call SaveMySetting(RegistryName, "ASCOMScopeProgID", scopeProgID)
    End If
    
    On Error GoTo 0
    Exit Sub
    
ASCOMMountError:
    On Error GoTo 0
    
    Call MsgBox("Error accessing the ASCOM components." & vbCrLf & "Please reinstall ASCOM from http://www.ascom-standards.org" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical, , Err.HelpFile, Err.HelpContext)
End Sub

Private Sub cmdDeleteFilter_Click()
    If Me.lstFilters.ListIndex = -1 Then
        MsgBox "You must select a filter before you can delete it."
        Exit Sub
    End If
    
    Call Me.lstFiltersFocusOffset.RemoveItem(Me.lstFilters.ListIndex)
    Call Me.lstFilters.RemoveItem(Me.lstFilters.ListIndex)
End Sub

Private Sub cmdEMailScript_Click()
    Me.MSComm.DialogTitle = "Open E-Mail Script"
    Me.MSComm.FileName = Me.txtEMailScript.Text
    Me.MSComm.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.CancelError = True
    Me.MSComm.InitDir = Me.txtEMailScript.Text
    Me.MSComm.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtEMailScript.Text = Me.MSComm.FileName
    Else
        Me.txtEMailScript.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub cmdGet_Click()
    Dim objImage As Object
    
    On Error GoTo cmdGetError
    
    Set objImage = CreateObject("CCDSoft.Image")
    
    With objImage
        .AttachToActive
        
        If .ScaleInArcsecondsPerPixel = 0 Then
            Err.Raise -1
        End If
        
        Me.txtNorthAngle.Text = Format(.NorthAngle, "0.0")
        Settings.NorthAngle = CDbl(Me.txtNorthAngle.Text)
        Me.txtPixelScale.Text = Format(.ScaleInArcsecondsPerPixel / CDbl(.FITSKeyword("XBINNING")), "0.00")
        Settings.PixelScale = CDbl(Me.txtPixelScale.Text)
    End With
    
cmdGetError:
    If Err.Number <> 0 Then
        Call MsgBox("Please manually take an image in CCDSoft and perform an Insert WCS operation." & vbCrLf & "Confirm the plate solve is correct and then push the 'Get' button again.", vbInformation + vbOKOnly)
    End If
    
    On Error GoTo 0
End Sub

Private Sub cmdGetAngleFromTheSky_Click()
    Dim PositionAngle As Double
    
    PositionAngle = Planetarium.GetFOVIPositionAngle
    
    If Me.chkRotFromSky(0).Value = vbChecked Then
        Me.txtHomeRotationAngle.Text = Format(PositionAngle, "0.00")
        Settings.HomeRotationAngle = PositionAngle
    End If
    
    If Me.chkRotFromSky(1).Value = vbChecked Then
        Me.txtGuiderCalAngle.Text = Format(PositionAngle, "0.00")
        Settings.GuiderCalibrationAngle = PositionAngle
    End If
End Sub

Private Sub cmdGetFilters_Click()
    Dim Counter As Integer
    Dim Filters As Variant
    
    On Error GoTo GetFiltersError
    
    Call Camera.CameraSetup
    Filters = Camera.objCameraControl.GetFilters
    
    If Not IsNull(Filters) Then
        Me.lstFilters.Clear
        Me.lstFiltersFocusOffset.Clear
        For Counter = 0 To UBound(Filters)
            Me.lstFilters.AddItem Filters(Counter)
            
            Filters(Counter) = Left(Filters(Counter), 55)
            Me.lstFiltersFocusOffset.AddItem Filters(Counter) & Space(55 - Len(Filters(Counter))) & vbTab & "0"
            Me.lstFiltersFocusOffset.ItemData(Me.lstFiltersFocusOffset.NewIndex) = 0
        Next Counter
    End If
    
GetFiltersError:
    If (Err.Number <> 0) Then
        Call MsgBox("Error getting filters. Please check your camera control program.", vbCritical, "Get Filters Error")
    End If
    On Error GoTo 0
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

Private Sub cmdOpenCatalogPath_Click()
    Me.MSComm.Filter = "Folders|Folders"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.DialogTitle = "Select the folder of the star catalog..."
    Me.MSComm.FileName = "Select folder"
    Me.MSComm.InitDir = Me.txtCatalogPath.Text
    Me.MSComm.flags = cdlOFNHideReadOnly + cdlOFNNoValidate + cdlOFNPathMustExist
    Me.MSComm.CancelError = True
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtCatalogPath.Text = Left(Me.MSComm.FileName, InStrRev(Me.MSComm.FileName, "\"))
    End If
    On Error GoTo 0
End Sub

Private Sub cmdWeatherMonitorRemoteFileOpen_Click()
    Me.MSComm.Filter = "*.*|*.*"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.DialogTitle = "Select the weather monitor output file..."
    Me.MSComm.FileName = "Select file"
    Me.MSComm.InitDir = Me.txtWeatherMonitorRemoteFile.Text
    Me.MSComm.flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    Me.MSComm.CancelError = True
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtWeatherMonitorRemoteFile.Text = Me.MSComm.FileName
    End If
    On Error GoTo 0
End Sub

Private Sub cmdRemoveToAddress_Click()
    If Me.lstToAddresses.ListIndex = -1 Then
        MsgBox "You must select an e-mail address before you can remove it."
        Exit Sub
    End If
    
    Call Me.lstToAddresses.RemoveItem(Me.lstToAddresses.ListIndex)
End Sub

Private Sub cmdRotatorConfigure_Click()
    Dim objRotator As Object
    Dim chsr As DriverHelper.Chooser
    Dim rotatorProgID As String
    
    Call Rotator.RotatorUnload
    
    On Error GoTo RotatorConfigureError
        
    If lstRotator.ListIndex = 3 Then
        Set objRotator = CreateObject("RCOS_AE.Rotator") 'CreateObject("RcosTCC.Rotator")
        Call objRotator.SetupDialog
        Set objRotator = Nothing
    Else
        Set chsr = New DriverHelper.Chooser
        
        rotatorProgID = GetMySetting(RegistryName, "ASCOMRotatorProgID", "")
        
        ' This will be a Rotator chooser
        chsr.DeviceType = "Rotator"
        ' Retrieve the ProgID of the previously chosen
        ' device, or set it to ""
        rotatorProgID = chsr.Choose(rotatorProgID)
        
        If rotatorProgID = "" Then
            rotatorProgID = GetMySetting(RegistryName, "ASCOMRotatorProgID", rotatorProgID)
        Else
            Call SaveMySetting(RegistryName, "ASCOMRotatorProgID", rotatorProgID)
        End If
    End If
        
    On Error GoTo 0
     
    Exit Sub
    
RotatorConfigureError:
    On Error GoTo 0
    
    If lstRotator.ListIndex = 3 Then
        Call MsgBox("Error connecting to PIR Driver." & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical, , Err.HelpFile, Err.HelpContext)
    Else
        Call MsgBox("Error accessing the ASCOM components." & vbCrLf & "Please reinstall ASCOM from http://www.ascom-standards.org" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical, , Err.HelpFile, Err.HelpContext)
    End If
End Sub

Private Sub cmdSendTestEMail_Click()
    Dim Result As Boolean
    Dim TestMessage As String
    Dim FileNo As Integer
    
    If Me.lstToAddresses.ListIndex = -1 Then
        MsgBox "You must select an e-mail address to test.", vbExclamation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Me.cmdSendTestEMail.Caption = "Sending E-Mail...."
    DoEvents
    Me.Enabled = False
    
    'send e-mail
    TestMessage = "Subject: CCD Commander Test Message" & vbCrLf & "From: CCDCommander <" & Me.txtFromAddress.Text & ">" & vbCrLf & "To: " & Me.lstToAddresses.List(Me.lstToAddresses.ListIndex) & vbCrLf & vbCrLf & _
        "Hello, this is a test message." & vbCrLf & "." & vbCrLf

    FileNo = FreeFile()
    Open App.Path & "\EmailTest.log" For Output As #FileNo

    If Me.chkAuthentication.Value = vbUnchecked Then
        Result = EMail.SendEMailNoAuth(Me, Me.txtSMTPServer.Text, Settings.SMTPPort, Me.txtFromAddress.Text, Me.lstToAddresses.List(Me.lstToAddresses.ListIndex), TestMessage, FileNo)
    Else
        Result = EMail.SendEMailWithAuth(Me, Me.txtSMTPServer.Text, Settings.SMTPPort, Me.txtFromAddress.Text, Me.lstToAddresses.List(Me.lstToAddresses.ListIndex), TestMessage, Me.txtSMTPUsername.Text, Me.txtSMTPPassword.Text, FileNo)
    End If
    
    Close #FileNo
    
    Me.cmdSendTestEMail.Caption = "Send Test E-Mail"
    Me.Enabled = True
    Me.MousePointer = vbNormal
    DoEvents
    
    If Result Then
        MsgBox "Test e-mail sent to: " & Me.lstToAddresses.List(Me.lstToAddresses.ListIndex), vbInformation
    Else
        MsgBox "Error sending test e-mail." & vbCrLf & "Please check your SMTP settings.", vbCritical
    End If
End Sub

Private Sub Form_Load()
    Dim TempList As String
    Dim TempList2 As String
    Dim Counter As Integer
    
    Call MainMod.SetOnTopMode(Me)
    
    Me.Visible = False
    
    'Add supported programs to the various lists
    If Me.lstCameraControl.ListCount <> 4 Then
        Me.lstCameraControl.Clear
        Me.lstCameraControl.AddItem "CCDSoft v5"
        Me.lstCameraControl.AddItem "MaxIm DL v4/5/6"
        Me.lstCameraControl.AddItem "CCDSoft v5 (w/AO)"
        Me.lstCameraControl.AddItem "TheSkyX"
    End If
    
    If Me.lstMountControl.ListCount <> 3 Then
        Me.lstMountControl.Clear
        Me.lstMountControl.AddItem "TheSky6"
        Me.lstMountControl.AddItem "ASCOM Driver"
        Me.lstMountControl.AddItem "TheSkyX"
    End If
    
    If Me.lstPlanetarium.ListCount <> 3 Then
        Me.lstPlanetarium.Clear
        Me.lstPlanetarium.AddItem "None"
        Me.lstPlanetarium.AddItem "TheSky6"
        Me.lstPlanetarium.AddItem "TheSkyX"
    End If
    
    If Me.lstFocuserControl.ListCount <> 7 Then
        Me.lstFocuserControl.Clear
        Me.lstFocuserControl.AddItem "None"
        Me.lstFocuserControl.AddItem "FocusMax"
        Me.lstFocuserControl.AddItem "FocusMax AcquireStar"
        Me.lstFocuserControl.AddItem "CCDSoft @Focus"
        Me.lstFocuserControl.AddItem "CCDSoft @Focus2"
        Me.lstFocuserControl.AddItem "MaxIm/DL Focus"
        Me.lstFocuserControl.AddItem "TheSkyX @Focus2"
        Me.lstFocuserControl.AddItem "TheSkyX @Focus3"
    End If
    
    If Me.lstRotator.ListCount <> 6 Then
        Me.lstRotator.Clear
        Me.lstRotator.AddItem "None"
        Me.lstRotator.AddItem "Optec Pyxis"
        Me.lstRotator.AddItem "Astrodon TAKometer"
        Me.lstRotator.AddItem "RCOS PIR"
        Me.lstRotator.AddItem "ASCOM Driver"
        Me.lstRotator.AddItem "Manual"
    End If
    
    If Me.lstDomeControl.ListCount <> 5 Then
        Me.lstDomeControl.Clear
        Me.lstDomeControl.AddItem "None/Custom"
        Me.lstDomeControl.AddItem "AutomaDome"
        Me.lstDomeControl.AddItem "Digital Dome Works"
        Me.lstDomeControl.AddItem "ASCOM Driver"
        Me.lstDomeControl.AddItem "AutomaDome (TheSkyX)"
    End If
    
    If Me.lstCloudSensor.ListCount <> 4 Then
        Me.lstCloudSensor.Clear
        Me.lstCloudSensor.AddItem "None"
        Me.lstCloudSensor.AddItem "Boltwood/Clarity I"
        Me.lstCloudSensor.AddItem "Boltwood/Clarity II"
        Me.lstCloudSensor.AddItem "Boltwood/Clarity II Remote"
        Me.lstCloudSensor.AddItem "AAG Cloud Watcher"
        Me.lstCloudSensor.AddItem "AAG Cloud Watcher Remote"
    End If
    
    If Me.lstPlateSolve.ListCount <> 4 Then
        Me.lstPlateSolve.Clear
        Me.lstPlateSolve.AddItem "CCDSoft/TheSky"
        Me.lstPlateSolve.AddItem "MaxIm/PinPoint LE"
        Me.lstPlateSolve.AddItem "PinPoint Full"
        Me.lstPlateSolve.AddItem "TheSkyX"
    End If
    
    If Me.lstCatalog.ListCount <> 6 Then
        Me.lstCatalog.Clear
        Me.lstCatalog.AddItem "GSC 1.1"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 0
        Me.lstCatalog.AddItem "GSC 1.1 Corrected"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 3
        Me.lstCatalog.AddItem "Tycho 2"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 4
        Me.lstCatalog.AddItem "USNO A2.0"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 5
        Me.lstCatalog.AddItem "USNO SA2.0"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 1
        Me.lstCatalog.AddItem "USNO UCAC2"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 6
        Me.lstCatalog.AddItem "USNO UCAC3"
        Me.lstCatalog.ItemData(Me.lstCatalog.NewIndex) = 10
    End If
    
    Me.lstCameraControl.ListIndex = CInt(GetMySetting(RegistryName, "CameraControl", "0"))
    TempList = GetMySetting(RegistryName, "FilterList", "Red,Green,Blue,Clear,")
    TempList2 = GetMySetting(RegistryName, "FilterFocusOffsetList", "")
    Me.lstFilters.Clear
    Me.lstFiltersFocusOffset.Clear
    Do While Len(TempList) > 0
        Me.lstFilters.AddItem Left(TempList, InStr(TempList, ",") - 1)
        
        If Len(TempList2) = 0 Then
            Me.lstFiltersFocusOffset.AddItem PadFilterName(Left(TempList, InStr(TempList, ",") - 1)) & vbTab & "0"
            Me.lstFiltersFocusOffset.ItemData(Me.lstFiltersFocusOffset.NewIndex) = 0
        Else
            Me.lstFiltersFocusOffset.AddItem PadFilterName(Left(TempList, InStr(TempList, ",") - 1)) & vbTab & Left(TempList2, InStr(TempList2, ",") - 1)
            Me.lstFiltersFocusOffset.ItemData(Me.lstFiltersFocusOffset.NewIndex) = CInt(Left(TempList2, InStr(TempList2, ",") - 1))
            TempList2 = Mid(TempList2, InStr(TempList2, ",") + 1)
        End If
        
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    
    Me.chkEnableFilterOffsets.Value = CInt(GetMySetting(RegistryName, "FilterFocusOffsetEnabled", "0"))
    Me.txtMaximumStarFaded.Text = GetMySetting(RegistryName, "MaximumStarFadedErrors", "5")
    
    Me.chkIgnoreStarFaded.Value = CInt(GetMySetting(RegistryName, "IgnoreStarFaded", "0"))
    Me.chkIgnoreExposureAborted.Value = CInt(GetMySetting(RegistryName, "IgnoreExposureAborted", "0"))
    
    Me.chkInternalGuider.Value = CInt(GetMySetting(RegistryName, "InternalGuider", "0"))
    Me.chkDisableForceFilterChange.Value = CInt(GetMySetting(RegistryName, "DisableForceFilterChange", "0"))
    Me.lstMountControl.ListIndex = CInt(GetMySetting(RegistryName, "MountControl", "0"))
    Me.lstPlanetarium.ListIndex = CInt(GetMySetting(RegistryName, "Planetarium", "1"))
    Me.lstFocuserControl.ListIndex = CInt(GetMySetting(RegistryName, "FocusControl", "1"))
    If CInt(GetMySetting(RegistryName, "Rotator", "0")) < Me.lstRotator.ListCount Then
        Me.lstRotator.ListIndex = CInt(GetMySetting(RegistryName, "Rotator", "0"))
    Else
        Me.lstRotator.ListIndex = 0
    End If
    
    Me.chkDisconnectAtEnd.Value = vbUnchecked 'CInt(GetMySetting(RegistryName, "DisconnectAtEnd", "0"))
    
    Me.txtCOMNum.Text = GetMySetting(RegistryName, "RotatorCOMNumber", "1")
    Me.txtHomeRotationAngle.Text = GetMySetting(RegistryName, "RotatorHomeAngle", "0")
    Me.txtGuiderCalAngle.Text = GetMySetting(RegistryName, "GuiderCalAngle", "0")
    Me.chkGuiderRotates.Value = GetMySetting(RegistryName, "GuiderRotates", "1")
    Me.chkReverseRotatorDirection.Value = GetMySetting(RegistryName, "ReverseRotatorDirection", "1")
    Me.chkGuiderMirrorImage.Value = GetMySetting(RegistryName, "GuiderMirrorImage", "0")
    Me.chkRotateTheSkyFOVI.Value = GetMySetting(RegistryName, "RotateTheSkyFOVI", "0")
    Me.optRotatorFlip(CInt(GetMySetting(RegistryName, "RotatorAtFlipOptions", "0"))).Value = True
    
    Me.txtSaveTo.Text = GetMySetting(RegistryName, "SaveToPath", App.Path & "\Images\")
    Me.chkMaxImCompression.Value = CInt(GetMySetting(RegistryName, "MaxImCompression", "0"))
    
    Me.lstPlateSolve.ListIndex = CInt(GetMySetting(RegistryName, "PlateSolve", "0"))
    Me.txtNorthAngle.Text = GetMySetting(RegistryName, "NorthAngle", "0")
    Me.chkIgnoreNorthAngle.Value = CInt(GetMySetting(RegistryName, "IgnoreNorthAngle", "0"))
    Me.txtPixelScale.Text = GetMySetting(RegistryName, "PixelScale", Format(1, "0.00"))
    Me.lstCatalog.ListIndex = CInt(GetMySetting(RegistryName, "PinPointCatalog", "1"))
    Me.txtCatalogPath.Text = GetMySetting(RegistryName, "PinPointCatalogPath", "C:\")
    
    Me.txtCatalogMagMax.Text = GetMySetting(RegistryName, "PinPointCatalogMagMax", Format(-2, "0.0"))
    Me.txtCatalogMagMin.Text = GetMySetting(RegistryName, "PinPointCatalogMagMin", Format(20, "0.0"))
    Me.txtSearchArea.Text = GetMySetting(RegistryName, "PinPointSearchArea", Format(256, "0.00"))
    Me.txtMinStarBrightness.Text = CLng(GetMySetting(RegistryName, "PinPointMinStarBrightnes", "200"))
    Me.txtStandardDeviation.Text = GetMySetting(RegistryName, "PinPointStandardDeviation", Format(4, "0.0"))
    Me.txtMaxNumStars.Text = CInt(GetMySetting(RegistryName, "PinPointMaxNumStars", "0"))
        
    Me.chkDarkSubtractPlateSolveImage.Value = CInt(GetMySetting(RegistryName, "DarkSubtractPlateSolveImages", "1"))
    Me.txtPinPointLETimeOut.Text = GetMySetting(RegistryName, "PinPointLETimeOut", "5")
    Me.chkPinPointLERetry.Value = CInt(GetMySetting(RegistryName, "PinPointLERetry", "1"))
    
    If CInt(GetMySetting(RegistryName, "DomeControl", "0")) < Me.lstDomeControl.ListCount Then
        Me.lstDomeControl.ListIndex = GetMySetting(RegistryName, "DomeControl", "0")
    Else
        Me.lstDomeControl.ListIndex = 0
    End If
    
    If CInt(GetMySetting(RegistryName, "CloudSensor", "0")) < Me.lstCloudSensor.ListCount Then
        Me.lstCloudSensor.ListIndex = GetMySetting(RegistryName, "CloudSensor", "0")
    Else
        Me.lstCloudSensor.ListIndex = 0
    End If

    TempList = GetMySetting(RegistryName, "CloudSensorPauseActionList", "0,1,2,")
    For Counter = 0 To Me.lstCloudSensorPauseActionWhen.ListCount - 1
        Me.lstCloudSensorPauseActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        If CInt(Left(TempList, InStr(TempList, ",") - 1)) = 3 Then
            'Parameter from old version, this goes to a different list
            Me.lstRainSensorPauseActionWhen.Selected(0) = True
        Else
            Me.lstCloudSensorPauseActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        End If
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstCloudSensorPauseActionWhen.ListIndex = -1

    TempList = GetMySetting(RegistryName, "RainSensorPauseActionList", "")
    If TempList <> "" Then
        For Counter = 0 To Me.lstRainSensorPauseActionWhen.ListCount - 1
            Me.lstRainSensorPauseActionWhen.Selected(Counter) = False
        Next Counter
        Do While Len(TempList) > 0
            If CInt(Left(TempList, InStr(TempList, ",") - 1)) < Me.lstRainSensorPauseActionWhen.ListCount Then
                Me.lstRainSensorPauseActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
            End If
            TempList = Mid(TempList, InStr(TempList, ",") + 1)
        Loop
    End If
    Me.lstRainSensorPauseActionWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "WindSensorPauseActionList", "")
    For Counter = 0 To Me.lstWindSensorPauseActionWhen.ListCount - 1
        Me.lstWindSensorPauseActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstWindSensorPauseActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstWindSensorPauseActionWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "LightSensorPauseActionList", "")
    For Counter = 0 To Me.lstLightSensorPauseActionWhen.ListCount - 1
        Me.lstLightSensorPauseActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstLightSensorPauseActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstLightSensorPauseActionWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "CloudSensorCloseDomeWhenList", "2,")
    For Counter = 0 To Me.lstCloudSensorCloseDomeWhen.ListCount - 1
        Me.lstCloudSensorCloseDomeWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        If CInt(Left(TempList, InStr(TempList, ",") - 1)) = 3 Then
            'Parameter from old version, this goes to a different list
            Me.lstRainSensorCloseDomeWhen.Selected(0) = True
        Else
            Me.lstCloudSensorCloseDomeWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        End If
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstCloudSensorCloseDomeWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "RainSensorCloseDomeWhenList", "")
    If TempList <> "" Then
        For Counter = 0 To Me.lstRainSensorCloseDomeWhen.ListCount - 1
            Me.lstRainSensorCloseDomeWhen.Selected(Counter) = False
        Next Counter
        Do While Len(TempList) > 0
            If CInt(Left(TempList, InStr(TempList, ",") - 1)) < Me.lstRainSensorCloseDomeWhen.ListCount Then
                Me.lstRainSensorCloseDomeWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
            End If
            TempList = Mid(TempList, InStr(TempList, ",") + 1)
        Loop
    End If
    Me.lstRainSensorCloseDomeWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "WindSensorCloseDomeWhenList", "")
    For Counter = 0 To Me.lstWindSensorCloseDomeWhen.ListCount - 1
        Me.lstWindSensorCloseDomeWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstWindSensorCloseDomeWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstWindSensorCloseDomeWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "LightSensorCloseDomeWhenList", "")
    For Counter = 0 To Me.lstLightSensorCloseDomeWhen.ListCount - 1
        Me.lstLightSensorCloseDomeWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstLightSensorCloseDomeWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstLightSensorCloseDomeWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "CloudSensorResumeWhenList", "1,")
    For Counter = 0 To Me.lstCloudSensorResumeActionWhen.ListCount - 1
        Me.lstCloudSensorResumeActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstCloudSensorResumeActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstCloudSensorResumeActionWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "RainSensorResumeWhenList", "")
    If TempList = "" Then
        If Me.lstCloudSensor.ListIndex = 1 Then
            TempList = "0,"
        Else
            TempList = "1,"
        End If
    End If
    For Counter = 0 To Me.lstRainSensorResumeActionWhen.ListCount - 1
        Me.lstRainSensorResumeActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        If CInt(Left(TempList, InStr(TempList, ",") - 1)) < Me.lstRainSensorResumeActionWhen.ListCount Then
            Me.lstRainSensorResumeActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        End If
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstRainSensorResumeActionWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "WindSensorResumeWhenList", "")
    For Counter = 0 To Me.lstWindSensorResumeActionWhen.ListCount - 1
        Me.lstWindSensorResumeActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstWindSensorResumeActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstWindSensorResumeActionWhen.ListIndex = -1
    
    TempList = GetMySetting(RegistryName, "LightSensorResumeWhenList", "")
    For Counter = 0 To Me.lstLightSensorResumeActionWhen.ListCount - 1
        Me.lstLightSensorResumeActionWhen.Selected(Counter) = False
    Next Counter
    Do While Len(TempList) > 0
        Me.lstLightSensorResumeActionWhen.Selected(CInt(Left(TempList, InStr(TempList, ",") - 1))) = True
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    Me.lstLightSensorResumeActionWhen.ListIndex = -1
    
    Me.chkAutoDomeClose.Value = CInt(GetMySetting(RegistryName, "DomeClosesAutomaticallyOnBadWeather", CStr(vbChecked)))
    Me.txtQuerySensorPeriod.Text = GetMySetting(RegistryName, "QueryCloudSensorPeriod", "5")
    Me.txtClearTime.Text = GetMySetting(RegistryName, "CloudSensorClearTime", "30")
    Me.chkParkMountWhenCloudy.Value = GetMySetting(RegistryName, "ParkMountWhenCloudy", "1")
    Me.chkParkMountFirst.Value = GetMySetting(RegistryName, "ParkMountBeforeCloseDome", "0")
    
    Me.chkEnableWeatherMonitorScripts.Value = GetMySetting(RegistryName, "EnableWeatherMonitorScripts", "0")
    Me.txtAfterPauseScript.Text = GetMySetting(RegistryName, "WeatherMonitorAfterPauseScript", "")
    Me.txtAfterCloseScript.Text = GetMySetting(RegistryName, "WeatherMonitorAfterCloseScript", "")
    Me.txtAfterGoodScript.Text = GetMySetting(RegistryName, "WeatherMonitorAfterGoodScript", "")

    Me.chkUncoupleDomeDuringSlews.Value = GetMySetting(RegistryName, "UncoupleDomeDuringSlews", "0")
    Me.chkHaltOnDomeError.Value = GetMySetting(RegistryName, "HaltOnDomeError", CStr(vbChecked))
    Me.chkCloseDomeOnError.Value = GetMySetting(RegistryName, "CloseDomeOnError", "1")
    Me.txtDDWTimeout.Text = GetMySetting(RegistryName, "DDWTimeout", "5")
    Me.txtDDWRetryCount.Text = GetMySetting(RegistryName, "DDWRetryCount", "5")
    
    Call Camera.SetupFormsForCameraControlProgram

    If GetMySetting(MountRegistryName, "MountType", "GEM") = "GEM" Then
        Call optMountType_Click(0)
    Else
        Call optMountType_Click(1)
    End If
    
    Me.txtDelayAfterSlew.Text = GetMySetting(MountRegistryName, "DelayAfterSlew", "0")
    Me.chkOnlyDelayAfterMeridianFlip.Value = GetMySetting(MountRegistryName, "OnlyDelayAfterMeridianFlip", "0")
    
    Me.chkDisableDecComp.Value = GetMySetting(MountRegistryName, "DisableDeclinationCompensation", "0")
    Me.chkParkMount.Value = GetMySetting(MountRegistryName, "ParkMountOnError", "1")
    
    Me.txtEasternLimit.Text = GetMySetting(MountRegistryName, "GEMEasternLimit", "5")
    Me.txtWesternLimit.Text = GetMySetting(MountRegistryName, "GEMWesternLimit", "5")

    If GetMySetting(MountRegistryName, "GEMGuideCal", "West") = "West" Then
        Me.optGuiderCal(0).Value = True
    Else
        Me.optGuiderCal(1).Value = True
    End If
    
    Me.chkEnableScripts.Value = GetMySetting(MountRegistryName, "EnableSlewScripts", "0")
    Me.txtAfterScript.Text = GetMySetting(MountRegistryName, "AfterSlewScript", "")
    Me.txtBeforeScript.Text = GetMySetting(MountRegistryName, "BeforeSlewScript", "")
    
    Me.chkVerifyTeleCoords.Value = GetMySetting(MountRegistryName, "VerifyTeleCoords", "1")
    Me.txtMaxPointingError.Text = GetMySetting(MountRegistryName, "MaxPointingError", "3")
    
    Me.chkAutoDetermineMountSide.Value = GetMySetting(MountRegistryName, "AutoDetermineMountSide", "1")

    Me.chkEnable.Value = GetMySetting(GuiderRegistryName, "Enabled", "0")
    Me.chkDisableGuideStarRecovery.Value = GetMySetting(GuiderRegistryName, "DisableGuideStarRecovery", "0")
    Me.txtMinExp.Text = GetMySetting(GuiderRegistryName, "MinimumExposure", "1")
    Me.txtGuideExposureIncrement.Text = GetMySetting(GuiderRegistryName, "GuideExposureIncrement", "2")
    Me.txtMaxExp.Text = GetMySetting(GuiderRegistryName, "MaximumExposure", "30")
    Me.txtMinBright.Text = GetMySetting(GuiderRegistryName, "MinimumADU", "3000")
    Me.txtMaxBright.Text = GetMySetting(GuiderRegistryName, "MaximumADU", "20000")
    Me.txtGuideBoxX.Text = GetMySetting(GuiderRegistryName, "GuideBoxXSize", "32")
    Me.txtGuideBoxY.Text = GetMySetting(GuiderRegistryName, "GuideBoxYSize", "32")
    Me.txtGuideStarFWHM.Text = GetMySetting(GuiderRegistryName, "GuideStarFWHM", "4")
    Me.chkIgnore1PixelStars.Value = GetMySetting(GuiderRegistryName, "Ignore1PixelStars", "0")
    Me.txtMaxStarMovement.Text = GetMySetting(GuiderRegistryName, "MaximumStarMovement", "1")
    Me.chkContinuousAutoguide.Value = GetMySetting(GuiderRegistryName, "ContinuousAutoguiding", "0")
    Me.chkRestartGuidingWhenLargeError.Value = GetMySetting(GuiderRegistryName, "RestartGuidingWhenLargeError", "0")
    Me.txtRestartError.Text = GetMySetting(GuiderRegistryName, "RestartGuidingError", "1")
    Me.txtRestartCycles.Text = GetMySetting(GuiderRegistryName, "RestartGuidingCycles", "1")
    
    Me.txtSMTPServer.Text = GetMySetting(RegistryName, "SMTPServer", "")
    Me.txtSMTPPort.Text = GetMySetting(RegistryName, "SMTPPort", "25")
    Me.chkAuthentication.Value = CInt(GetMySetting(RegistryName, "UseSMTPAuthentication", "0"))
    Me.txtSMTPUsername.Text = GetMySetting(RegistryName, "SMTPUsername", "")
    Me.txtSMTPPassword.Text = GetMySetting(RegistryName, "SMTPPassword", "")
    Me.txtFromAddress.Text = GetMySetting(RegistryName, "FromAddress", "")
    TempList = GetMySetting(RegistryName, "ToAddresses", "")
    Me.lstToAddresses.Clear
    Do While Len(TempList) > 0
        Me.lstToAddresses.AddItem Left(TempList, InStr(TempList, ",") - 1)
        TempList = Mid(TempList, InStr(TempList, ",") + 1)
    Loop
    For Counter = 0 To Me.chkEMailAlert.Count - 1
        Me.chkEMailAlert(Counter).Value = CInt(GetMySetting(RegistryName, "EMailAlert" & Counter, "0"))
    Next Counter
    Me.txtEMailScript.Text = GetMySetting(RegistryName, "EMailScript", "")
    
    Me.chkMeasureAverageHFD.Value = CInt(GetMySetting(RegistryName, "FocusMaxMeasureAverageHFD", "1"))
    
    Me.chkRetryFocusRunOnFailure.Value = CInt(GetMySetting(RegistryName, "RetryFocusRunOnFailure", "0"))
    Me.txtFocusRetryCount.Text = GetMySetting(RegistryName, "FocusRetryCount", "1")
    Me.txtFocusTimeOut.Text = GetMySetting(RegistryName, "FocusTimeOut", "10")
    
    Me.txtWeatherMonitorRemoteFile.Text = GetMySetting(RegistryName, "WeatherMonitorRemoteFile", "")
    Me.chkWeatherMonitorRepeatAction.Value = CInt(GetMySetting(RegistryName, "WeatherMonitorRepeatAction", "0"))
    
    Me.SSTab1.Tab = 0
    
    Me.chkEnableWatchdog.Value = CInt(GetMySetting(RegistryName, "WatchdogEnable", "0"))
    
    Call ValidateAllEntries
End Sub

Private Sub lstCameraControl_Click()
    If InStr(Me.cmdGetFilters.Caption, Me.lstCameraControl.List(Me.lstCameraControl.ListIndex)) = 0 And Not frmMain.RunningAction Then
        frmTempGraph.Timer1.Enabled = False
        Call Camera.CameraUnload
    End If
    
    If Me.lstCameraControl.ListIndex = CameraControl.MaxIm Then
        Me.chkMaxImCompression.Enabled = True
        Me.chkIgnoreStarFaded.Enabled = True
        Me.chkIgnoreExposureAborted.Enabled = True
        Me.lblMaximumStarFaded.Enabled = True
        Me.txtMaximumStarFaded.Enabled = True
    Else
        Me.chkMaxImCompression.Enabled = False
        Me.chkMaxImCompression.Value = vbUnchecked
        Me.chkIgnoreStarFaded.Enabled = False
        Me.chkIgnoreStarFaded.Value = vbUnchecked
        Me.chkIgnoreExposureAborted.Enabled = False
        Me.chkIgnoreExposureAborted.Value = vbUnchecked
        Me.txtMaximumStarFaded.Enabled = False
        Me.lblMaximumStarFaded.Enabled = False
    End If

    Me.cmdGetFilters.Caption = "Get Filters from " & Left(Me.lstCameraControl.List(Me.lstCameraControl.ListIndex), 8)
End Sub

Private Sub lstCloudSensor_Click()
    Dim Counter As Integer
    If Me.lstCloudSensor.ListIndex = 0 Then
        Me.SSTab1.TabEnabled(3) = False
    Else
        Me.SSTab1.TabEnabled(3) = True
        
        If Me.lstCloudSensor.ListIndex = 1 Then
            For Counter = 0 To 2
                Me.lblWindConditions(Counter).Visible = False
                Me.lblLightConditions(Counter).Visible = False
            Next Counter
            
            Me.lstWindSensorCloseDomeWhen.Visible = False
            Me.lstWindSensorPauseActionWhen.Visible = False
            Me.lstWindSensorResumeActionWhen.Visible = False
            
            Me.lstLightSensorCloseDomeWhen.Visible = False
            Me.lstLightSensorPauseActionWhen.Visible = False
            Me.lstLightSensorResumeActionWhen.Visible = False
            
            Call lstRainSensorPauseActionWhen_ItemCheck(0)
            Call lstRainSensorCloseDomeWhen_ItemCheck(0)
            Call lstRainSensorResumeActionWhen_ItemCheck(0)
        Else
            For Counter = 0 To 2
                Me.lblWindConditions(Counter).Visible = True
                Me.lblLightConditions(Counter).Visible = True
            Next Counter
        
            Me.lstWindSensorCloseDomeWhen.Visible = True
            Me.lstWindSensorPauseActionWhen.Visible = True
            Me.lstWindSensorResumeActionWhen.Visible = True
            
            Me.lstLightSensorCloseDomeWhen.Visible = True
            Me.lstLightSensorPauseActionWhen.Visible = True
            Me.lstLightSensorResumeActionWhen.Visible = True
        End If
    
        If Me.chkParkMountFirst.Value = vbChecked Then
            'toggle value to make sure everything is setup properly
            Me.chkParkMountFirst.Value = vbUnchecked
            DoEvents
            Me.chkParkMountFirst.Value = vbChecked
        End If
    End If
    
    If Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityIIRemote Or _
        Me.lstCloudSensor.ListIndex = WeatherMonitorControl.AAGRemote Then
        Me.txtWeatherMonitorRemoteFile.Visible = True
        Me.cmdWeatherMonitorRemoteFileOpen.Visible = True
    Else
        Me.txtWeatherMonitorRemoteFile.Visible = False
        Me.cmdWeatherMonitorRemoteFileOpen.Visible = False
    End If
End Sub

Private Sub lstCloudSensorCloseDomeWhen_ItemCheck(Item As Integer)
    If Me.chkAutoDomeClose.Value = vbChecked Then
        If Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityII Or _
        Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityIIRemote Or _
        Me.lstCloudSensor.ListIndex = WeatherMonitorControl.AAGRemote Or _
            Me.lstCloudSensor.ListIndex = WeatherMonitorControl.AAG Then
            'Clarity I has a hard coded "Very Cloudy" level that may be worse than the software programmed level
            'So allow the Very Cloudy condition to be enabled for CCD Commander to close the dome
            Me.lstCloudSensorCloseDomeWhen.Selected(2) = False
        End If
    End If

    If Me.chkParkMountFirst.Value = vbChecked Then
        Me.lstCloudSensorPauseActionWhen.Selected(Item) = Me.lstCloudSensorPauseActionWhen.Selected(Item) Or Me.lstCloudSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstCloudSensorPauseActionWhen_ItemCheck(Item As Integer)
    If Me.chkParkMountFirst.Value = vbChecked And Me.lstCloudSensorCloseDomeWhen.Enabled Then
        Me.lstCloudSensorPauseActionWhen.Selected(Item) = Me.lstCloudSensorPauseActionWhen.Selected(Item) Or Me.lstCloudSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstDomeControl_Click()
    If Me.lstDomeControl.ListIndex = 0 Then
        Me.lstCloudSensorCloseDomeWhen.Enabled = False
        Me.lstRainSensorCloseDomeWhen.Enabled = False
        Me.lstWindSensorCloseDomeWhen.Enabled = False
        Me.lstLightSensorCloseDomeWhen.Enabled = False
        Me.SSTab1.TabEnabled(7) = False
        Me.cmdASCOMDomeConfigurre.Visible = False
    Else
        Me.lstCloudSensorCloseDomeWhen.Enabled = True
        Me.lstRainSensorCloseDomeWhen.Enabled = True
        Me.lstWindSensorCloseDomeWhen.Enabled = True
        Me.lstLightSensorCloseDomeWhen.Enabled = True
        
        Me.SSTab1.TabEnabled(7) = True
        
        If Me.lstDomeControl.List(Me.lstDomeControl.ListIndex) = "Digital Dome Works" Then
            Me.fraDDWRetry.Enabled = True
        Else
            Me.fraDDWRetry.Enabled = False
        End If
    
        If Left(Me.lstDomeControl.List(Me.lstDomeControl.ListIndex), 5) = "ASCOM" Then
            Me.cmdASCOMDomeConfigurre.Visible = True
        Else
            Me.cmdASCOMDomeConfigurre.Visible = False
        End If
        
        If Me.chkParkMountFirst.Value = vbChecked Then
            'toggle value to make sure everything is setup properly
            Me.chkParkMountFirst.Value = vbUnchecked
            DoEvents
            Me.chkParkMountFirst.Value = vbChecked
        End If
    End If
End Sub

Private Sub lstFilters_DblClick()
    Dim NewDescription As String
    
    'Changing filter name
    If Me.lstFilters.ListIndex = -1 Then
        MsgBox "You must select a filter before you can change it."
        Exit Sub
    End If

    NewDescription = InputBox("Enter the new description for filter #" & Me.lstFilters.ListIndex + 1, "Filter #" & Me.lstFilters.ListIndex + 1 & " Descirption", Me.lstFilters.List(Me.lstFilters.ListIndex))

    If NewDescription <> "" Then
        Me.lstFilters.List(Me.lstFilters.ListIndex) = NewDescription
        Me.lstFiltersFocusOffset.List(Me.lstFilters.ListIndex) = PadFilterName(NewDescription) & vbTab & Me.lstFiltersFocusOffset.ItemData(Me.lstFilters.ListIndex)
    End If
End Sub

Private Sub lstFiltersFocusOffset_DblClick()
    Dim NewOffset As String
    
    'Changing filter name
    If Me.lstFiltersFocusOffset.ListIndex = -1 Then
        MsgBox "You must select a filter before you can change the offset."
        Exit Sub
    End If

    NewOffset = InputBox("Enter the new offset for filter #" & Me.lstFiltersFocusOffset.ListIndex + 1, "Filter #" & Me.lstFiltersFocusOffset.ListIndex + 1 & " Focus Offset", Me.lstFiltersFocusOffset.ItemData(Me.lstFiltersFocusOffset.ListIndex))

    If NewOffset <> "" Then
        On Error Resume Next
        Me.lstFiltersFocusOffset.ItemData(Me.lstFiltersFocusOffset.ListIndex) = CInt(NewOffset)
        If Err.Number <> 0 Then
            On Error GoTo 0
        Else
            On Error GoTo 0
            Me.lstFiltersFocusOffset.List(Me.lstFiltersFocusOffset.ListIndex) = PadFilterName(Me.lstFilters.List(Me.lstFiltersFocusOffset.ListIndex)) & vbTab & Me.lstFiltersFocusOffset.ItemData(Me.lstFiltersFocusOffset.ListIndex)
        End If
    End If
End Sub

Private Sub lstFocuserControl_Click()
    If lstFocuserControl.ListIndex = FocusControl.None Then
        Me.SSTab1.TabEnabled(9) = False
    Else
        Me.SSTab1.TabEnabled(9) = True
    End If
    
    If lstFocuserControl.ListIndex = FocusControl.FocusMax Then
        Me.chkMeasureAverageHFD.Enabled = True
    Else
        Me.chkMeasureAverageHFD.Value = vbUnchecked
        Me.chkMeasureAverageHFD.Enabled = False
    End If
End Sub

Private Sub lstLightSensorCloseDomeWhen_ItemCheck(Item As Integer)
    If Me.chkAutoDomeClose.Value = vbChecked Then
        Me.lstLightSensorCloseDomeWhen.Selected(2) = False
    End If
    
    If Me.chkParkMountFirst.Value = vbChecked Then
        Me.lstLightSensorPauseActionWhen.Selected(Item) = Me.lstLightSensorPauseActionWhen.Selected(Item) Or Me.lstLightSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstLightSensorPauseActionWhen_ItemCheck(Item As Integer)
    If Me.chkParkMountFirst.Value = vbChecked And Me.lstLightSensorCloseDomeWhen.Enabled Then
        Me.lstLightSensorPauseActionWhen.Selected(Item) = Me.lstLightSensorPauseActionWhen.Selected(Item) Or Me.lstLightSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstMountControl_Click()
    If Left(Me.lstMountControl.List(Me.lstMountControl.ListIndex), 5) = "ASCOM" Then
        Me.cmdASCOMMountConfigurre.Visible = True
    Else
        Me.cmdASCOMMountConfigurre.Visible = False
    End If
    
    If Not frmMain.RunningAction Then
        Call Mount.MountUnload
    End If
End Sub

Private Sub lstPlanetarium_Click()
    If Me.lstPlanetarium.ListIndex = 0 Then
        Me.chkRotateTheSkyFOVI.Value = vbUnchecked
        Me.chkRotateTheSkyFOVI.Enabled = False
        Me.cmdGetAngleFromTheSky.Enabled = False
        Me.chkRotFromSky(0).Enabled = False
        Me.chkRotFromSky(1).Enabled = False
    Else
        Me.chkRotateTheSkyFOVI.Enabled = True
        Me.cmdGetAngleFromTheSky.Enabled = True
        Me.chkRotFromSky(0).Enabled = True
        Me.chkRotFromSky(1).Enabled = True
    End If
End Sub

Private Sub lstPlateSolve_Click()
    If Me.lstPlateSolve.ListIndex = 3 Then
        Me.fraPinPointSetup.Enabled = False
        Me.fraPinPointLESetup.Enabled = False
        Me.cmdGet.Visible = False
    ElseIf Me.lstPlateSolve.ListIndex = 2 Then
        Me.fraPinPointSetup.Enabled = True
        Me.fraPinPointLESetup.Enabled = True
        Me.cmdGet.Visible = False
    ElseIf Me.lstPlateSolve.ListIndex = 1 Then
        Me.fraPinPointSetup.Enabled = False
        Me.fraPinPointLESetup.Enabled = True
        Me.cmdGet.Visible = False
    Else
        Me.fraPinPointLESetup.Enabled = False
        Me.fraPinPointSetup.Enabled = False
        Me.cmdGet.Visible = True
    End If
End Sub

Private Sub lstRainSensorCloseDomeWhen_ItemCheck(Item As Integer)
    If Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityI Then
        Me.lstRainSensorCloseDomeWhen.Selected(0) = False
        Me.lstRainSensorCloseDomeWhen.Selected(2) = False
    End If
    
    If Me.chkAutoDomeClose.Value = vbChecked Then
        Me.lstRainSensorCloseDomeWhen.Selected(1) = False
        Me.lstRainSensorCloseDomeWhen.Selected(2) = False
    End If

    If Me.chkParkMountFirst.Value = vbChecked Then
        Me.lstRainSensorPauseActionWhen.Selected(Item) = Me.lstRainSensorPauseActionWhen.Selected(Item) Or Me.lstRainSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstRainSensorPauseActionWhen_ItemCheck(Item As Integer)
    If Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityI Then
        Me.lstRainSensorPauseActionWhen.Selected(0) = False
        Me.lstRainSensorPauseActionWhen.Selected(2) = False
    End If

    If Me.chkParkMountFirst.Value = vbChecked And Me.lstRainSensorCloseDomeWhen.Enabled Then
        Me.lstRainSensorPauseActionWhen.Selected(Item) = Me.lstRainSensorPauseActionWhen.Selected(Item) Or Me.lstRainSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstRainSensorResumeActionWhen_ItemCheck(Item As Integer)
    If Me.lstCloudSensor.ListIndex = WeatherMonitorControl.ClarityI Then
        Me.lstRainSensorResumeActionWhen.Selected(0) = False
    End If
End Sub

Private Sub lstRotator_Click()
    If lstRotator.ListIndex = 0 Then
        Me.SSTab1.TabEnabled(2) = False
        Me.txtNorthAngle.Enabled = True
        Me.cmdRotatorConfigure.Visible = False
    Else
        Me.SSTab1.TabEnabled(2) = True
        Me.txtNorthAngle.Enabled = False
        
        If lstRotator.ListIndex = 3 Or lstRotator.ListIndex = 4 Then
            Me.cmdRotatorConfigure.Visible = True
            Me.txtCOMNum.Enabled = False
        Else
            Me.cmdRotatorConfigure.Visible = False
            Me.txtCOMNum.Enabled = True
        End If
    End If
End Sub

Private Sub lstWindSensorCloseDomeWhen_ItemCheck(Item As Integer)
    If Me.chkAutoDomeClose.Value = vbChecked Then
        Me.lstWindSensorCloseDomeWhen.Selected(2) = False
    End If
    
    If Me.chkParkMountFirst.Value = vbChecked Then
        Me.lstWindSensorPauseActionWhen.Selected(Item) = Me.lstWindSensorPauseActionWhen.Selected(Item) Or Me.lstWindSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub lstWindSensorPauseActionWhen_ItemCheck(Item As Integer)
    If Me.chkParkMountFirst.Value = vbChecked And Me.lstWindSensorCloseDomeWhen.Enabled Then
        Me.lstWindSensorPauseActionWhen.Selected(Item) = Me.lstWindSensorPauseActionWhen.Selected(Item) Or Me.lstWindSensorCloseDomeWhen.Selected(Item)
    End If
End Sub

Private Sub mnuLoadSettings_Click()
    Dim MySettings As Variant
    Dim FileNo As Integer
    Dim Counter As Integer
    
    On Error Resume Next
    Call MkDir(App.Path & "\SavedSettings")
    On Error GoTo 0
    Me.ComDlg.DialogTitle = "Load Settings"
    Me.ComDlg.Filter = "CCDCommander Settings (*.bin)|*.bin"
    Me.ComDlg.FilterIndex = 1
    Me.ComDlg.InitDir = App.Path & "\SavedSettings\"
    Me.ComDlg.FileName = ""
    Me.ComDlg.CancelError = True
    Me.ComDlg.flags = cdlOFNOverwritePrompt
    On Error Resume Next
    Me.ComDlg.ShowOpen
    If Err.Number = 0 Then
        On Error GoTo 0
        
        FileNo = FreeFile
        Open Me.ComDlg.FileName For Binary Access Read As #FileNo
        
        Get #FileNo, , MySettings
        For Counter = 0 To UBound(MySettings, 1)
            Call MainMod.SaveMySetting(RegistryName, CStr(MySettings(Counter, 0)), CStr(MySettings(Counter, 1)))
        Next Counter
        Get #FileNo, , MySettings
        For Counter = 0 To UBound(MySettings, 1)
            Call MainMod.SaveMySetting(MountRegistryName, CStr(MySettings(Counter, 0)), CStr(MySettings(Counter, 1)))
        Next Counter
        Get #FileNo, , MySettings
        For Counter = 0 To UBound(MySettings, 1)
            Call MainMod.SaveMySetting(GuiderRegistryName, CStr(MySettings(Counter, 0)), CStr(MySettings(Counter, 1)))
        Next Counter
        
        Close #FileNo
        
        'Reload the settings from the registry
        Call Form_Load
        
        'Form_Load hides the window (since load is called normally at initial startup), need to make it visible again.
        Me.Visible = True
    Else
        On Error GoTo 0
    End If
    

End Sub

Private Sub mnuSaveSettings_Click()
    Dim MySettings As Variant
    Dim FileNo As Integer
    
    Call CheckAndSaveSettings
    
    On Error Resume Next
    Call MkDir(App.Path & "\SavedSettings")
    On Error GoTo 0
    Me.ComDlg.DialogTitle = "Save Settings"
    Me.ComDlg.Filter = "CCDCommander Settings (*.bin)|*.bin"
    Me.ComDlg.FilterIndex = 1
    Me.ComDlg.InitDir = App.Path & "\SavedSettings\"
    Me.ComDlg.FileName = Format(Now, "yymmdd") & ".bin"
    Me.ComDlg.CancelError = True
    Me.ComDlg.flags = cdlOFNOverwritePrompt
    On Error Resume Next
    Me.ComDlg.ShowSave
    If Err.Number = 0 Then
        On Error GoTo 0
        
        FileNo = FreeFile
        Open Me.ComDlg.FileName For Binary Access Write As #FileNo
        
        MySettings = GetAllSettings("CCDCommander", RegistryName)
        Put #FileNo, , MySettings
        MySettings = GetAllSettings("CCDCommander", MountRegistryName)
        Put #FileNo, , MySettings
        MySettings = GetAllSettings("CCDCommander", GuiderRegistryName)
        Put #FileNo, , MySettings
        
        Close #FileNo
    Else
        On Error GoTo 0
    End If
    
End Sub

Private Sub OKButton_Click()
    Call CheckAndSaveSettings
    
    Me.Hide
End Sub

Private Sub CheckAndSaveSettings()
    Dim TempList As String
    Dim Counter As Integer
    
    If Not ValidateAllEntries() Then
        Beep
        Exit Sub
    End If
    
    If InStr(Me.lstCameraControl.List(Me.lstCameraControl.ListIndex), "CCDSoft") = 0 And _
        InStr(Me.lstFocuserControl.List(Me.lstFocuserControl.ListIndex), "CCDSoft") > 0 Then
        Beep
        Me.SSTab1.Tab = 0
        Exit Sub
    End If
    
    Call SaveMySetting(GuiderRegistryName, "Enabled", Me.chkEnable.Value)
    Call SaveMySetting(GuiderRegistryName, "DisableGuideStarRecovery", Me.chkDisableGuideStarRecovery.Value)
    Call SaveMySetting(GuiderRegistryName, "MinimumExposure", Me.txtMinExp.Text)
    Call SaveMySetting(GuiderRegistryName, "GuideExposureIncrement", Me.txtGuideExposureIncrement.Text)
    Call SaveMySetting(GuiderRegistryName, "MaximumExposure", Me.txtMaxExp.Text)
    Call SaveMySetting(GuiderRegistryName, "MinimumADU", Me.txtMinBright.Text)
    Call SaveMySetting(GuiderRegistryName, "MaximumADU", Me.txtMaxBright.Text)
    Call SaveMySetting(GuiderRegistryName, "GuideBoxXSize", Me.txtGuideBoxX.Text)
    Call SaveMySetting(GuiderRegistryName, "GuideBoxYSize", Me.txtGuideBoxY.Text)
    Call SaveMySetting(GuiderRegistryName, "GuideStarFWHM", Me.txtGuideStarFWHM.Text)
    Call SaveMySetting(GuiderRegistryName, "Ignore1PixelStars", Me.chkIgnore1PixelStars.Value)
    Call SaveMySetting(GuiderRegistryName, "MaximumStarMovement", Me.txtMaxStarMovement.Text)
    Call SaveMySetting(GuiderRegistryName, "ContinuousAutoguiding", Me.chkContinuousAutoguide.Value)
    Call SaveMySetting(GuiderRegistryName, "RestartGuidingWhenLargeError", Me.chkRestartGuidingWhenLargeError.Value)
    Call SaveMySetting(GuiderRegistryName, "RestartGuidingError", Me.txtRestartError.Text)
    Call SaveMySetting(GuiderRegistryName, "RestartGuidingCycles", Me.txtRestartCycles.Text)
    
    If Me.optMountType(0).Value Then
        Call SaveMySetting(MountRegistryName, "MountType", "GEM")
    Else
        Call SaveMySetting(MountRegistryName, "MountType", "Fork")
    End If
    
    Call SaveMySetting(MountRegistryName, "DisableDeclinationCompensation", Me.chkDisableDecComp.Value)
    Call SaveMySetting(MountRegistryName, "ParkMountOnError", Me.chkParkMount.Value)
    
    Call SaveMySetting(MountRegistryName, "DelayAfterSlew", Me.txtDelayAfterSlew.Text)
    Call SaveMySetting(MountRegistryName, "OnlyDelayAfterMeridianFlip", Me.chkOnlyDelayAfterMeridianFlip.Value)
    
    Call SaveMySetting(MountRegistryName, "GEMEasternLimit", Me.txtEasternLimit.Text)
    Call SaveMySetting(MountRegistryName, "GEMWesternLimit", Me.txtWesternLimit.Text)

    If Me.optGuiderCal(0).Value Then
        Call SaveMySetting(MountRegistryName, "GEMGuideCal", "West")
    Else
        Call SaveMySetting(MountRegistryName, "GEMGuideCal", "East")
    End If

    Call SaveMySetting(MountRegistryName, "EnableSlewScripts", Me.chkEnableScripts.Value)
    Call SaveMySetting(MountRegistryName, "AfterSlewScript", Me.txtAfterScript.Text)
    Call SaveMySetting(MountRegistryName, "BeforeSlewScript", Me.txtBeforeScript.Text)
    
    Call SaveMySetting(MountRegistryName, "AutoDetermineMountSide", Me.chkAutoDetermineMountSide.Value)
    
    Call SaveMySetting(MountRegistryName, "VerifyTeleCoords", Me.chkVerifyTeleCoords.Value)
    Call SaveMySetting(MountRegistryName, "MaxPointingError", Me.txtMaxPointingError.Text)
    
    Call SaveMySetting(RegistryName, "CameraControl", Me.lstCameraControl.ListIndex)
    TempList = ""
    For Counter = 0 To Me.lstFilters.ListCount - 1
        TempList = TempList & Me.lstFilters.List(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "FilterList", TempList)
    Call SaveMySetting(RegistryName, "IgnoreStarFaded", Me.chkIgnoreStarFaded.Value)
    Call SaveMySetting(RegistryName, "MaximumStarFadedErrors", Me.txtMaximumStarFaded.Text)
    
    Call SaveMySetting(RegistryName, "IgnoreExposureAborted", Me.chkIgnoreExposureAborted.Value)
    
    Call SaveMySetting(RegistryName, "InternalGuider", Me.chkInternalGuider.Value)
    Call SaveMySetting(RegistryName, "DisableForceFilterChange", Me.chkDisableForceFilterChange.Value)
    Call SaveMySetting(RegistryName, "MountControl", Me.lstMountControl.ListIndex)
    Call SaveMySetting(RegistryName, "Planetarium", Me.lstPlanetarium.ListIndex)
    Call SaveMySetting(RegistryName, "FocusControl", Me.lstFocuserControl.ListIndex)
    Call SaveMySetting(RegistryName, "Rotator", Me.lstRotator.ListIndex)
    
    Call SaveMySetting(RegistryName, "DisconnectAtEnd", Me.chkDisconnectAtEnd.Value)
    
    Call SaveMySetting(RegistryName, "RotatorCOMNumber", Me.txtCOMNum.Text)
    Call SaveMySetting(RegistryName, "RotatorHomeAngle", Me.txtHomeRotationAngle.Text)
    Call SaveMySetting(RegistryName, "GuiderCalAngle", Me.txtGuiderCalAngle.Text)
    Call SaveMySetting(RegistryName, "GuiderRotates", Me.chkGuiderRotates.Value)
    Call SaveMySetting(RegistryName, "ReverseRotatorDirection", Me.chkReverseRotatorDirection.Value)
    Call SaveMySetting(RegistryName, "GuiderMirrorImage", Me.chkGuiderMirrorImage.Value)
    Call SaveMySetting(RegistryName, "RotateTheSkyFOVI", Me.chkRotateTheSkyFOVI.Value)
    
    If Me.optRotatorFlip(0).Value Then
        Call SaveMySetting(RegistryName, "RotatorAtFlipOptions", "0")
    Else
        Call SaveMySetting(RegistryName, "RotatorAtFlipOptions", "1")
    End If
    
    Call SaveMySetting(RegistryName, "SaveToPath", Me.txtSaveTo.Text)
    Call SaveMySetting(RegistryName, "MaxImCompression", Me.chkMaxImCompression.Value)
    
    Call SaveMySetting(RegistryName, "PlateSolve", Me.lstPlateSolve.ListIndex)
    Call SaveMySetting(RegistryName, "NorthAngle", Me.txtNorthAngle.Text)
    Call SaveMySetting(RegistryName, "IgnoreNorthAngle", Me.chkIgnoreNorthAngle.Value)
    Call SaveMySetting(RegistryName, "PixelScale", Me.txtPixelScale.Text)
    Call SaveMySetting(RegistryName, "PinPointCatalog", Me.lstCatalog.ListIndex)
    Call SaveMySetting(RegistryName, "PinPointCatalogPath", Me.txtCatalogPath.Text)
    
    Call SaveMySetting(RegistryName, "PinPointCatalogMagMax", Me.txtCatalogMagMax.Text)
    Call SaveMySetting(RegistryName, "PinPointCatalogMagMin", Me.txtCatalogMagMin.Text)
    Call SaveMySetting(RegistryName, "PinPointSearchArea", Me.txtSearchArea.Text)
    Call SaveMySetting(RegistryName, "PinPointMinStarBrightnes", Me.txtMinStarBrightness.Text)
    Call SaveMySetting(RegistryName, "PinPointStandardDeviation", Me.txtStandardDeviation.Text)
    Call SaveMySetting(RegistryName, "PinPointMaxNumStars", Me.txtMaxNumStars.Text)
    
    Call SaveMySetting(RegistryName, "DarkSubtractPlateSolveImages", Me.chkDarkSubtractPlateSolveImage.Value)
    Call SaveMySetting(RegistryName, "PinPointLETimeOut", Me.txtPinPointLETimeOut.Text)
    Call SaveMySetting(RegistryName, "PinPointLERetry", Me.chkPinPointLERetry.Value)
    
    Call SaveMySetting(RegistryName, "DomeControl", Me.lstDomeControl.ListIndex)
    Call SaveMySetting(RegistryName, "CloudSensor", Me.lstCloudSensor.ListIndex)
    
    TempList = ""
    For Counter = 0 To Me.lstCloudSensorPauseActionWhen.ListCount - 1
        If Me.lstCloudSensorPauseActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "CloudSensorPauseActionList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstRainSensorPauseActionWhen.ListCount - 1
        If Me.lstRainSensorPauseActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "RainSensorPauseActionList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstWindSensorPauseActionWhen.ListCount - 1
        If Me.lstWindSensorPauseActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "WindSensorPauseActionList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstLightSensorPauseActionWhen.ListCount - 1
        If Me.lstLightSensorPauseActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "LightSensorPauseActionList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstCloudSensorCloseDomeWhen.ListCount - 1
        If Me.lstCloudSensorCloseDomeWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "CloudSensorCloseDomeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstRainSensorCloseDomeWhen.ListCount - 1
        If Me.lstRainSensorCloseDomeWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "RainSensorCloseDomeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstWindSensorCloseDomeWhen.ListCount - 1
        If Me.lstWindSensorCloseDomeWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "WindSensorCloseDomeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstLightSensorCloseDomeWhen.ListCount - 1
        If Me.lstLightSensorCloseDomeWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "LightSensorCloseDomeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstCloudSensorResumeActionWhen.ListCount - 1
        If Me.lstCloudSensorResumeActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "CloudSensorResumeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstRainSensorResumeActionWhen.ListCount - 1
        If Me.lstRainSensorResumeActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "RainSensorResumeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstWindSensorResumeActionWhen.ListCount - 1
        If Me.lstWindSensorResumeActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "WindSensorResumeWhenList", TempList)
    
    TempList = ""
    For Counter = 0 To Me.lstLightSensorResumeActionWhen.ListCount - 1
        If Me.lstLightSensorResumeActionWhen.Selected(Counter) Then
            TempList = TempList & Counter & ","
        End If
    Next Counter
    Call SaveMySetting(RegistryName, "LightSensorResumeWhenList", TempList)
    
    Call SaveMySetting(RegistryName, "DomeClosesAutomaticallyOnBadWeather", Me.chkAutoDomeClose.Value)
    
    Call SaveMySetting(RegistryName, "QueryCloudSensorPeriod", Me.txtQuerySensorPeriod.Text)
    Call SaveMySetting(RegistryName, "CloudSensorClearTime", Me.txtClearTime.Text)
    Call SaveMySetting(RegistryName, "ParkMountWhenCloudy", Me.chkParkMountWhenCloudy.Value)
    Call SaveMySetting(RegistryName, "ParkMountBeforeCloseDome", Me.chkParkMountFirst.Value)
    
    Call SaveMySetting(RegistryName, "EnableWeatherMonitorScripts", Me.chkEnableWeatherMonitorScripts.Value)
    Call SaveMySetting(RegistryName, "WeatherMonitorAfterPauseScript", Me.txtAfterPauseScript.Text)
    Call SaveMySetting(RegistryName, "WeatherMonitorAfterCloseScript", Me.txtAfterCloseScript.Text)
    Call SaveMySetting(RegistryName, "WeatherMonitorAfterGoodScript", Me.txtAfterGoodScript.Text)
    
    Call SaveMySetting(RegistryName, "UncoupleDomeDuringSlews", Me.chkUncoupleDomeDuringSlews.Value)
    Call SaveMySetting(RegistryName, "HaltOnDomeError", Me.chkHaltOnDomeError.Value)
    Call SaveMySetting(RegistryName, "CloseDomeOnError", Me.chkCloseDomeOnError.Value)
    Call SaveMySetting(RegistryName, "DDWTimeout", Me.txtDDWTimeout.Text)
    Call SaveMySetting(RegistryName, "DDWRetryCount", Me.txtDDWRetryCount.Text)
    
    Call SaveMySetting(RegistryName, "SMTPServer", Me.txtSMTPServer.Text)
    Call SaveMySetting(RegistryName, "SMTPPort", Me.txtSMTPPort.Text)
    Call SaveMySetting(RegistryName, "UseSMTPAuthentication", Me.chkAuthentication.Value)
    Call SaveMySetting(RegistryName, "SMTPUsername", Me.txtSMTPUsername.Text)
    Call SaveMySetting(RegistryName, "SMTPPassword", Me.txtSMTPPassword.Text)
    Call SaveMySetting(RegistryName, "FromAddress", Me.txtFromAddress.Text)
    TempList = ""
    For Counter = 0 To Me.lstToAddresses.ListCount - 1
        TempList = TempList & Me.lstToAddresses.List(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "ToAddresses", TempList)
    For Counter = 0 To Me.chkEMailAlert.Count - 1
        Call SaveMySetting(RegistryName, "EMailAlert" & Counter, Me.chkEMailAlert(Counter).Value)
    Next Counter
    Call SaveMySetting(RegistryName, "EMailScript", Me.txtEMailScript.Text)
        
    TempList = ""
    For Counter = 0 To Me.lstFiltersFocusOffset.ListCount - 1
        TempList = TempList & Me.lstFiltersFocusOffset.ItemData(Counter) & ","
    Next Counter
    Call SaveMySetting(RegistryName, "FilterFocusOffsetList", TempList)

    Call SaveMySetting(RegistryName, "FilterFocusOffsetEnabled", Me.chkEnableFilterOffsets.Value)
    
    Call SaveMySetting(RegistryName, "FocusMaxMeasureAverageHFD", Me.chkMeasureAverageHFD.Value)
    
    Call SaveMySetting(RegistryName, "RetryFocusRunOnFailure", Me.chkRetryFocusRunOnFailure.Value)
    Call SaveMySetting(RegistryName, "FocusRetryCount", Me.txtFocusRetryCount.Text)
    Call SaveMySetting(RegistryName, "FocusTimeOut", Me.txtFocusTimeOut.Text)
    
    Call SaveMySetting(RegistryName, "WeatherMonitorRemoteFile", Me.txtWeatherMonitorRemoteFile.Text)
    Call SaveMySetting(RegistryName, "WeatherMonitorRepeatAction", Me.chkWeatherMonitorRepeatAction.Value)
    
    Call SaveMySetting(RegistryName, "WatchdogEnable", Me.chkEnableWatchdog.Value)
    
    Call Camera.PutFilterDataIntoForms
    
    Call Camera.SetupFormsForCameraControlProgram
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.SSTab1.HelpContextID = 200 + Me.SSTab1.Tab
    Me.HelpContextID = 200 + Me.SSTab1.Tab
End Sub

Private Sub txtAfterPauseScript_GotFocus()
    Me.txtAfterPauseScript.SelStart = 0
    Me.txtAfterPauseScript.SelLength = Len(Me.txtAfterPauseScript.Text)
End Sub

Private Sub txtAfterPauseScript_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKeyDelete
            txtAfterPauseScript.Text = ""
    End Select
End Sub

Private Sub txtAfterCloseScript_GotFocus()
    Me.txtAfterCloseScript.SelStart = 0
    Me.txtAfterCloseScript.SelLength = Len(Me.txtAfterCloseScript.Text)
End Sub

Private Sub txtAfterCloseScript_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKeyDelete
            txtAfterCloseScript.Text = ""
    End Select
End Sub

Private Sub txtAfterGoodScript_GotFocus()
    Me.txtAfterGoodScript.SelStart = 0
    Me.txtAfterGoodScript.SelLength = Len(Me.txtAfterGoodScript.Text)
End Sub

Private Sub txtAfterGoodScript_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack, vbKeyDelete
            txtAfterGoodScript.Text = ""
    End Select
End Sub

Private Sub txtCatalogMagMax_GotFocus()
    Me.txtCatalogMagMax.SelStart = 0
    Me.txtCatalogMagMax.SelLength = Len(Me.txtCatalogMagMax.Text)
End Sub

Private Sub txtCatalogMagMax_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtCatalogMagMax.Text)
    If Err.Number <> 0 Or Test <> Me.txtCatalogMagMax.Text Then
        Beep
        Cancel = True
    Else
        Settings.CatalogMagMax = Test
        
        Me.txtCatalogMagMax.Text = Format(Test, "0.0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtCatalogMagMin_GotFocus()
    Me.txtCatalogMagMin.SelStart = 0
    Me.txtCatalogMagMin.SelLength = Len(Me.txtCatalogMagMin.Text)
End Sub

Private Sub txtCatalogMagMin_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtCatalogMagMin.Text)
    If Err.Number <> 0 Or Test <> Me.txtCatalogMagMin.Text Then
        Beep
        Cancel = True
    Else
        Settings.CatalogMagMin = Test
        
        Me.txtCatalogMagMin.Text = Format(Test, "0.0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtDDWRetryCount_GotFocus()
    Me.txtDDWRetryCount.SelStart = 0
    Me.txtDDWRetryCount.SelLength = Len(Me.txtDDWRetryCount.Text)
End Sub

Private Sub txtDDWRetryCount_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDDWRetryCount.Text)
    If Err.Number <> 0 Or Test <> Me.txtDDWRetryCount.Text Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.DDWRetryCount = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtDDWTimeout_GotFocus()
    Me.txtDDWTimeout.SelStart = 0
    Me.txtDDWTimeout.SelLength = Len(Me.txtDDWTimeout.Text)
End Sub

Private Sub txtDDWTimeout_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtDDWTimeout.Text)
    If Err.Number <> 0 Or Test <> Me.txtDDWTimeout.Text Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.DDWTimeout = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtFocusRetryCount_GotFocus()
    Me.txtFocusRetryCount.SelStart = 0
    Me.txtFocusRetryCount.SelLength = Len(Me.txtFocusRetryCount.Text)
End Sub

Private Sub txtFocusRetryCount_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtFocusRetryCount.Text)
    If Err.Number <> 0 Or Test <> Me.txtFocusRetryCount.Text Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.RetryFocusCount = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtFocusTimeOut_GotFocus()
    Me.txtFocusTimeOut.SelStart = 0
    Me.txtFocusTimeOut.SelLength = Len(Me.txtFocusTimeOut.Text)
End Sub

Private Sub txtFocusTimeOut_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtFocusTimeOut.Text)
    If Err.Number <> 0 Or Test <> Me.txtFocusTimeOut.Text Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.FocusTimeOut = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtRestartCycles_GotFocus()
    Me.txtRestartCycles.SelStart = 0
    Me.txtRestartCycles.SelLength = Len(Me.txtRestartCycles.Text)
End Sub

Private Sub txtRestartCycles_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtRestartCycles.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtRestartCycles.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuiderRestartCycles = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtRestartError_GotFocus()
    Me.txtRestartError.SelStart = 0
    Me.txtRestartError.SelLength = Len(Me.txtRestartError.Text)
End Sub

Private Sub txtRestartError_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtRestartError.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtRestartError.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuiderRestartError = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtSearchArea_GotFocus()
    Me.txtSearchArea.SelStart = 0
    Me.txtSearchArea.SelLength = Len(Me.txtSearchArea.Text)
End Sub

Private Sub txtSearchArea_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtSearchArea.Text)
    If Err.Number <> 0 Or Test <> Me.txtSearchArea.Text Then
        Beep
        Cancel = True
    Else
        If Test < 100 Then
            Test = 100
        ElseIf Test > 676 Then
            Test = 676
        End If
                
        Settings.SearchArea = Test
        
        Me.txtSearchArea.Text = Format(Test, "0.0")
    End If
    On Error GoTo 0
End Sub

Private Sub txtSMTPServer_GotFocus()
    Me.txtSMTPServer.SelStart = 0
    Me.txtSMTPServer.SelLength = Len(Me.txtSMTPServer.Text)
End Sub

Private Sub txtSMTPPort_GotFocus()
    Me.txtSMTPPort.SelStart = 0
    Me.txtSMTPPort.SelLength = Len(Me.txtSMTPPort.Text)
End Sub

Private Sub txtSMTPPort_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtSMTPPort.Text)
    If Err.Number <> 0 Or Test <> Me.txtSMTPPort.Text Or Test < 1 Then
        Beep
        Cancel = True
    Else
        Settings.SMTPPort = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtSMTPUsername_GotFocus()
    Me.txtSMTPUsername.SelStart = 0
    Me.txtSMTPUsername.SelLength = Len(Me.txtSMTPUsername.Text)
End Sub

Private Sub txtSMTPPassword_GotFocus()
    Me.txtSMTPPassword.SelStart = 0
    Me.txtSMTPPassword.SelLength = Len(Me.txtSMTPPassword.Text)
End Sub

Private Sub txtFromAddress_GotFocus()
    Me.txtFromAddress.SelStart = 0
    Me.txtFromAddress.SelLength = Len(Me.txtFromAddress.Text)
End Sub

Private Sub txtStandardDeviation_GotFocus()
    Me.txtStandardDeviation.SelStart = 0
    Me.txtStandardDeviation.SelLength = Len(Me.txtStandardDeviation.Text)
End Sub

Private Sub txtStandardDeviation_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtStandardDeviation.Text)
    If Err.Number <> 0 Or Test <> Me.txtStandardDeviation.Text Then
        Beep
        Cancel = True
    Else
        If Test < 1 Then
            Test = 1
        ElseIf Test > 8 Then
            Test = 8
        End If
        
        Settings.StandardDeviation = Test
        
        Me.txtStandardDeviation.Text = Format(Test, "0.00")
    End If
    On Error GoTo 0
End Sub

Private Sub txtMinStarBrightness_GotFocus()
    Me.txtMinStarBrightness.SelStart = 0
    Me.txtMinStarBrightness.SelLength = Len(Me.txtMinStarBrightness.Text)
End Sub

Private Sub txtMinStarBrightness_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CDbl(Me.txtMinStarBrightness.Text)
    If Err.Number <> 0 Or Test <> Me.txtMinStarBrightness.Text Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.MinStarBrightness = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxNumStars_GotFocus()
    Me.txtMaxNumStars.SelStart = 0
    Me.txtMaxNumStars.SelLength = Len(Me.txtMaxNumStars.Text)
End Sub

Private Sub txtMaxNumStars_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CDbl(Me.txtMaxNumStars.Text)
    If Err.Number <> 0 Or Test <> Me.txtMaxNumStars.Text Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.MaxNumStars = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtCOMNum_GotFocus()
    Me.txtCOMNum.SelStart = 0
    Me.txtCOMNum.SelLength = Len(Me.txtCOMNum.Text)
End Sub

Private Sub txtCOMNum_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtCOMNum.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtCOMNum.Text Or Test > 255 Then
        Beep
        Cancel = True
    Else
        Settings.RotatorCOMNumber = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuiderCalAngle_GotFocus()
    Me.txtGuiderCalAngle.SelStart = 0
    Me.txtGuiderCalAngle.SelLength = Len(Me.txtGuiderCalAngle.Text)
End Sub

Private Sub txtGuiderCalAngle_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtGuiderCalAngle.Text)
    If Err.Number <> 0 Or Test <> Me.txtGuiderCalAngle.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuiderCalibrationAngle = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtHomeRotationAngle_GotFocus()
    Me.txtHomeRotationAngle.SelStart = 0
    Me.txtHomeRotationAngle.SelLength = Len(Me.txtHomeRotationAngle.Text)
End Sub

Private Sub txtHomeRotationAngle_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtHomeRotationAngle.Text)
    If Err.Number <> 0 Or Test <> Me.txtHomeRotationAngle.Text Then
        Beep
        Cancel = True
    Else
        Settings.HomeRotationAngle = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxPointingError_GotFocus()
    Me.txtMaxPointingError.SelStart = 0
    Me.txtMaxPointingError.SelLength = Len(Me.txtMaxPointingError.Text)
End Sub

Private Sub txtMaxPointingError_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMaxPointingError.Text)
    If Err.Number <> 0 Or Test <> Me.txtMaxPointingError.Text Then
        Beep
        Cancel = True
    Else
        Settings.MaxPointingError = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtNorthAngle_GotFocus()
    Me.txtNorthAngle.SelStart = 0
    Me.txtNorthAngle.SelLength = Len(Me.txtNorthAngle.Text)
End Sub

Private Sub txtNorthAngle_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtNorthAngle.Text)
    If Err.Number <> 0 Or Test <> Me.txtNorthAngle.Text Then
        Beep
        Cancel = True
    Else
        Settings.NorthAngle = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtPinPointLETimeOut_GotFocus()
    Me.txtPinPointLETimeOut.SelStart = 0
    Me.txtPinPointLETimeOut.SelLength = Len(Me.txtPinPointLETimeOut.Text)
End Sub

Private Sub txtPinPointLETimeOut_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtPinPointLETimeOut.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtPinPointLETimeOut.Text Then
        Beep
        Cancel = True
    Else
        Settings.PinPointLETimeout = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtPixelScale_GotFocus()
    Me.txtPixelScale.SelStart = 0
    Me.txtPixelScale.SelLength = Len(Me.txtPixelScale.Text)
End Sub

Private Sub txtPixelScale_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtPixelScale.Text)
    If Err.Number <> 0 Or Test < 0 Or Test <> Me.txtPixelScale.Text Then
        Beep
        Cancel = True
    Else
        Settings.PixelScale = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtQuerySensorPeriod_GotFocus()
    Me.txtQuerySensorPeriod.SelStart = 0
    Me.txtQuerySensorPeriod.SelLength = Len(Me.txtQuerySensorPeriod.Text)
End Sub

Private Sub txtQuerySensorPeriod_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtQuerySensorPeriod.Text)
    If Err.Number <> 0 Or Test <> Me.txtQuerySensorPeriod.Text Then
        Beep
        Cancel = True
    Else
        Settings.CloudMonitorQueryPeriod = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtClearTime_GotFocus()
    Me.txtClearTime.SelStart = 0
    Me.txtClearTime.SelLength = Len(Me.txtClearTime.Text)
End Sub

Private Sub txtClearTime_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtClearTime.Text)
    If Err.Number <> 0 Or Test <> Me.txtClearTime.Text Then
        Beep
        Cancel = True
    Else
        Settings.CloudMonitorClearTime = Test
    End If
    On Error GoTo 0
End Sub

Private Sub chkEnableScripts_Click()
    If Me.chkEnableScripts.Value = vbChecked Then
        Me.fraScripts.Enabled = True
    Else
        Me.fraScripts.Enabled = False
    End If
End Sub

Private Sub cmdOpenBefore_Click()
    Me.MSComm.DialogTitle = "Open Before Slew Script"
    Me.MSComm.FileName = Me.txtBeforeScript.Text
    Me.MSComm.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.CancelError = True
    Me.MSComm.InitDir = Me.txtBeforeScript.Text
    Me.MSComm.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtBeforeScript.Text = Me.MSComm.FileName
    Else
        Me.txtBeforeScript.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub cmdOpenAfter_Click()
    Me.MSComm.DialogTitle = "Open After Slew Script"
    Me.MSComm.FileName = Me.txtAfterScript.Text
    Me.MSComm.Filter = "Script (*.vbs;*.wsf)|*.vbs;*.wsf"
    Me.MSComm.FilterIndex = 1
    Me.MSComm.CancelError = True
    Me.MSComm.InitDir = Me.txtAfterScript.Text
    Me.MSComm.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    On Error Resume Next
    Me.MSComm.ShowOpen
    If Err.Number = 0 Then
        Me.txtAfterScript.Text = Me.MSComm.FileName
    Else
        Me.txtAfterScript.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Sub optGuiderCal_Click(Index As Integer)
    Me.optGuiderCal(Index).Value = True
    
    If Index = 0 Then
        Me.optGuiderCal(1).Value = False
        Me.lblAngles.Caption = "Angles determined when the telescope was pointing to the eastern sky."
    Else
        Me.optGuiderCal(0).Value = False
        Me.lblAngles.Caption = "Angles determined when the telescope was pointing to the western sky."
    End If
    
    
End Sub

Private Sub optMountType_Click(Index As Integer)
    Me.optMountType(Index).Value = True
    
    If Index = 0 Then
        Me.optMountType(1).Value = False
        Me.fraGEMSetup.Enabled = True
        Me.lblAngles.Visible = True
        Me.fraRotatorFlip.Enabled = True
        Me.chkOnlyDelayAfterMeridianFlip.Enabled = True
    Else
        Me.optMountType(0).Value = False
        Me.fraGEMSetup.Enabled = False
        Me.fraRotatorFlip.Enabled = False
        Me.chkOnlyDelayAfterMeridianFlip.Value = vbUnchecked
        Me.chkOnlyDelayAfterMeridianFlip.Enabled = False
    End If
End Sub

Private Sub txtDelayAfterSlew_GotFocus()
    Me.txtDelayAfterSlew.SelStart = 0
    Me.txtDelayAfterSlew.SelLength = Len(Me.txtDelayAfterSlew.Text)
End Sub

Private Sub txtDelayAfterSlew_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtDelayAfterSlew.Text)
    If Err.Number <> 0 Or Test > 180 Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.DelayAfterSlew = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtEasternLimit_GotFocus()
    Me.txtEasternLimit.SelStart = 0
    Me.txtEasternLimit.SelLength = Len(Me.txtEasternLimit.Text)
End Sub

Private Sub txtEasternLimit_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtEasternLimit.Text)
    If Err.Number <> 0 Or Test > 180 Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.EasternLimit = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtWesternLimit_GotFocus()
    Me.txtWesternLimit.SelStart = 0
    Me.txtWesternLimit.SelLength = Len(Me.txtWesternLimit.Text)
End Sub

Private Sub txtWesternLimit_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtWesternLimit.Text)
    If Err.Number <> 0 Or Test > 180 Or Test < 0 Then
        Beep
        Cancel = True
    Else
        Settings.WesternLimit = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuideBoxX_GotFocus()
    Me.txtGuideBoxX.SelStart = 0
    Me.txtGuideBoxX.SelLength = Len(Me.txtGuideBoxX.Text)
End Sub

Private Sub txtGuideBoxX_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtGuideBoxX.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtGuideBoxX.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuideBoxX = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuideBoxY_GotFocus()
    Me.txtGuideBoxY.SelStart = 0
    Me.txtGuideBoxY.SelLength = Len(Me.txtGuideBoxY.Text)
End Sub

Private Sub txtGuideBoxY_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtGuideBoxY.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtGuideBoxY.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuideBoxY = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtGuideStarFWHM_GotFocus()
    Me.txtGuideStarFWHM.SelStart = 0
    Me.txtGuideStarFWHM.SelLength = Len(Me.txtGuideStarFWHM.Text)
End Sub

Private Sub txtGuideStarFWHM_Validate(Cancel As Boolean)
    Dim Test As Integer
    On Error Resume Next
    Test = CInt(Me.txtGuideStarFWHM.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtGuideStarFWHM.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuideStarFWHM = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxStarMovement_GotFocus()
    Me.txtMaxStarMovement.SelStart = 0
    Me.txtMaxStarMovement.SelLength = Len(Me.txtMaxStarMovement.Text)
End Sub

Private Sub txtMaxStarMovement_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtMaxStarMovement.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxStarMovement.Text Then
        Beep
        Cancel = True
    Else
        Settings.MaxStarMovement = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMinBright_GotFocus()
    Me.txtMinBright.SelStart = 0
    Me.txtMinBright.SelLength = Len(Me.txtMinBright.Text)
End Sub

Private Sub txtMinBright_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtMinBright.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMinBright.Text Then
        Beep
        Cancel = True
    Else
        Settings.MinimumGuideStarADU = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxBright_GotFocus()
    Me.txtMaxBright.SelStart = 0
    Me.txtMaxBright.SelLength = Len(Me.txtMaxBright.Text)
End Sub

Private Sub txtMaxBright_Validate(Cancel As Boolean)
    Dim Test As Long
    On Error Resume Next
    Test = CLng(Me.txtMaxBright.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxBright.Text Then
        Beep
        Cancel = True
    Else
        Settings.MaximumGuideStarADU = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMinExp_GotFocus()
    Me.txtMinExp.SelStart = 0
    Me.txtMinExp.SelLength = Len(Me.txtMinExp.Text)
End Sub

Private Sub txtMinExp_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMinExp.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMinExp.Text Then
        Beep
        Cancel = True
    Else
        Settings.MinimumGuideStarExposure = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaximumStarFaded_GotFocus()
    Me.txtMaximumStarFaded.SelStart = 0
    Me.txtMaximumStarFaded.SelLength = Len(Me.txtMaximumStarFaded.Text)
End Sub

Private Sub txtMaximumStarFaded_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMaximumStarFaded.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaximumStarFaded.Text Then
        Beep
        Cancel = True
    Else
        Settings.MaximumStarFadedErrors = Test
    End If
    On Error GoTo 0
End Sub


Private Sub txtGuideExposureIncrement_GotFocus()
    Me.txtGuideExposureIncrement.SelStart = 0
    Me.txtGuideExposureIncrement.SelLength = Len(Me.txtGuideExposureIncrement.Text)
End Sub

Private Sub txtGuideExposureIncrement_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtGuideExposureIncrement.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtGuideExposureIncrement.Text Then
        Beep
        Cancel = True
    Else
        Settings.GuideStarExposureIncrement = Test
    End If
    On Error GoTo 0
End Sub

Private Sub txtMaxExp_GotFocus()
    Me.txtMaxExp.SelStart = 0
    Me.txtMaxExp.SelLength = Len(Me.txtMaxExp.Text)
End Sub

Private Sub txtMaxExp_Validate(Cancel As Boolean)
    Dim Test As Double
    On Error Resume Next
    Test = CDbl(Me.txtMaxExp.Text)
    If Err.Number <> 0 Or Test <= 0 Or Test <> Me.txtMaxExp.Text Then
        Beep
        Cancel = True
    Else
        Settings.MaximumGuideStarExposure = Test
    End If
    On Error GoTo 0
End Sub

Private Function ValidateAllEntries() As Boolean
    Dim Cancel As Boolean
    
    Cancel = False
    If Not Cancel Then Call txtCOMNum_Validate(Cancel)
    If Not Cancel Then Call txtHomeRotationAngle_Validate(Cancel)
    If Not Cancel Then Call txtGuiderCalAngle_Validate(Cancel)
    If Not Cancel Then Call txtNorthAngle_Validate(Cancel)
    If Not Cancel Then Call txtPixelScale_Validate(Cancel)
    If Not Cancel Then Call txtDelayAfterSlew_Validate(Cancel)
    If Not Cancel Then txtEasternLimit_Validate (Cancel)
    If Not Cancel Then txtWesternLimit_Validate (Cancel)
    If Not Cancel Then Call txtGuideBoxX_Validate(Cancel)
    If Not Cancel Then Call txtGuideBoxY_Validate(Cancel)
    If Not Cancel Then Call txtGuideStarFWHM_Validate(Cancel)
    If Not Cancel Then Call txtRestartCycles_Validate(Cancel)
    If Not Cancel Then Call txtRestartError_Validate(Cancel)
    If Not Cancel Then Call txtMaxStarMovement_Validate(Cancel)
    If Not Cancel Then Call txtMinBright_Validate(Cancel)
    If Not Cancel Then Call txtMaxBright_Validate(Cancel)
    If Not Cancel Then Call txtMinExp_Validate(Cancel)
    If Not Cancel Then Call txtMaxExp_Validate(Cancel)
    If Not Cancel Then Call txtGuideExposureIncrement_Validate(Cancel)
    If Not Cancel Then Call txtMaxPointingError_Validate(Cancel)
    If Not Cancel Then Call txtPinPointLETimeOut_Validate(Cancel)
    If Not Cancel Then Call txtClearTime_Validate(Cancel)
    If Not Cancel Then Call txtQuerySensorPeriod_Validate(Cancel)
    If Not Cancel Then Call txtCatalogMagMax_Validate(Cancel)
    If Not Cancel Then Call txtCatalogMagMin_Validate(Cancel)
    If Not Cancel Then Call txtSearchArea_Validate(Cancel)
    If Not Cancel Then Call txtMinStarBrightness_Validate(Cancel)
    If Not Cancel Then Call txtMaxNumStars_Validate(Cancel)
    If Not Cancel Then Call txtStandardDeviation_Validate(Cancel)
    If Not Cancel Then Call txtSMTPPort_Validate(Cancel)
    If Not Cancel Then Call txtFocusRetryCount_Validate(Cancel)
    If Not Cancel Then Call txtFocusTimeOut_Validate(Cancel)
    If Not Cancel Then Call txtDDWTimeout_Validate(Cancel)
    If Not Cancel Then Call txtDDWRetryCount_Validate(Cancel)
    If Not Cancel Then Call txtMaximumStarFaded_Validate(Cancel)
    
    ValidateAllEntries = Not Cancel
End Function


Private Function PadFilterName(FilterName As String) As String
    Dim NonSpaceChar As Integer
    Dim SpaceChar As Integer
    Dim Counter As Integer
    Dim NewFilterName As String
    
    NewFilterName = FilterName
    
    Do
        NonSpaceChar = 0
        SpaceChar = 0
        For Counter = 1 To Len(NewFilterName)
            If Mid(NewFilterName, Counter, 1) = " " Then
                SpaceChar = SpaceChar + 1
            Else
                NonSpaceChar = NonSpaceChar + 1
            End If
        Next Counter
        
        If ((NonSpaceChar * 2) + SpaceChar) > 61 Then
            NewFilterName = Left(NewFilterName, Len(NewFilterName) - 1)
        End If
    Loop Until ((NonSpaceChar * 2) + SpaceChar) <= 61
    
    NewFilterName = NewFilterName + Space(61 - ((NonSpaceChar * 2) + SpaceChar))
    
    PadFilterName = NewFilterName
End Function
