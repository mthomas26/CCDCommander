VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CCD Commander"
   ClientHeight    =   8655
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   5895
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About Project1"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   479.532
      ScaleMode       =   0  'User
      ScaleWidth      =   479.532
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   8280
      Width           =   1467
   End
   Begin VB.TextBox txtRegistration 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4380
      Width           =   4872
   End
   Begin VB.Label lblOthers 
      ForeColor       =   &H00000000&
      Height          =   2685
      Left            =   780
      TabIndex        =   7
      Tag             =   "App Description"
      Top             =   1620
      Width           =   4875
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   228
      Left            =   780
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   600
      Width           =   4092
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      Caption         =   "http://www.ccdcommander.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   780
      MouseIcon       =   "frmAbout.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1320
      Width           =   2310
   End
   Begin VB.Label lblDescription 
      Caption         =   "Copyright 2004-2009 by Matthew Thomas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Left            =   780
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   840
      Width           =   4092
   End
   Begin VB.Label lblTitle 
      Caption         =   "CCD Commander"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   780
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   180
      Width           =   4092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5552
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5595
      Y1              =   5940
      Y2              =   5940
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   1365
      Left            =   180
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   7200
      Width           =   3870
   End
   Begin VB.Label Label1 
      Caption         =   "All rights reserved."
      Height          =   228
      Left            =   780
      TabIndex        =   8
      Tag             =   "Version"
      Top             =   1080
      Width           =   4092
   End
End
Attribute VB_Name = "frmAbout"
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

' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub lblWeb_Click()
    Dim Ret&
    Ret = ShellExecute(Me.hwnd, "Open", _
        "http://www.ccdcommander.com", _
        "", "", 1)
End Sub

Private Sub lblWeb_MouseDown(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    Me.lblWeb.ForeColor = vbRed
End Sub

Private Sub lblWeb_MouseUp(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
    Me.lblWeb.ForeColor = vbBlue
End Sub


Private Sub Form_Load()
    Call MainMod.SetOnTopMode(Me)
    
    If App.Revision > 1000 Then
        lblVersion.Caption = "Version " & App.Major & "." & (App.Minor + 1) & " Release Candidate " & (App.Revision - 1000)
    ElseIf App.Revision > 200 Then
        lblVersion.Caption = "Version " & App.Major & "." & (App.Minor + 1) & " Beta " & (App.Revision - 200)
    Else
        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    End If
    lblTitle.Caption = App.Title
    lblDescription.Caption = "Copyright " & Chr(169) & " 2004-" & Format$(Date, "yyyy") & " by Matthew Thomas"
    Me.lblOthers.Caption = "CCDSoft5 is copyrighted by Software Bisque, Inc. and " & vbCrLf & String(10, " ") & "Santa Barbara Instruments Group" & _
        vbCrLf & "TheSky6 && TheSkyX are copyrighted by Software Bisque, Inc." & vbCrLf & _
        "ASCOM is copyrighted by the ASCOM Initiative" & vbCrLf & _
        "MaxIm DL is copyrighted by Diffraction Limited" & vbCrLf & _
        "PinPoint is copyrighted by DC-3 Dreams" & vbCrLf & _
        "FocusMax is copyrighted by Larry Weber and Steve Brady" & vbCrLf & _
        "Pyxis is copyrighted by Optec Inc." & vbCrLf & _
        "RCOS && PIR are copyrighted by Optical Systems Incorporated" & vbCrLf & _
        "TAKometer is copyrighted by Don Goldman/AstroDon" & vbCrLf & _
        "Automadome is copyrighted by Software Bisque, Inc." & vbCrLf & _
        "Digital Dome Works is copyrighted by Technical Innovations" & vbCrLf & _
        "Boltwood/Clarity is copyrighted by Diffraction Limited"
        
        
    Me.lblDisclaimer.Caption = "CCDCommander is provided to you " & Chr(34) & "AS IS" & Chr(34) & _
        " without any warantee, implied or otherwise. By using this program you agree to assume all risks associated " & _
        "with the use of CCDCommander, including but not limited to the risks or program errors, damage to or loss of " & _
        "data, programs, or equipment."
        
    Call RegistrationInfo
End Sub

Private Sub cmdOk_Click()
        Unload Me
End Sub

Public Sub RegistrationInfo()
    Me.txtRegistration.Text = "This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by " & _
                            "the Free Software Foundation, version 3 of the License." & vbCrLf & vbCrLf & _
                            "This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of " & _
                            "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details." & vbCrLf & vbCrLf & _
                            "You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>."
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

