VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAutoguiderError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autoguider Error"
   ClientHeight    =   8220
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9360
   HelpContextID   =   400
   Icon            =   "AutoguiderError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   960
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   492
      Left            =   2292
      TabIndex        =   5
      Top             =   7620
      Width           =   972
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   492
      Left            =   6096
      TabIndex        =   4
      Top             =   7620
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8940
      Top             =   7440
   End
   Begin MSChart20Lib.MSChart chtXError 
      CausesValidation=   0   'False
      Height          =   3432
      Left            =   72
      OleObjectBlob   =   "AutoguiderError.frx":030A
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   9192
   End
   Begin MSChart20Lib.MSChart chtYError 
      CausesValidation=   0   'False
      Height          =   3432
      Left            =   96
      OleObjectBlob   =   "AutoguiderError.frx":1C9F
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4020
      Width           =   9192
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Y-axis Error"
      Height          =   192
      Left            =   4296
      TabIndex        =   3
      Top             =   3780
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "X-axis Error"
      Height          =   192
      Left            =   4296
      TabIndex        =   0
      Top             =   60
      Width           =   816
   End
End
Attribute VB_Name = "frmAutoguiderError"
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

Private Sub cmdClear_Click()
    Me.chtXError.RowCount = 0
    Me.chtYError.RowCount = 0
End Sub

Private Sub cmdSave_Click()
    Dim FileNo As Integer
    Dim Counter As Long
    
    Me.CommonDialog.Filter = "Comma Delimited (*.csv)|*.csv|Text (*.txt)|*.txt"
    Me.CommonDialog.FilterIndex = 1
    Me.CommonDialog.DialogTitle = "Save Autoguider Data"
    Me.CommonDialog.FileName = "AutoguiderData.csv"
    Me.CommonDialog.InitDir = App.Path
    Me.CommonDialog.flags = cdlOFNHideReadOnly + cdlOFNNoValidate + cdlOFNPathMustExist
    Me.CommonDialog.CancelError = True
    
    On Error GoTo cmdSaveError
    Me.CommonDialog.ShowSave
    
    FileNo = FreeFile()
    Open Me.CommonDialog.FileName For Output As #FileNo
    
    Print #FileNo, "Time,XError,YError"
    
    Timer1.Enabled = False
    For Counter = 1 To Me.chtXError.RowCount
        Me.chtXError.Row = Counter
        Me.chtYError.Row = Counter
        
        Print #FileNo, Me.chtXError.RowLabel & "," & Me.chtXError.Data & "," & Me.chtYError.Data
    Next Counter
    Timer1.Enabled = True
    
    Close #FileNo
cmdSaveError:
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    Me.Top = GetMySetting("WindowPositions", "AGErrorTop", 200)
    Me.Left = GetMySetting("WindowPositions", "AGErrorLeft", 200)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
        frmMain.mnuViewItem(1).Checked = False
    Else
        Me.Timer1.Enabled = False
    End If

    If Me.WindowState <> vbMinimized Then
        Call SaveMySetting("WindowPositions", "AGErrorTop", Me.Top)
        Call SaveMySetting("WindowPositions", "AGErrorLeft", Me.Left)
    End If
End Sub

Private Sub GraphDeleteFirstPoints(myChart As MSChart)
    Dim Counter As Long
    Dim RowLabel As String
    Dim Data As Double
    
    For Counter = 1 To myChart.RowCount - 3000
        myChart.Row = Counter + 3000
        RowLabel = myChart.RowLabel
        Data = myChart.Data
        
        myChart.Row = Counter
        myChart.RowLabel = RowLabel
        myChart.Data = Data
    Next Counter
    
    myChart.RowCount = myChart.RowCount - 3000
End Sub

Private Sub Timer1_Timer()
    Dim myGuideError As AutoguiderError
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
        
    On Error GoTo AutoguiderErrorGraphTimerError
    
    Do While colAutoguiderErrors.Count > 0
        Set myGuideError = colAutoguiderErrors.Item(1)
        Call colAutoguiderErrors.Remove(1)
        
        If Me.chtXError.RowCount = 32767 Then
            Call GraphDeleteFirstPoints(Me.chtXError)
        End If
        
        Me.chtXError.RowCount = Me.chtXError.RowCount + 1
        Me.chtXError.Row = Me.chtXError.RowCount
        Me.chtXError.RowLabel = Format(myGuideError.TimeStamp, "h:mm:ss")
        Me.chtXError.Plot.Axis(VtChAxisIdX).Tick.Style = VtChAxisTickStyleNone
        Me.chtXError.Data = myGuideError.XError

        If Me.chtYError.RowCount = 32767 Then
            Call GraphDeleteFirstPoints(Me.chtYError)
        End If
        
        Me.chtYError.RowCount = Me.chtYError.RowCount + 1
        Me.chtYError.Row = Me.chtYError.RowCount
        Me.chtYError.RowLabel = Format(myGuideError.TimeStamp, "h:mm:ss")
        Me.chtYError.Plot.Axis(VtChAxisIdX).Tick.Style = VtChAxisTickStyleNone
        Me.chtYError.Data = myGuideError.YError
    Loop

AutoguiderErrorGraphTimerError:
    ErrorNum = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description
        
    On Error GoTo 0
    
    If ErrorNum = &H80010005 Then
        'error is automation error - I can just ignore this and try again next time around
    ElseIf ErrorNum <> 0 Then
        Call Err.Raise(ErrorNum, ErrorSource, ErrorDescription)
    End If
End Sub
