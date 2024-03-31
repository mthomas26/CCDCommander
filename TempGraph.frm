VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTempGraph 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CCD Temperature"
   ClientHeight    =   4485
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9360
   HelpContextID   =   500
   Icon            =   "TempGraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9360
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   612
      Left            =   3540
      TabIndex        =   7
      Top             =   3660
      Width           =   1032
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   8820
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox lblUpdate 
      Height          =   312
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "5"
      Top             =   3780
      Width           =   312
   End
   Begin VB.VScrollBar vsUpdate 
      CausesValidation=   0   'False
      Height          =   312
      Left            =   7620
      Max             =   30
      Min             =   5
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3780
      Value           =   5
      Width           =   192
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   612
      Left            =   2100
      TabIndex        =   2
      Top             =   3660
      Width           =   1032
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   612
      Left            =   660
      TabIndex        =   1
      Top             =   3660
      Width           =   1032
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8940
      Top             =   3180
   End
   Begin MSChart20Lib.MSChart Chart 
      CausesValidation=   0   'False
      Height          =   3432
      Left            =   60
      OleObjectBlob   =   "TempGraph.frx":030A
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   9192
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "seconds"
      Height          =   195
      Left            =   7920
      TabIndex        =   5
      Top             =   3840
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Update Interval"
      Height          =   195
      Left            =   6180
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmTempGraph"
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
    Me.Timer1.Enabled = False
    
    Me.Chart.RowCount = 0
    
    If Not (Camera.objCameraControl Is Nothing) Then
        Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = Int(Camera.objCameraControl.Temperature - 1.5)
        Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Int(Camera.objCameraControl.Temperature + 2.5)
        Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = Int(Camera.objCameraControl.CoolerPower - 9.5)
        Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = Int(Camera.objCameraControl.CoolerPower + 10.5)
    
        Me.Timer1.Interval = CDbl(Me.lblUpdate.Text) * 1000
        Me.Timer1.Enabled = True
    End If
End Sub

Public Sub Clear()
    Call cmdClear_Click
End Sub

Private Sub cmdSave_Click()
    Dim FileNo As Integer
    Dim Counter As Long
    
    Me.CommonDialog.Filter = "Comma Delimited (*.csv)|*.csv|Text (*.txt)|*.txt"
    Me.CommonDialog.FilterIndex = 1
    Me.CommonDialog.DialogTitle = "Save CCD Temperature Data"
    Me.CommonDialog.FileName = "CCDTemperatureData.csv"
    Me.CommonDialog.InitDir = App.Path
    Me.CommonDialog.flags = cdlOFNHideReadOnly + cdlOFNNoValidate + cdlOFNPathMustExist
    Me.CommonDialog.CancelError = True
    
    On Error GoTo cmdSaveError
    Me.CommonDialog.ShowSave
    
    FileNo = FreeFile()
    Open Me.CommonDialog.FileName For Output As #FileNo
    
    Print #FileNo, "Time,Temperature,Cooler %"
    
    Timer1.Enabled = False
    For Counter = 1 To Me.Chart.RowCount
        Me.Chart.Row = Counter
        
        Me.Chart.Column = 1
        Print #FileNo, Me.Chart.RowLabel & "," & Me.Chart.Data & ",";
        Me.Chart.Column = 2
        Print #FileNo, Me.Chart.Data
    Next Counter
    Timer1.Enabled = True
    
    Close #FileNo
cmdSaveError:
    On Error GoTo 0
End Sub

Private Sub cmdStop_Click()
    If Me.Timer1.Enabled = True Then
        cmdStop.Caption = "&Start"
        Me.Timer1.Enabled = False
    Else
        cmdStop.Caption = "&Stop"
        
        Me.Chart.RowCount = 0
        Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = Int(Camera.objCameraControl.Temperature - 1.5)
        Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Int(Camera.objCameraControl.Temperature + 2.5)
        If Camera.objCameraControl.CoolerPower = 0 Then
            Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = 0
        Else
            Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = Int(Camera.objCameraControl.CoolerPower - 9.5)
        End If
        If Camera.objCameraControl.CoolerPower = 100 Then
            Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = 100
        Else
            Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = Int(Camera.objCameraControl.CoolerPower + 10.5)
        End If
    
        Me.Timer1.Interval = Me.vsUpdate.Value * 1000
        Me.Timer1.Enabled = True
    End If
End Sub

Public Sub SetupTempGraph()
    Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = Int(Camera.objCameraControl.Temperature - 1.5)
    Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Int(Camera.objCameraControl.Temperature + 2.5)
    If Camera.objCameraControl.CoolerPower = 0 Then
        Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = 0
    Else
        Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = Int(Camera.objCameraControl.CoolerPower - 9.5)
    End If
    If Camera.objCameraControl.CoolerPower = 100 Then
        Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = 100
    Else
        Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = Int(Camera.objCameraControl.CoolerPower + 10.5)
    End If

    Me.vsUpdate.Value = CInt(GetMySetting("TempGraph", "UpdateInterval", "5"))

    Me.Timer1.Interval = Me.vsUpdate.Value * 1000
End Sub

Private Sub Form_Load()
    Me.Top = GetMySetting("WindowPositions", "TempGraphTop", 200)
    Me.Left = GetMySetting("WindowPositions", "TempGraphLeft", 200)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Visible = False
        frmMain.mnuViewItem(0).Checked = False
    Else
        Me.Timer1.Enabled = False
    End If

    If Me.WindowState <> vbMinimized Then
        Call SaveMySetting("WindowPositions", "TempGraphTop", Me.Top)
        Call SaveMySetting("WindowPositions", "TempGraphLeft", Me.Left)
    End If
End Sub

Private Sub lblUpdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 104 Or KeyCode = vbKeyUp Then
        If Me.vsUpdate.Value > Me.vsUpdate.Min Then
            Me.vsUpdate.Value = Me.vsUpdate.Value - 1
        End If
    ElseIf KeyCode = 98 Or KeyCode = vbKeyDown Then
        If Me.vsUpdate.Value < Me.vsUpdate.Max Then
            Me.vsUpdate.Value = Me.vsUpdate.Value + 1
        End If
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
    Dim ErrorNum As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String
        
    On Error GoTo TempGraphTimerError
    
    If Camera.objCameraControl.CoolerState And Camera.objCameraControl.CoolerPower < 101 And Camera.objCameraControl.Temperature < 100 Then
        If Me.Chart.RowCount = 32767 Then
            Call GraphDeleteFirstPoints(Me.Chart)
        End If
    
        Me.Chart.RowCount = Me.Chart.RowCount + 1
        Me.Chart.Row = Me.Chart.RowCount
        Me.Chart.RowLabel = Format(Time, "h:mm:ss")
        Me.Chart.Plot.Axis(VtChAxisIdX).Tick.Style = VtChAxisTickStyleNone
    
        Me.Chart.Column = 1
        Me.Chart.Data = Camera.objCameraControl.Temperature
    
        If Me.Chart.Data < Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum Then
            Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = Int(Me.Chart.Data - 0.5)
        ElseIf Me.Chart.Data > Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum Then
            Me.Chart.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = Int(Me.Chart.Data + 1.5)
        End If
    
        Me.Chart.Column = 2
        Me.Chart.Data = Camera.objCameraControl.CoolerPower
    
        If Me.Chart.Data < Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum Then
            If Me.Chart.Data < 0.5 Then
                Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = 0
            Else
                Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Minimum = Int(Me.Chart.Data - 0.5)
            End If
        ElseIf Me.Chart.Data > Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum Then
            If Me.Chart.Data > (100 - 1.5) Then
                Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = 100
            Else
                Me.Chart.Plot.Axis(VtChAxisIdY2).ValueScale.Maximum = Int(Me.Chart.Data + 1.5)
            End If
        End If
    End If

TempGraphTimerError:
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

Private Sub vsUpdate_Change()
    Me.Timer1.Interval = Me.vsUpdate.Value * 1000
    Me.lblUpdate.Text = Me.vsUpdate.Value
    Me.lblUpdate.DataChanged = True
    
    Call SaveMySetting("TempGraph", "UpdateInterval", Me.lblUpdate.Text)
End Sub
