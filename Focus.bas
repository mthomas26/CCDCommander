Attribute VB_Name = "Focus"
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

Public Enum FocusControl
    None = 0
    FocusMax = 1
    FocusMaxAcquireStar = 2
    CCDSoftAtFocus = 3
    CCDSoftAtFocus2 = 4
    MaxImFocus = 5
    TheSkyXAtFocus2 = 6
    TheSkyXAtFocus3 = 7
End Enum

Dim objFocuser As Object

Public Sub FocusMaxAction(clsAction As FocusAction)
    Dim SetFastDownload As Boolean
    Dim FocusCounter As Integer

    Call AddToStatus("In Focus Action.")
    
    If objFocuser Is Nothing Then
        Call AddToStatus("No focuser selected - skippping focus action.")
    Else
        
        If clsAction.ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 Then
            Call AddToStatus("Setting filter to: " & frmOptions.lstFilters.List(clsAction.ImagerFilter))
            Call AddToStatus("Forcing filter change...")
            Call Camera.ForceFilterChange(clsAction.ImagerFilter)
        End If
        
        If Not Camera.objCameraControl.FastDownloadEnabled And clsAction.FastReadout And Not Aborted Then
            Call AddToStatus("Setting FastReadout true.")
            Camera.objCameraControl.FastDownloadEnabled = True
            SetFastDownload = True
        End If
        
        If Not Aborted Then
            Call AddToStatus("Starting focus run...")
            
            FocusCounter = 0
            Do While Not objFocuser.Focus(clsAction) And FocusCounter < Settings.RetryFocusCount And frmOptions.chkRetryFocusRunOnFailure.Value = vbChecked
            
                If Aborted Then Exit Do
                
                Call AddToStatus("Retrying focus run...")
                FocusCounter = FocusCounter + 1
                
                Call Wait(2)
            Loop
        End If
        
        If SetFastDownload Then
            Call AddToStatus("Setting FastReadout false.")
            Camera.objCameraControl.FastDownloadEnabled = False
        End If
    
        If clsAction.ImagerFilter >= 0 And frmOptions.lstFilters.ListCount > 0 And Not Aborted Then
            Call AddToStatus("Resetting filter to: " & frmOptions.lstFilters.List(clsAction.ImagerFilter))
            Call AddToStatus("Forcing filter change...")
            Call Camera.ForceFilterChange(clsAction.ImagerFilter)
        End If
    End If
End Sub

Public Sub FocuserFilterOffset(CurrentFilter As Integer, NewFilter As Integer)
    Dim FilterOffset As Integer
    Dim FocuserPosition As Integer
    
    FilterOffset = frmOptions.lstFiltersFocusOffset.ItemData(NewFilter) - frmOptions.lstFiltersFocusOffset.ItemData(CurrentFilter)

    If FilterOffset <> 0 Then
        Call AddToStatus("Offsetting focuser " & FilterOffset & " steps for " & frmOptions.lstFilters.List(NewFilter) & " filter.")
        FocuserPosition = objFocuser.OffsetFocuser(FilterOffset)
        Call AddToStatus("Done offseting filter.  Final position = " & FocuserPosition)
    End If
End Sub

Public Sub FocusSetup()
    Select Case frmOptions.lstFocuserControl.ListIndex
        Case FocusControl.None
            Set objFocuser = Nothing
        Case FocusControl.FocusMax
            Set objFocuser = New clsFocusMaxControl
        Case FocusControl.FocusMaxAcquireStar
            Set objFocuser = New clsFocusMaxAcqControl
        Case FocusControl.CCDSoftAtFocus
            Set objFocuser = New clsAtFocusControl
        Case FocusControl.CCDSoftAtFocus2
            Set objFocuser = New clsAtFocus2Control
        Case FocusControl.MaxImFocus
            Set objFocuser = New clsMaxImFocusControl
        Case FocusControl.TheSkyXAtFocus2
            Set objFocuser = New clsTheSkyXAtFocus2Control
        Case FocusControl.TheSkyXAtFocus3
            Set objFocuser = New clsTheSkyXAtFocus3Control
    End Select
    
    If Not (objFocuser Is Nothing) Then _
        Call objFocuser.ConnectToFocuser
End Sub

Public Sub FocusUnload()
    Set objFocuser = Nothing
End Sub
