Attribute VB_Name = "ActionList"
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

Public Sub RunActionListAction(clsAction As RunActionList, ByVal ActionListLevel As Integer)
    Dim myCollection As Collection
    Dim Counter As Integer
    Dim Time1 As Date
    Dim Quit As Boolean
    
    Quit = False
    
    If clsAction.LinkToFile Then
        Call AddToStatus("Running " & Mid(clsAction.ActionListName, InStrRev(clsAction.ActionListName, "\") + 1))
    Else
        Call AddToStatus("Running sub-action list.")
    End If
        
    Set myCollection = clsAction.ActionCollection
    
    Select Case clsAction.RepeatMode
        Case 0
            Call SetupSubActionTab(clsAction, ActionListLevel)
            
            Call MainMod.RunActionList(myCollection, ActionListLevel)
            
            Call CloseSubActionTab(clsAction, ActionListLevel)
        Case 1
            Counter = 0
            Do
                Counter = Counter + 1
                Call AddToStatus("Running Action List Iteration #" & Counter)
                                
                Call SetupSubActionTab(clsAction, ActionListLevel)
                Call CleanUpGUIActionList(clsAction.ActionCollection, ActionListLevel)
                
                If Not MainMod.RunActionList(myCollection, ActionListLevel) Then
                    Quit = True
                End If
            
                Call CloseSubActionTab(clsAction, ActionListLevel)
                
            'Added > to check if Counter got ahead of the TimesToRepeat.  Can happen if TimesToRepeat is 0
            Loop Until Counter >= clsAction.TimesToRepeat Or Aborted Or Quit
        Case 2
            Counter = 0
            Time1 = Now
            Do
                Counter = Counter + 1
                Call AddToStatus("Running Action List Iteration #" & Counter & ", " & Format(clsAction.RepeatTime - (DateDiff("s", Time1, Now) / 60), "0.00") & " minutes remaining.")
                
                Call SetupSubActionTab(clsAction, ActionListLevel)
                Call CleanUpGUIActionList(clsAction.ActionCollection, ActionListLevel)
                
                If Not MainMod.RunActionList(myCollection, ActionListLevel) Then
                    Quit = True
                End If
            
                Call CloseSubActionTab(clsAction, ActionListLevel)
            Loop Until (DateDiff("s", Time1, Now) / 60) > clsAction.RepeatTime Or Aborted Or MainMod.SkipToNextSkipToAction Or Quit
        Case 3
            Do While Not Aborted And Not SkipToNextSkipToAction And Not Quit
                Call SetupSubActionTab(clsAction, ActionListLevel)
                Call CleanUpGUIActionList(clsAction.ActionCollection, ActionListLevel)
                
                If Not MainMod.RunActionList(myCollection, ActionListLevel) Then
                    Quit = True
                End If
            
                Call CloseSubActionTab(clsAction, ActionListLevel)
            Loop
        Case Else
            Call AddToStatus("Repeat mode not specified, defaulting to Run Once.")
                
            Call SetupSubActionTab(clsAction, ActionListLevel)
            Call MainMod.RunActionList(myCollection, ActionListLevel)
            Call CloseSubActionTab(clsAction, ActionListLevel)
    End Select
    
    Set myCollection = Nothing
        
    If clsAction.LinkToFile Then
        Call AddToStatus("Done running " & Mid(clsAction.ActionListName, InStrRev(clsAction.ActionListName, "\") + 1))
    Else
        Call AddToStatus("Done running sub-action list.")
    End If
End Sub

Private Sub SetupSubActionTab(clsAction As RunActionList, ByVal ActionListLevel As Integer)
    If MainMod.FollowRunningAction Then
        Call frmMain.AddStuffForSubAction(clsAction)
    End If
    
    On Error GoTo SetupSubActionTabError
    If frmMain.ActionCollections(ActionListLevel) Is clsAction.ActionCollection Then
        frmMain.optRunAborted(ActionListLevel).Enabled = False
        frmMain.optRunMultiple(ActionListLevel).Enabled = False
        frmMain.optRunOnce(ActionListLevel).Enabled = False
        frmMain.optRunPeriod(ActionListLevel).Enabled = False
    End If
    
SetupSubActionTabError:
    On Error GoTo 0
End Sub

Private Sub CloseSubActionTab(clsAction As RunActionList, ByVal ActionListLevel As Integer)
    If MainMod.FollowRunningAction Then
        Call frmMain.RemoveStuffForSubAction(ActionListLevel)
    End If
End Sub
