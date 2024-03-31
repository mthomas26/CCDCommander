Attribute VB_Name = "Comment"
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

Public Sub RunCommentAction(clsAction As clsCommentAction)
    'just put the comment into the log
    If InStr(clsAction.Comment, vbCrLf) Then
        Call MainMod.AddToStatus("Comment: " & vbCrLf & clsAction.Comment)
    Else
        Call MainMod.AddToStatus("Comment: " & clsAction.Comment)
    End If

    If frmOptions.chkEMailAlert(EMailAlertIndexes.CommentActions).Value = vbChecked And clsAction.SendAsEmail Then
        'Send e-mail!
        Call EMail.SendEMail(frmMain, "CCD Commander Comment", Format(Now, "hh:mm:ss") & vbCrLf & clsAction.Comment)
    End If
End Sub
