Option Strict On
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmTextView
	Inherits System.Windows.Forms.Form
	'   This file is part of WebbIE.
	'
	'    WebbIE is free software: you can redistribute it and/or modify
	'    it under the terms of the GNU General Public License as published by
	'    the Free Software Foundation, either version 3 of the License, or
	'    (at your option) any later version.
	'
	'    WebbIE is distributed in the hope that it will be useful,
	'    but WITHOUT ANY WARRANTY; without even the implied warranty of
	'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	'    GNU General Public License for more details.
	'
	'    You should have received a copy of the GNU General Public License
	'    along with WebbIE.  If not, see <http://www.gnu.org/licenses/>.
	
	
    Public gTargetForm As frmMain
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        On Error Resume Next
		Call Me.Hide()
	End Sub
	
	Private Sub frmTextView_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error Resume Next
        txtTextView.Font = gTargetForm.txtText.Font
	End Sub
	
	Public Sub mnuEditCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditCopy.Click
        On Error Resume Next
		If txtTextView.SelectedText <> "" Then ' only copy if we have something selected
			Call My.Computer.Clipboard.Clear()
			Call My.Computer.Clipboard.SetText(txtTextView.SelectedText)
		End If
	End Sub
	
	Public Sub mnuEditCut_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditCut.Click
        On Error Resume Next
		Dim selStart As Integer
		Call My.Computer.Clipboard.Clear()
		Call My.Computer.Clipboard.SetText(txtTextView.SelectedText)
		selStart = txtTextView.SelectionStart
		txtTextView.Text = VB.Left(txtTextView.Text, txtTextView.SelectionStart) & VB.Right(txtTextView.Text, Len(txtTextView.Text) - txtTextView.SelectionStart - txtTextView.SelectionLength)
		txtTextView.SelectionStart = selStart
        Call txtTextView.ScrollToCaret()
    End Sub

    Public Sub mnuEditFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFind.Click
        On Error Resume Next
        Dim where As Integer

        gfindText = InputBox(modI18N.GetText("Enter text to find"), modI18N.GetText("Search string") & ":", gfindText)
        'if OK is clicked
        If gfindText <> "" Then
            ' Find string in text
            where = InStr(1, txtTextView.Text, gfindText, CompareMethod.Text)
            If where > 0 Then ' If found..
                txtTextView.Focus()
                txtTextView.SelectionStart = where - 1 ' set selection start and
                txtTextView.SelectionLength = Len(gfindText) ' set selection length
                Call txtTextView.ScrollToCaret()
            Else
                MsgBox(modI18N.GetText("Text not found"), MsgBoxStyle.Information, Application.ProductName)
            End If
        End If
    End Sub

    Public Sub mnuEditFindnext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFindnext.Click
        On Error Resume Next
        Dim where As Integer
        Dim originalPos As Integer

        If txtTextView.Text <> "" Then
            originalPos = txtTextView.SelectionStart 'go to start of the found word
            'search onwards from the existing word
            where = InStr(Mid(txtTextView.Text, txtTextView.SelectionStart + 2, Len(txtTextView.Text) - (txtTextView.SelectionStart + 2)), gfindText, CompareMethod.Text)
            If where > 0 Then 'if found
                where += originalPos
                txtTextView.SelectionStart = where
                txtTextView.SelectionLength = Len(gfindText) 'highlight the word
                Call txtTextView.ScrollToCaret()
            Else
                'if unfound, display a warning
                MsgBox(modI18N.GetText("No further occurrences found"), MsgBoxStyle.Information, Application.ProductName)
            End If
        Else 'if there is no text to search
            Call PlayErrorSound()
        End If
    End Sub

    Public Sub mnuEditSelectall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditSelectall.Click
        On Error Resume Next
        txtTextView.SelectionStart = 0
        txtTextView.SelectionLength = Len(txtTextView.Text)
        Call txtTextView.ScrollToCaret()
    End Sub

    Public Sub mnuFileClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileClose.Click
        On Error Resume Next
        Call cmdOK_Click(cmdOK, New System.EventArgs())
    End Sub

    Private Sub txtTextView_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTextView.KeyUp
        On Error Resume Next
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000

        If txtTextView.SelectedText <> "" Then
            mnuEditCopy.Enabled = True
        Else
            mnuEditCopy.Enabled = False
        End If
    End Sub


    Private Sub txtTextView_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles txtTextView.MouseUp
        On Error Resume Next
        Dim Button As Integer = eventArgs.Button \ &H100000
        Dim Shift As Integer = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = eventArgs.X
        Dim Y As Single = eventArgs.Y
        'TODO What should this do?
        'process a mouse action like a key action - check for a selected area.
        Call txtTextView_KeyUp(txtTextView, New System.Windows.Forms.KeyEventArgs(0))
    End Sub

    Private Sub frmTextView_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub
End Class