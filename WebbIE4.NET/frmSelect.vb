Option Strict On
Option Explicit On
Friend Class frmSelect
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
	
	'This module puts up the selection menu for webforms that appears when the user clicks on a drop-down
	'box in a webpage.   Alasdair 20 September 2002
	
	
    Private mSelectNumber As Integer ' the index of the select item in the selects array
    Public gTargetForm As frmMain

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error Resume Next
        Call Me.Hide()
    End Sub

    Private Sub cmdOKfrmSelect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOKfrmSelect.Click
        On Error Resume Next
        Call frmMain.UpdateSelection(mSelectNumber, lstSelect.SelectedIndex + 1)
        Call Me.Hide()
    End Sub

    Private Sub frmSelect_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error Resume Next
        lstSelect.Font = frmMain.txtText.Font
    End Sub

    Private Sub frmSelect_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        On Error Resume Next
        Dim KeyCode As Integer = eventArgs.KeyCode
        Dim Shift As Integer = eventArgs.KeyData \ &H10000

        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Hide()
    End Sub

    Public Sub Populate(ByRef sNumber As Integer)
        Try
            Dim i As Integer ' counter

            mSelectNumber = sNumber
            'clear the list
            Call lstSelect.Items.Clear()
            'populate the list
            For i = 1 To selects(mSelectNumber).size
                Call lstSelect.Items.Add(selects(mSelectNumber).options(i))
            Next i
            'select the current item
            lstSelect.SelectedIndex = selects(mSelectNumber).selected - 1
        Catch
        End Try
    End Sub
	
	Private Sub lstSelect_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lstSelect.KeyDown
        Try
            Dim KeyCode As Integer = eventArgs.KeyCode
            Dim Shift As Integer = eventArgs.KeyData \ &H10000
            'respond to user actions


            If KeyCode = System.Windows.Forms.Keys.Escape Then
                'exit form
                Call Me.Hide()
                Call Me.gTargetForm.Show()
            End If
            If KeyCode = System.Windows.Forms.Keys.Return Or KeyCode = System.Windows.Forms.Keys.Space Then
                'user has chosen an item
                If lstSelect.SelectedIndex < 0 Then
                    'don't do anything - need user to select an item
                Else
                    'user has selected an item
                    'set the correct new one in selects
                    'now return to the main form, telling it to update the page with the
                    'new selection
                    Call cmdOKfrmSelect_Click(cmdOKfrmSelect, New System.EventArgs())
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub frmSelect_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub
End Class