Option Strict On
Option Explicit On
Friend Class frmTextareaInput
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
	
	
	Private mobjTextAreaNode As mshtml.IHTMLDOMNode ' the text area node we're working on
    Public gAreaLabel As String ' holds any label given the input by the page
	
    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error Resume Next
        Call Me.Hide()
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        On Error Resume Next
        'exit the form

        Call Me.Hide()
        Call frmMain.UpdateTextarea((txtInput.Text), mobjTextAreaNode)
    End Sub

    Private Sub frmTextareaInput_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            txtInput.Font = frmMain.txtText.Font
            If Len(gAreaLabel) > 0 Then
                'we've had the areaLabel variable set to something: this should be
                'the label applied to the element in the web page
                Me.Text = gAreaLabel
                lblInputText.Text = gAreaLabel
            Else
                'nope, use the default phrase
                Me.Text = modI18N.GetText("Input text")
                lblInputText.Text = Me.Text
            End If
            Call txtInput.Focus()
            Call Application.DoEvents()
            txtInput.SelectionStart = 0
            txtInput.SelectionLength = 0
        Catch
        End Try
    End Sub

    Public Sub Populate(ByRef textAreaNode As mshtml.IHTMLDOMNode)
        Try
            'show the original text
            Dim e As mshtml.IHTMLElement

            e = CType(textAreaNode, mshtml.IHTMLElement)
            txtInput.Text = e.innerText
            mobjTextAreaNode = textAreaNode
        Catch
        End Try
    End Sub

    Private Sub txtInput_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInput.KeyDown
        Try
            Dim KeyCode As Integer = eventArgs.KeyCode
            Dim Shift As Integer = eventArgs.KeyData \ &H10000

            If KeyCode = CInt(System.Windows.Forms.Keys.A) And eventArgs.Control Then
                txtInput.SelectionStart = 0
                txtInput.SelectionLength = Len(txtInput.Text)
                Call txtInput.ScrollToCaret()
            ElseIf KeyCode = System.Windows.Forms.Keys.Return And eventArgs.Control Then
                Call cmdOK_Click(cmdOK, New System.EventArgs())
            End If
        Catch
        End Try
    End Sub

    Private Sub frmTextareaInput_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub

End Class