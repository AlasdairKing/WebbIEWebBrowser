Option Strict On
Option Explicit On
Friend Class frmDownloadProgress
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
	
    'Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
    '    'stop downloading

    '    Call frmMain.Winsock.Close()
    '    Call Me.Hide()
    'End Sub

    Private Sub frmDownloadProgress_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Me.Left = CInt(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width / 2 - Me.Width / 2)
        Me.Top = CInt(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height / 2 - Me.Height / 2)
        lblDownloading.TabIndex = 0
        Me.progressBar.TabIndex = 1
        Me.cmdCancel.TabIndex = 2
    End Sub

    Private Sub tmrProgressSound_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrProgressSound.Tick
        On Error Resume Next
        If Me.Visible Then
            Call PlayProgressSound(CInt(Me.progressBar.Value / 10))
        End If
    End Sub
End Class