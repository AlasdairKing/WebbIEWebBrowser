Public Class frmPassword

    Private mPassword As String

    Public Function GetPassword() As String
        On Error Resume Next
        Return mPassword
    End Function

    Public Sub SetLabel(label As String)
        On Error Resume Next
        lblPassword.Text = label
        txtPassword.AccessibleDescription = label
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        On Error Resume Next
        mPassword = ""
        Me.DialogResult = DialogResult.Cancel
        Call Me.Hide()
    End Sub

    Private Sub frmPassword_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub

    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        On Error Resume Next
        mPassword = txtPassword.Text
        Me.DialogResult = DialogResult.OK
        Call Me.Hide()
    End Sub
End Class