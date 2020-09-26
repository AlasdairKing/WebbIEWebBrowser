Public Class frmOptionsForm

    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        On Error Resume Next
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Call Me.Close()
    End Sub

    Private Sub frmOptionsForm_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error Resume Next
        My.Settings.ShowToolbar = chkShowToolbar.Checked
        My.Settings.ShowImages = chkShowImages.Checked
        My.Settings.AllowPopupWindows = chkAllowPopupWindows.Checked
        My.Settings.AllowMessages = chkAllowMessages.Checked
        My.Settings.NavigationSounds = chkNavigationSounds.Checked
        My.Settings.QuickKeys = chkUseQuickkeys.Checked
        My.Settings.ToolbarCaptions = chkShowToolbarCaptions.Checked
        My.Settings.UseIEHomepage = chkIEHomepage.Checked
        My.Settings.NumberLinks = chkNumberLinks.Checked
    End Sub

    Private Sub frmOptionsForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)

        'User Preferences
        chkShowToolbar.Checked = My.Settings.ShowToolbar
        chkShowImages.Checked = My.Settings.ShowImages
        chkAllowPopupWindows.Checked = My.Settings.AllowPopupWindows
        chkAllowMessages.Checked = My.Settings.AllowMessages
        chkNavigationSounds.Checked = My.Settings.NavigationSounds
        chkUseQuickkeys.Checked = My.Settings.QuickKeys
        chkShowToolbarCaptions.Checked = My.Settings.ToolbarCaptions
        chkShowToolbarCaptions.Enabled = chkShowToolbar.Checked
        chkIEHomepage.Checked = My.Settings.UseIEHomepage
        chkNumberLinks.Checked = My.Settings.NumberLinks
    End Sub

    Private Sub chkShowToolbarCaptions_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowToolbarCaptions.CheckedChanged

    End Sub

    Private Sub chkIEHomepage_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkIEHomepage.CheckedChanged

    End Sub

    Private Sub chkShowToolbar_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowToolbar.CheckedChanged
        On Error Resume Next
        chkShowToolbarCaptions.Enabled = chkShowToolbar.Checked
    End Sub

    Private Sub chkNavigationSounds_CheckedChanged(sender As Object, e As EventArgs) Handles chkNavigationSounds.CheckedChanged

    End Sub
End Class