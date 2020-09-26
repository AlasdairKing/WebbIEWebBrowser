Public Class frmSlider

    Public Title As String

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        Try
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Call Me.Hide()
        Catch
        End Try
    End Sub

    Private Sub cmdOK_Click(sender As System.Object, e As System.EventArgs) Handles cmdOK.Click
        On Error Resume Next
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Hide()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub frmSlider_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Me.Text = Title & " (" & trbMain.Minimum & " " & modI18N.GetText("to") & " " & trbMain.Maximum & ")"
        Me.trbMain.AccessibleName = Title
    End Sub
End Class