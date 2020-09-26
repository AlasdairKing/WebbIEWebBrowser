<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDownloadProgress
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents tmrProgressSound As System.Windows.Forms.Timer
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents progressBar As System.Windows.Forms.ProgressBar
	Public WithEvents lblDownloading As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDownloadProgress))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.tmrProgressSound = New System.Windows.Forms.Timer(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.progressBar = New System.Windows.Forms.ProgressBar()
        Me.lblDownloading = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'tmrProgressSound
        '
        Me.tmrProgressSound.Enabled = True
        Me.tmrProgressSound.Interval = 1500
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(176, 88)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(81, 33)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Tag = "frmDownloadProgress.cmdCancel"
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'progressBar
        '
        Me.progressBar.Location = New System.Drawing.Point(8, 56)
        Me.progressBar.Name = "progressBar"
        Me.progressBar.Size = New System.Drawing.Size(425, 25)
        Me.progressBar.TabIndex = 0
        '
        'lblDownloading
        '
        Me.lblDownloading.AutoSize = True
        Me.lblDownloading.BackColor = System.Drawing.SystemColors.Control
        Me.lblDownloading.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDownloading.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDownloading.Location = New System.Drawing.Point(16, 8)
        Me.lblDownloading.Name = "lblDownloading"
        Me.lblDownloading.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDownloading.Size = New System.Drawing.Size(69, 13)
        Me.lblDownloading.TabIndex = 2
        Me.lblDownloading.Text = "Downloading"
        '
        'frmDownloadProgress
        '
        Me.AcceptButton = Me.cmdCancel
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(436, 126)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.progressBar)
        Me.Controls.Add(Me.lblDownloading)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDownloadProgress"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Tag = "frmDownloadProgress"
        Me.Text = "Downloading"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class