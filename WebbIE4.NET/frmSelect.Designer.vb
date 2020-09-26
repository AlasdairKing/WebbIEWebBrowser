<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSelect
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
    Public WithEvents lstSelect As System.Windows.Forms.ListBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOKfrmSelect As System.Windows.Forms.Button
	Public WithEvents lblSelect As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelect))
        Me.lstSelect = New System.Windows.Forms.ListBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOKfrmSelect = New System.Windows.Forms.Button()
        Me.lblSelect = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lstSelect
        '
        Me.lstSelect.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstSelect.BackColor = System.Drawing.SystemColors.Window
        Me.lstSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstSelect.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstSelect.IntegralHeight = False
        Me.lstSelect.ItemHeight = 30
        Me.lstSelect.Location = New System.Drawing.Point(0, 0)
        Me.lstSelect.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.lstSelect.Name = "lstSelect"
        Me.lstSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstSelect.Size = New System.Drawing.Size(814, 536)
        Me.lstSelect.TabIndex = 1
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(826, 106)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(162, 76)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Tag = "frmSelect.cmdCancel"
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOKfrmSelect
        '
        Me.cmdOKfrmSelect.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOKfrmSelect.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOKfrmSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOKfrmSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOKfrmSelect.Location = New System.Drawing.Point(826, 16)
        Me.cmdOKfrmSelect.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.cmdOKfrmSelect.Name = "cmdOKfrmSelect"
        Me.cmdOKfrmSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOKfrmSelect.Size = New System.Drawing.Size(162, 76)
        Me.cmdOKfrmSelect.TabIndex = 2
        Me.cmdOKfrmSelect.Text = "OK"
        Me.cmdOKfrmSelect.UseVisualStyleBackColor = False
        '
        'lblSelect
        '
        Me.lblSelect.AutoSize = True
        Me.lblSelect.BackColor = System.Drawing.SystemColors.Control
        Me.lblSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSelect.Location = New System.Drawing.Point(480, 222)
        Me.lblSelect.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblSelect.Name = "lblSelect"
        Me.lblSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSelect.Size = New System.Drawing.Size(163, 30)
        Me.lblSelect.TabIndex = 0
        Me.lblSelect.Text = "&Select an option"
        '
        'frmSelect
        '
        Me.AcceptButton = Me.cmdOKfrmSelect
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(1003, 533)
        Me.ControlBox = False
        Me.Controls.Add(Me.lstSelect)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOKfrmSelect)
        Me.Controls.Add(Me.lblSelect)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(3, 23)
        Me.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSelect"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Tag = "frmSelect"
        Me.Text = "Select an option"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class