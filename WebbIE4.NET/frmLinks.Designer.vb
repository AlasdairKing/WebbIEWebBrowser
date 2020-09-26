<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmLinks
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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLinks))
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.pnlControls = New System.Windows.Forms.Panel()
        Me.cmdGo = New System.Windows.Forms.Button()
        Me.optSortAlphabetical = New System.Windows.Forms.RadioButton()
        Me.optSortPageOrder = New System.Windows.Forms.RadioButton()
        Me.lstLinks = New System.Windows.Forms.ListBox()
        Me.pnlControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(314, 300)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(5, 8, 5, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(142, 71)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.TabStop = False
        Me.cmdCancel.Tag = ""
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'pnlControls
        '
        Me.pnlControls.Controls.Add(Me.cmdGo)
        Me.pnlControls.Controls.Add(Me.optSortAlphabetical)
        Me.pnlControls.Controls.Add(Me.optSortPageOrder)
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlControls.Location = New System.Drawing.Point(0, 583)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(636, 158)
        Me.pnlControls.TabIndex = 4
        '
        'cmdGo
        '
        Me.cmdGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdGo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGo.Location = New System.Drawing.Point(464, 70)
        Me.cmdGo.Margin = New System.Windows.Forms.Padding(5, 8, 5, 8)
        Me.cmdGo.Name = "cmdGo"
        Me.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGo.Size = New System.Drawing.Size(158, 71)
        Me.cmdGo.TabIndex = 5
        Me.cmdGo.Text = "Go"
        Me.cmdGo.UseVisualStyleBackColor = False
        '
        'optSortAlphabetical
        '
        Me.optSortAlphabetical.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.optSortAlphabetical.BackColor = System.Drawing.SystemColors.Control
        Me.optSortAlphabetical.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSortAlphabetical.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSortAlphabetical.Location = New System.Drawing.Point(14, 70)
        Me.optSortAlphabetical.Margin = New System.Windows.Forms.Padding(5, 8, 5, 8)
        Me.optSortAlphabetical.Name = "optSortAlphabetical"
        Me.optSortAlphabetical.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSortAlphabetical.Size = New System.Drawing.Size(562, 77)
        Me.optSortAlphabetical.TabIndex = 6
        Me.optSortAlphabetical.TabStop = True
        Me.optSortAlphabetical.Tag = "frmLinks.optSort(1)"
        Me.optSortAlphabetical.Text = "Sort by &Alphabetical order"
        Me.optSortAlphabetical.UseVisualStyleBackColor = False
        '
        'optSortPageOrder
        '
        Me.optSortPageOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.optSortPageOrder.BackColor = System.Drawing.SystemColors.Control
        Me.optSortPageOrder.Checked = True
        Me.optSortPageOrder.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSortPageOrder.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSortPageOrder.Location = New System.Drawing.Point(14, 2)
        Me.optSortPageOrder.Margin = New System.Windows.Forms.Padding(5, 8, 5, 8)
        Me.optSortPageOrder.Name = "optSortPageOrder"
        Me.optSortPageOrder.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSortPageOrder.Size = New System.Drawing.Size(578, 77)
        Me.optSortPageOrder.TabIndex = 4
        Me.optSortPageOrder.TabStop = True
        Me.optSortPageOrder.Tag = "frmLinks.optSort(0)"
        Me.optSortPageOrder.Text = "Sort by &Page order"
        Me.optSortPageOrder.UseVisualStyleBackColor = False
        '
        'lstLinks
        '
        Me.lstLinks.AccessibleName = "&Links"
        Me.lstLinks.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstLinks.FormattingEnabled = True
        Me.lstLinks.IntegralHeight = False
        Me.lstLinks.ItemHeight = 30
        Me.lstLinks.Location = New System.Drawing.Point(0, 0)
        Me.lstLinks.Name = "lstLinks"
        Me.lstLinks.Size = New System.Drawing.Size(636, 583)
        Me.lstLinks.TabIndex = 8
        '
        'frmLinks
        '
        Me.AcceptButton = Me.cmdGo
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(636, 741)
        Me.Controls.Add(Me.lstLinks)
        Me.Controls.Add(Me.pnlControls)
        Me.Controls.Add(Me.cmdCancel)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(222, 261)
        Me.Margin = New System.Windows.Forms.Padding(5, 8, 5, 8)
        Me.Name = "frmLinks"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Tag = "frmLinks"
        Me.Text = "Webpage Links"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlControls.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Public WithEvents cmdGo As System.Windows.Forms.Button
    Public WithEvents optSortAlphabetical As System.Windows.Forms.RadioButton
    Public WithEvents optSortPageOrder As System.Windows.Forms.RadioButton
    Friend WithEvents lstLinks As System.Windows.Forms.ListBox
#End Region
End Class