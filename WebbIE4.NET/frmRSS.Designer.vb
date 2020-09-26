<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRSS
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
    Public WithEvents lstItems As System.Windows.Forms.ListBox
    Public WithEvents btnGo As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRSS))
        Me.lstItems = New System.Windows.Forms.ListBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblItems = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lstItems
        '
        Me.lstItems.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstItems.BackColor = System.Drawing.SystemColors.Window
        Me.lstItems.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstItems.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstItems.IntegralHeight = False
        Me.lstItems.ItemHeight = 30
        Me.lstItems.Location = New System.Drawing.Point(10, 41)
        Me.lstItems.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstItems.Name = "lstItems"
        Me.lstItems.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstItems.Size = New System.Drawing.Size(920, 275)
        Me.lstItems.TabIndex = 1
        '
        'btnGo
        '
        Me.btnGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGo.BackColor = System.Drawing.SystemColors.Control
        Me.btnGo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGo.Enabled = False
        Me.btnGo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGo.Location = New System.Drawing.Point(817, 326)
        Me.btnGo.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGo.Size = New System.Drawing.Size(116, 52)
        Me.btnGo.TabIndex = 2
        Me.btnGo.Tag = ""
        Me.btnGo.Text = "OK"
        Me.btnGo.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(50, 50)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.TabStop = False
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblItems
        '
        Me.lblItems.BackColor = System.Drawing.SystemColors.Control
        Me.lblItems.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblItems.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblItems.Location = New System.Drawing.Point(5, 5)
        Me.lblItems.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblItems.Name = "lblItems"
        Me.lblItems.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblItems.Size = New System.Drawing.Size(126, 27)
        Me.lblItems.TabIndex = 0
        Me.lblItems.Text = "&Items"
        '
        'frmRSS
        '
        Me.AcceptButton = Me.btnGo
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(946, 392)
        Me.Controls.Add(Me.lstItems)
        Me.Controls.Add(Me.btnGo)
        Me.Controls.Add(Me.lblItems)
        Me.Controls.Add(Me.btnCancel)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmRSS"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "RSS News feed"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents lblItems As System.Windows.Forms.Label
#End Region 
End Class