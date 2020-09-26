<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmGoogle
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
    Public WithEvents lstResults As System.Windows.Forms.ListBox
    Public WithEvents txtSearch As System.Windows.Forms.TextBox
    Public WithEvents mWebBrowser As System.Windows.Forms.WebBrowser
    Public WithEvents tmrNavigating As System.Windows.Forms.Timer
    Public WithEvents cmdSearch As System.Windows.Forms.Button
    Public WithEvents cmdGo As System.Windows.Forms.Button
    Public WithEvents lblResults As System.Windows.Forms.Label
    Public WithEvents lblSearch As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGoogle))
        Me.lstResults = New System.Windows.Forms.ListBox()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.mWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.tmrNavigating = New System.Windows.Forms.Timer(Me.components)
        Me.cmdSearch = New System.Windows.Forms.Button()
        Me.cmdGo = New System.Windows.Forms.Button()
        Me.lblResults = New System.Windows.Forms.Label()
        Me.lblSearch = New System.Windows.Forms.Label()
        Me.tmrControlFocus = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'lstResults
        '
        Me.lstResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstResults.BackColor = System.Drawing.SystemColors.Window
        Me.lstResults.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstResults.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstResults.IntegralHeight = False
        Me.lstResults.ItemHeight = 30
        Me.lstResults.Location = New System.Drawing.Point(10, 120)
        Me.lstResults.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstResults.Size = New System.Drawing.Size(753, 371)
        Me.lstResults.TabIndex = 9
        '
        'txtSearch
        '
        Me.txtSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearch.BackColor = System.Drawing.SystemColors.Window
        Me.txtSearch.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSearch.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSearch.Location = New System.Drawing.Point(10, 37)
        Me.txtSearch.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.txtSearch.MaxLength = 0
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSearch.Size = New System.Drawing.Size(590, 35)
        Me.txtSearch.TabIndex = 8
        '
        'mWebBrowser
        '
        Me.mWebBrowser.Location = New System.Drawing.Point(419, 165)
        Me.mWebBrowser.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.mWebBrowser.Name = "mWebBrowser"
        Me.mWebBrowser.Size = New System.Drawing.Size(153, 167)
        Me.mWebBrowser.TabIndex = 7
        Me.mWebBrowser.TabStop = False
        '
        'tmrNavigating
        '
        Me.tmrNavigating.Enabled = True
        Me.tmrNavigating.Interval = 300
        '
        'cmdSearch
        '
        Me.cmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSearch.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSearch.Enabled = False
        Me.cmdSearch.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSearch.Location = New System.Drawing.Point(624, 24)
        Me.cmdSearch.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSearch.Size = New System.Drawing.Size(139, 62)
        Me.cmdSearch.TabIndex = 2
        Me.cmdSearch.TabStop = False
        Me.cmdSearch.Tag = "frmGoogle.cmdSearch"
        Me.cmdSearch.Text = "S&earch"
        Me.cmdSearch.UseVisualStyleBackColor = False
        '
        'cmdGo
        '
        Me.cmdGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdGo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGo.Enabled = False
        Me.cmdGo.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGo.Location = New System.Drawing.Point(624, 503)
        Me.cmdGo.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.cmdGo.Name = "cmdGo"
        Me.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGo.Size = New System.Drawing.Size(139, 43)
        Me.cmdGo.TabIndex = 3
        Me.cmdGo.TabStop = False
        Me.cmdGo.Tag = "frmGoogle.cmdGo"
        Me.cmdGo.Text = "&Go"
        Me.cmdGo.UseVisualStyleBackColor = False
        '
        'lblResults
        '
        Me.lblResults.AutoSize = True
        Me.lblResults.BackColor = System.Drawing.SystemColors.Control
        Me.lblResults.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblResults.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblResults.Location = New System.Drawing.Point(5, 84)
        Me.lblResults.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblResults.Name = "lblResults"
        Me.lblResults.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblResults.Size = New System.Drawing.Size(79, 30)
        Me.lblResults.TabIndex = 6
        Me.lblResults.Tag = "frmGoogle.lblResults"
        Me.lblResults.Text = "&Results"
        '
        'lblSearch
        '
        Me.lblSearch.AutoSize = True
        Me.lblSearch.BackColor = System.Drawing.SystemColors.Control
        Me.lblSearch.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSearch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSearch.Location = New System.Drawing.Point(5, 5)
        Me.lblSearch.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblSearch.Name = "lblSearch"
        Me.lblSearch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSearch.Size = New System.Drawing.Size(75, 30)
        Me.lblSearch.TabIndex = 5
        Me.lblSearch.Tag = "frmGoogle.lblSearch"
        Me.lblSearch.Text = "&Search"
        '
        'tmrControlFocus
        '
        Me.tmrControlFocus.Enabled = True
        Me.tmrControlFocus.Interval = 200
        '
        'frmGoogle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(784, 561)
        Me.Controls.Add(Me.lstResults)
        Me.Controls.Add(Me.txtSearch)
        Me.Controls.Add(Me.mWebBrowser)
        Me.Controls.Add(Me.cmdSearch)
        Me.Controls.Add(Me.cmdGo)
        Me.Controls.Add(Me.lblResults)
        Me.Controls.Add(Me.lblSearch)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(4, 10)
        Me.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmGoogle"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Tag = "frmGoogle"
        Me.Text = "Google Results"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tmrControlFocus As System.Windows.Forms.Timer
#End Region 
End Class