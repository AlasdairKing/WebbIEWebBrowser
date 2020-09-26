<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmTextView
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
	Public WithEvents mnuFileClose As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditCut As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditCopy As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditPaste As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditSelectall As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditFind As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEditFindnext As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents txtTextView As System.Windows.Forms.RichTextBox
    Public WithEvents cmdOK As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTextView))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileClose = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditSelectall = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditFindnext = New System.Windows.Forms.ToolStripMenuItem()
        Me.txtTextView = New System.Windows.Forms.RichTextBox()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.MainMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Padding = New System.Windows.Forms.Padding(12, 5, 0, 5)
        Me.MainMenu1.Size = New System.Drawing.Size(706, 39)
        Me.MainMenu1.TabIndex = 2
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileClose})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(53, 29)
        Me.mnuFile.Tag = "frmTextView.mnuFile"
        Me.mnuFile.Text = "&File"
        '
        'mnuFileClose
        '
        Me.mnuFileClose.Name = "mnuFileClose"
        Me.mnuFileClose.Size = New System.Drawing.Size(130, 30)
        Me.mnuFileClose.Tag = "frmTextView.mnuFileClose"
        Me.mnuFileClose.Text = "&Close"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuEditCut, Me.mnuEditCopy, Me.mnuEditPaste, Me.mnuEditSelectall, Me.mnuEditFind, Me.mnuEditFindnext})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(56, 29)
        Me.mnuEdit.Tag = "frmTextView.mnuEdit"
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditCut
        '
        Me.mnuEditCut.Enabled = False
        Me.mnuEditCut.Name = "mnuEditCut"
        Me.mnuEditCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuEditCut.Size = New System.Drawing.Size(305, 30)
        Me.mnuEditCut.Tag = "frmTextView.mnuEditCut"
        Me.mnuEditCut.Text = "Cu&t"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Enabled = False
        Me.mnuEditCopy.Name = "mnuEditCopy"
        Me.mnuEditCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuEditCopy.Size = New System.Drawing.Size(305, 30)
        Me.mnuEditCopy.Tag = "frmTextView.mnuEditCopy"
        Me.mnuEditCopy.Text = "&Copy"
        '
        'mnuEditPaste
        '
        Me.mnuEditPaste.Enabled = False
        Me.mnuEditPaste.Name = "mnuEditPaste"
        Me.mnuEditPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.mnuEditPaste.Size = New System.Drawing.Size(305, 30)
        Me.mnuEditPaste.Tag = "frmTextView.mnuEditPaste"
        Me.mnuEditPaste.Text = "&Paste"
        '
        'mnuEditSelectall
        '
        Me.mnuEditSelectall.Name = "mnuEditSelectall"
        Me.mnuEditSelectall.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.mnuEditSelectall.Size = New System.Drawing.Size(305, 30)
        Me.mnuEditSelectall.Tag = "frmTextView.mnuEditSelectall"
        Me.mnuEditSelectall.Text = "Select &All"
        '
        'mnuEditFind
        '
        Me.mnuEditFind.Name = "mnuEditFind"
        Me.mnuEditFind.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.mnuEditFind.Size = New System.Drawing.Size(305, 30)
        Me.mnuEditFind.Tag = "frmTextView.mnuEditFind"
        Me.mnuEditFind.Text = "&Find (on this page)"
        '
        'mnuEditFindnext
        '
        Me.mnuEditFindnext.Name = "mnuEditFindnext"
        Me.mnuEditFindnext.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.mnuEditFindnext.Size = New System.Drawing.Size(305, 30)
        Me.mnuEditFindnext.Tag = "frmTextView.mnuEditFindnext"
        Me.mnuEditFindnext.Text = "Find &Next"
        '
        'txtTextView
        '
        Me.txtTextView.BackColor = System.Drawing.SystemColors.Window
        Me.txtTextView.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTextView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtTextView.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTextView.Location = New System.Drawing.Point(0, 39)
        Me.txtTextView.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.txtTextView.MaxLength = 0
        Me.txtTextView.Name = "txtTextView"
        Me.txtTextView.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTextView.Size = New System.Drawing.Size(706, 328)
        Me.txtTextView.TabIndex = 1
        Me.txtTextView.Text = ""
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(720, 1274)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(162, 76)
        Me.cmdOK.TabIndex = 0
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'frmTextView
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdOK
        Me.ClientSize = New System.Drawing.Size(706, 367)
        Me.Controls.Add(Me.txtTextView)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(15, 57)
        Me.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.Name = "frmTextView"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Tag = "frmTextView"
        Me.Text = "WebbIE Text View"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class