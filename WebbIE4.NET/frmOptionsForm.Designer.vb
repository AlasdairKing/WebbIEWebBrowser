<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptionsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOptionsForm))
        Me.chkUseQuickkeys = New System.Windows.Forms.CheckBox()
        Me.chkAllowPopupWindows = New System.Windows.Forms.CheckBox()
        Me.chkAllowMessages = New System.Windows.Forms.CheckBox()
        Me.chkNavigationSounds = New System.Windows.Forms.CheckBox()
        Me.chkShowImages = New System.Windows.Forms.CheckBox()
        Me.chkShowToolbarCaptions = New System.Windows.Forms.CheckBox()
        Me.chkShowToolbar = New System.Windows.Forms.CheckBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.chkIEHomepage = New System.Windows.Forms.CheckBox()
        Me.chkNumberLinks = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'chkUseQuickkeys
        '
        Me.chkUseQuickkeys.AutoSize = True
        Me.chkUseQuickkeys.Location = New System.Drawing.Point(15, 16)
        Me.chkUseQuickkeys.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkUseQuickkeys.Name = "chkUseQuickkeys"
        Me.chkUseQuickkeys.Size = New System.Drawing.Size(330, 61)
        Me.chkUseQuickkeys.TabIndex = 0
        Me.chkUseQuickkeys.Text = "Use quick keys"
        Me.chkUseQuickkeys.UseVisualStyleBackColor = True
        '
        'chkAllowPopupWindows
        '
        Me.chkAllowPopupWindows.AutoSize = True
        Me.chkAllowPopupWindows.Location = New System.Drawing.Point(15, 64)
        Me.chkAllowPopupWindows.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkAllowPopupWindows.Name = "chkAllowPopupWindows"
        Me.chkAllowPopupWindows.Size = New System.Drawing.Size(469, 61)
        Me.chkAllowPopupWindows.TabIndex = 1
        Me.chkAllowPopupWindows.Text = "Allow popup windows"
        Me.chkAllowPopupWindows.UseVisualStyleBackColor = True
        '
        'chkAllowMessages
        '
        Me.chkAllowMessages.AutoSize = True
        Me.chkAllowMessages.Location = New System.Drawing.Point(15, 112)
        Me.chkAllowMessages.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkAllowMessages.Name = "chkAllowMessages"
        Me.chkAllowMessages.Size = New System.Drawing.Size(351, 61)
        Me.chkAllowMessages.TabIndex = 2
        Me.chkAllowMessages.Text = "Allow messages"
        Me.chkAllowMessages.UseVisualStyleBackColor = True
        '
        'chkNavigationSounds
        '
        Me.chkNavigationSounds.AutoSize = True
        Me.chkNavigationSounds.Location = New System.Drawing.Point(15, 160)
        Me.chkNavigationSounds.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkNavigationSounds.Name = "chkNavigationSounds"
        Me.chkNavigationSounds.Size = New System.Drawing.Size(404, 61)
        Me.chkNavigationSounds.TabIndex = 3
        Me.chkNavigationSounds.Text = "Navigation sounds"
        Me.chkNavigationSounds.UseVisualStyleBackColor = True
        '
        'chkShowImages
        '
        Me.chkShowImages.AutoSize = True
        Me.chkShowImages.Location = New System.Drawing.Point(15, 208)
        Me.chkShowImages.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkShowImages.Name = "chkShowImages"
        Me.chkShowImages.Size = New System.Drawing.Size(302, 61)
        Me.chkShowImages.TabIndex = 4
        Me.chkShowImages.Text = "Show images"
        Me.chkShowImages.UseVisualStyleBackColor = True
        '
        'chkShowToolbarCaptions
        '
        Me.chkShowToolbarCaptions.AutoSize = True
        Me.chkShowToolbarCaptions.Location = New System.Drawing.Point(39, 304)
        Me.chkShowToolbarCaptions.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkShowToolbarCaptions.Name = "chkShowToolbarCaptions"
        Me.chkShowToolbarCaptions.Size = New System.Drawing.Size(473, 61)
        Me.chkShowToolbarCaptions.TabIndex = 6
        Me.chkShowToolbarCaptions.Text = "Show toolbar captions"
        Me.chkShowToolbarCaptions.UseVisualStyleBackColor = True
        '
        'chkShowToolbar
        '
        Me.chkShowToolbar.AutoSize = True
        Me.chkShowToolbar.Location = New System.Drawing.Point(15, 256)
        Me.chkShowToolbar.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkShowToolbar.Name = "chkShowToolbar"
        Me.chkShowToolbar.Size = New System.Drawing.Size(305, 61)
        Me.chkShowToolbar.TabIndex = 5
        Me.chkShowToolbar.Text = "Show toolbar"
        Me.chkShowToolbar.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOK.Location = New System.Drawing.Point(394, 28)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(150, 53)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'chkIEHomepage
        '
        Me.chkIEHomepage.AutoSize = True
        Me.chkIEHomepage.Location = New System.Drawing.Point(15, 352)
        Me.chkIEHomepage.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkIEHomepage.Name = "chkIEHomepage"
        Me.chkIEHomepage.Size = New System.Drawing.Size(660, 61)
        Me.chkIEHomepage.TabIndex = 7
        Me.chkIEHomepage.Text = "Use Internet Explorer homepage"
        Me.chkIEHomepage.UseVisualStyleBackColor = True
        '
        'chkNumberLinks
        '
        Me.chkNumberLinks.AutoSize = True
        Me.chkNumberLinks.Location = New System.Drawing.Point(15, 395)
        Me.chkNumberLinks.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.chkNumberLinks.Name = "chkNumberLinks"
        Me.chkNumberLinks.Size = New System.Drawing.Size(305, 61)
        Me.chkNumberLinks.TabIndex = 9
        Me.chkNumberLinks.Text = "Number links"
        Me.chkNumberLinks.UseVisualStyleBackColor = True
        '
        'frmOptionsForm
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(23.0!, 57.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnOK
        Me.ClientSize = New System.Drawing.Size(569, 445)
        Me.Controls.Add(Me.chkNumberLinks)
        Me.Controls.Add(Me.chkIEHomepage)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.chkShowToolbar)
        Me.Controls.Add(Me.chkShowToolbarCaptions)
        Me.Controls.Add(Me.chkShowImages)
        Me.Controls.Add(Me.chkNavigationSounds)
        Me.Controls.Add(Me.chkAllowMessages)
        Me.Controls.Add(Me.chkAllowPopupWindows)
        Me.Controls.Add(Me.chkUseQuickkeys)
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.Name = "frmOptionsForm"
        Me.Text = "Options"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkUseQuickkeys As System.Windows.Forms.CheckBox
    Friend WithEvents chkAllowPopupWindows As System.Windows.Forms.CheckBox
    Friend WithEvents chkAllowMessages As System.Windows.Forms.CheckBox
    Friend WithEvents chkNavigationSounds As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowImages As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowToolbarCaptions As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowToolbar As System.Windows.Forms.CheckBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents chkIEHomepage As System.Windows.Forms.CheckBox
    Friend WithEvents chkNumberLinks As System.Windows.Forms.CheckBox
End Class
