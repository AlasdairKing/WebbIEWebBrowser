<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMediaControl
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMediaControl))
        Me.btnPlay = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.tmrMedia = New System.Windows.Forms.Timer(Me.components)
        Me.lblMedia = New System.Windows.Forms.Label()
        Me.trbMedia = New System.Windows.Forms.TrackBar()
        Me.staMedia = New System.Windows.Forms.StatusStrip()
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnOpen = New System.Windows.Forms.Button()
        CType(Me.trbMedia, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.staMedia.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPlay
        '
        Me.btnPlay.Location = New System.Drawing.Point(20, 105)
        Me.btnPlay.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.btnPlay.Name = "btnPlay"
        Me.btnPlay.Size = New System.Drawing.Size(150, 53)
        Me.btnPlay.TabIndex = 0
        Me.btnPlay.Text = "Play"
        Me.btnPlay.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(182, 105)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(150, 53)
        Me.btnStop.TabIndex = 1
        Me.btnStop.Text = "Pause"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'tmrMedia
        '
        Me.tmrMedia.Enabled = True
        Me.tmrMedia.Interval = 400
        '
        'lblMedia
        '
        Me.lblMedia.AutoSize = True
        Me.lblMedia.Location = New System.Drawing.Point(15, 9)
        Me.lblMedia.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblMedia.Name = "lblMedia"
        Me.lblMedia.Size = New System.Drawing.Size(109, 45)
        Me.lblMedia.TabIndex = 2
        Me.lblMedia.Text = "Media"
        '
        'trbMedia
        '
        Me.trbMedia.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.trbMedia.Location = New System.Drawing.Point(15, 46)
        Me.trbMedia.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.trbMedia.Name = "trbMedia"
        Me.trbMedia.Size = New System.Drawing.Size(594, 69)
        Me.trbMedia.TabIndex = 3
        Me.trbMedia.TickStyle = System.Windows.Forms.TickStyle.None
        '
        'staMedia
        '
        Me.staMedia.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus})
        Me.staMedia.Location = New System.Drawing.Point(0, 181)
        Me.staMedia.Name = "staMedia"
        Me.staMedia.Size = New System.Drawing.Size(624, 22)
        Me.staMedia.TabIndex = 4
        Me.staMedia.Text = "StatusStrip1"
        '
        'lblStatus
        '
        Me.lblStatus.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.lblStatus.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 17)
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(344, 105)
        Me.btnOpen.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(150, 53)
        Me.btnOpen.TabIndex = 5
        Me.btnOpen.Text = "&Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'frmMediaControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(18.0!, 45.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(624, 203)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.staMedia)
        Me.Controls.Add(Me.trbMedia)
        Me.Controls.Add(Me.lblMedia)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnPlay)
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.Name = "frmMediaControl"
        Me.Text = "Media Control"
        CType(Me.trbMedia, System.ComponentModel.ISupportInitialize).EndInit()
        Me.staMedia.ResumeLayout(False)
        Me.staMedia.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPlay As System.Windows.Forms.Button
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents tmrMedia As System.Windows.Forms.Timer
    Friend WithEvents lblMedia As System.Windows.Forms.Label
    Friend WithEvents trbMedia As System.Windows.Forms.TrackBar
    Friend WithEvents staMedia As System.Windows.Forms.StatusStrip
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnOpen As System.Windows.Forms.Button
End Class
