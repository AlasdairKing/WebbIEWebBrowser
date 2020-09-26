<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWeb
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWeb))
        Me.webMain = New System.Windows.Forms.WebBrowser()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.tmrCheckForEscape = New System.Windows.Forms.Timer(Me.components)
        Me.tmrCheckForClosing = New System.Windows.Forms.Timer(Me.components)
        Me.tmrCheckForNavigating = New System.Windows.Forms.Timer(Me.components)
        Me.tmrCheckForNavigationComplete = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'webMain
        '
        Me.webMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.webMain.Location = New System.Drawing.Point(0, 0)
        Me.webMain.MinimumSize = New System.Drawing.Size(20, 20)
        Me.webMain.Name = "webMain"
        Me.webMain.ScriptErrorsSuppressed = True
        Me.webMain.Size = New System.Drawing.Size(806, 337)
        Me.webMain.TabIndex = 0
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(-1000, 0)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'tmrCheckForEscape
        '
        Me.tmrCheckForEscape.Enabled = True
        Me.tmrCheckForEscape.Interval = 110
        '
        'tmrCheckForClosing
        '
        Me.tmrCheckForClosing.Enabled = True
        Me.tmrCheckForClosing.Interval = 200
        '
        'tmrCheckForNavigating
        '
        Me.tmrCheckForNavigating.Enabled = True
        '
        'tmrCheckForNavigationComplete
        '
        Me.tmrCheckForNavigationComplete.Enabled = True
        Me.tmrCheckForNavigationComplete.Interval = 1000
        '
        'frmWeb
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(806, 337)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.webMain)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmWeb"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WebbIE"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents webMain As System.Windows.Forms.WebBrowser
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents tmrCheckForEscape As System.Windows.Forms.Timer
    Friend WithEvents tmrCheckForClosing As System.Windows.Forms.Timer
    Friend WithEvents tmrCheckForNavigating As Timer
    Friend WithEvents tmrCheckForNavigationComplete As Timer
End Class
