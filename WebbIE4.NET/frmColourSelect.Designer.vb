<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmColourSelect
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
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.radWindowsDefault = New System.Windows.Forms.RadioButton()
        Me.radBlackOnWhite = New System.Windows.Forms.RadioButton()
        Me.radWhiteOnBlack = New System.Windows.Forms.RadioButton()
        Me.radYellowOnBlack = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(345, 12)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(117, 42)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(345, 60)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(117, 42)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'radWindowsDefault
        '
        Me.radWindowsDefault.AutoSize = True
        Me.radWindowsDefault.Location = New System.Drawing.Point(12, 12)
        Me.radWindowsDefault.Name = "radWindowsDefault"
        Me.radWindowsDefault.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.radWindowsDefault.Size = New System.Drawing.Size(232, 34)
        Me.radWindowsDefault.TabIndex = 2
        Me.radWindowsDefault.TabStop = True
        Me.radWindowsDefault.Text = "Use Windows default"
        Me.radWindowsDefault.UseVisualStyleBackColor = True
        '
        'radBlackOnWhite
        '
        Me.radBlackOnWhite.AutoSize = True
        Me.radBlackOnWhite.BackColor = System.Drawing.Color.White
        Me.radBlackOnWhite.ForeColor = System.Drawing.Color.Black
        Me.radBlackOnWhite.Location = New System.Drawing.Point(12, 52)
        Me.radBlackOnWhite.Name = "radBlackOnWhite"
        Me.radBlackOnWhite.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.radBlackOnWhite.Size = New System.Drawing.Size(170, 34)
        Me.radBlackOnWhite.TabIndex = 3
        Me.radBlackOnWhite.TabStop = True
        Me.radBlackOnWhite.Text = "Black on white"
        Me.radBlackOnWhite.UseVisualStyleBackColor = False
        '
        'radWhiteOnBlack
        '
        Me.radWhiteOnBlack.AutoSize = True
        Me.radWhiteOnBlack.BackColor = System.Drawing.Color.Black
        Me.radWhiteOnBlack.ForeColor = System.Drawing.Color.White
        Me.radWhiteOnBlack.Location = New System.Drawing.Point(12, 92)
        Me.radWhiteOnBlack.Name = "radWhiteOnBlack"
        Me.radWhiteOnBlack.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.radWhiteOnBlack.Size = New System.Drawing.Size(175, 34)
        Me.radWhiteOnBlack.TabIndex = 4
        Me.radWhiteOnBlack.TabStop = True
        Me.radWhiteOnBlack.Text = "White on black"
        Me.radWhiteOnBlack.UseVisualStyleBackColor = False
        '
        'radYellowOnBlack
        '
        Me.radYellowOnBlack.AutoSize = True
        Me.radYellowOnBlack.BackColor = System.Drawing.Color.Black
        Me.radYellowOnBlack.ForeColor = System.Drawing.Color.Yellow
        Me.radYellowOnBlack.Location = New System.Drawing.Point(12, 132)
        Me.radYellowOnBlack.Name = "radYellowOnBlack"
        Me.radYellowOnBlack.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.radYellowOnBlack.Size = New System.Drawing.Size(180, 34)
        Me.radYellowOnBlack.TabIndex = 5
        Me.radYellowOnBlack.TabStop = True
        Me.radYellowOnBlack.Text = "Yellow on black"
        Me.radYellowOnBlack.UseVisualStyleBackColor = False
        '
        'frmColourSelect
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 30.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(474, 185)
        Me.Controls.Add(Me.radYellowOnBlack)
        Me.Controls.Add(Me.radWhiteOnBlack)
        Me.Controls.Add(Me.radBlackOnWhite)
        Me.Controls.Add(Me.radWindowsDefault)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.Name = "frmColourSelect"
        Me.Text = "Colour"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents radWindowsDefault As System.Windows.Forms.RadioButton
    Friend WithEvents radBlackOnWhite As System.Windows.Forms.RadioButton
    Friend WithEvents radWhiteOnBlack As System.Windows.Forms.RadioButton
    Friend WithEvents radYellowOnBlack As System.Windows.Forms.RadioButton
End Class
