Public Enum ColourScheme
    WindowsDefault = 0
    BlackOnWhite
    WhiteOnBlack
    YellowOnBlack
End Enum

Public Class frmColourSelect

    Public Sub SetCurrentSelection(selectedColourScheme As Integer)
        Dim cs As ColourScheme = CType(selectedColourScheme, ColourScheme)
        Select Case cs
            Case colourScheme.BlackOnWhite
                Me.radBlackOnWhite.Checked = True
            Case colourScheme.WhiteOnBlack
                Me.radWhiteOnBlack.Checked = True
            Case colourScheme.YellowOnBlack
                Me.radYellowOnBlack.Checked = True
            Case Else
                Me.radWindowsDefault.Checked = True
        End Select
    End Sub


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Call Me.Close()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Call Me.Close()
    End Sub

    Public Property SelectedColourScheme As ColourScheme
        Get
            If Me.radBlackOnWhite.Checked Then
                Return ColourScheme.BlackOnWhite
            ElseIf Me.radYellowOnBlack.Checked Then
                Return ColourScheme.YellowOnBlack
            ElseIf Me.radWhiteOnBlack.Checked Then
                Return ColourScheme.WhiteOnBlack
            Else
                Return ColourScheme.WindowsDefault
            End If
        End Get
        Set(value As ColourScheme)
            Select Case value
                Case ColourScheme.BlackOnWhite
                    Me.radBlackOnWhite.Checked = True
                Case ColourScheme.WhiteOnBlack
                    Me.radWhiteOnBlack.Checked = True
                Case ColourScheme.YellowOnBlack
                    Me.radYellowOnBlack.Checked = True
                Case Else
                    Me.radWindowsDefault.Checked = True
            End Select
        End Set
    End Property

    Public Shared Sub SetColourScheme(ByRef f As Form, cs As ColourScheme)
        Dim fc As Color
        Dim bc As Color
        Select Case cs
            Case ColourScheme.BlackOnWhite
                fc = Color.Black
                bc = Color.White
            Case ColourScheme.WhiteOnBlack
                fc = Color.White
                bc = Color.Black
            Case ColourScheme.YellowOnBlack
                fc = Color.Yellow
                bc = Color.Black
            Case Else
                fc = System.Drawing.Color.FromName(System.Drawing.KnownColor.WindowText.ToString())
                bc = System.Drawing.Color.FromName(System.Drawing.KnownColor.Window.ToString())

        End Select

        For Each c As Control In f.Controls
            Call SetColourSchemeRecurse(c, fc, bc)
        Next c
    End Sub

    Private Shared Sub SetColourSchemeControl(ByRef c As Control, fc As Color, bc As Color)
        Dim apply As Boolean = False
        Select Case c.GetType().ToString
            Case "System.Windows.Forms.ComboBox"
                apply = True
            Case "System.Windows.Forms.RichTextBox"
                apply = True
            Case "System.Windows.Forms.TreeView"
                apply = True
            Case "System.Windows.Forms.ListBox"
                apply = True
            Case "System.Windows.Forms.Button"
                apply = False
            Case "System.Windows.Forms.RadioButton"
                apply = False
            Case "System.Windows.Forms.Label"
                apply = False
            Case "System.Windows.Forms.TextBox"
                apply = True
            Case "System.Windows.Forms.StatusStrip"
                apply = False
            Case "System.Windows.Forms.ToolStrip"
                apply = False
            Case "System.Windows.Forms.Panel"
                apply = False
            Case "System.Windows.Forms.MenuStrip"
                apply = False
            Case Else
                Debug.Print("Type:" & c.GetType().ToString)
        End Select
        If apply Then
            c.ForeColor = fc
            c.BackColor = bc
        End If
    End Sub

    Private Shared Sub SetColourSchemeRecurse(ByRef parentControl As Control, fc As Color, bc As Color)
        Call SetColourSchemeControl(parentControl, fc, bc)
        For Each c As Control In parentControl.Controls
            Call SetColourSchemeControl(c, fc, bc)
            If c.HasChildren Then
                For Each child In c.Controls
                    Call SetColourSchemeRecurse(CType(child, System.Windows.Forms.Control), fc, bc)
                Next child
            End If
        Next c
    End Sub

End Class