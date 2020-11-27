Public Class frmMain

    Private Contents As System.Collections.Generic.Dictionary(Of Integer, mshtml.IHTMLDOMNode)

    Private Frames As System.Collections.Generic.List(Of Object)

    Private Sub txtAddress_KeyDown(sender As Object, e As KeyEventArgs) Handles txtAddress.KeyDown
        If e.KeyCode = Keys.Return Then
            Call webMain.Navigate(txtAddress.Text)
            e.Handled = True
        End If
    End Sub

    Private Sub txtAddress_TextChanged(sender As Object, e As EventArgs) Handles txtAddress.TextChanged

    End Sub

    Private Sub webMain_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles webMain.DocumentCompleted
        If webMain.ReadyState = WebBrowserReadyState.Complete Then
            Dim domDocument As mshtml.IHTMLDocument2 = CType(webMain.Document.DomDocument, mshtml.IHTMLDocument2)
            txtOutput.Text = Parse(CType(domDocument.body, mshtml.IHTMLDOMNode))
        Else
            Frames.Add(sender)
        End If
    End Sub


    Private Function Parse(node As mshtml.IHTMLDOMNode) As String
        Dim s As String = ""
        If node.nodeType = 1 Then ' Element
        ElseIf node.nodeType = 3 Then 'Text Node
            Return node.nodeValue
        Else
            Return ""
        End If
        Select Case node.nodeName
            Case "SCRIPT"
                Return ""
            Case "STYLE"
                Return ""
            Case "A"
                Contents.Add(node.GetHashCode(), node)
                s += Environment.NewLine & "LINK (" + node.GetHashCode().ToString() + ") "
            Case "INPUT"
                Contents.Add(node.GetHashCode(), node)
                Dim input As mshtml.IHTMLInputElement = CType(node, mshtml.IHTMLInputElement)
                Select Case input.type.ToUpperInvariant
                    Case "BUTTON"
                        s &= Environment.NewLine & "INPUT BUTTON (" + node.GetHashCode().ToString() & ") " & input.value & Environment.NewLine
                    Case "PASSWORD"
                        s &= Environment.NewLine & "PASSWORD INPUT (" + node.GetHashCode().ToString() & ") " & input.value & Environment.NewLine
                    Case "SUBMIT"
                        s &= Environment.NewLine & "SUBMIT (" + node.GetHashCode().ToString() & ") " & input.value & Environment.NewLine
                    Case "RESET"
                        s &= Environment.NewLine & "RESET (" + node.GetHashCode().ToString() & ") " & input.value & Environment.NewLine
                    Case "HIDDEN"
                        'Add nothing.
                    Case Else ' TEXT
                        s &= Environment.NewLine & "INPUT TEXTBOX (" + node.GetHashCode().ToString() & ") " & input.value & Environment.NewLine
                End Select
                Return s
            Case "IFRAME"
                Contents.Add(node.GetHashCode(), node)
                s &= Environment.NewLine & "FRAME: (" & node.GetHashCode().ToString & ") "
        End Select

        Dim children As mshtml.IHTMLDOMChildrenCollection = node.childNodes
        For i As Integer = 0 To children.length - 1
            s &= Parse(children(i))
        Next i
        Select Case node.nodeName
            Case "P", "H1", "H2", "H3", "H4", "H5", "H6"
                s &= Environment.NewLine
        End Select
        Return s
    End Function

    Private Sub webMain_DocumentTitleChanged(sender As Object, e As EventArgs) Handles webMain.DocumentTitleChanged
        Me.Text = webMain.DocumentTitle
    End Sub

    Private Sub webMain_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles webMain.Navigating
        Me.txtOutput.Clear()
        Contents = New Dictionary(Of Integer, mshtml.IHTMLDOMNode)
        Frames = New List(Of Object)
    End Sub

    Private Sub txtOutput_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles txtOutput.PreviewKeyDown

    End Sub

    Private Sub txtOutput_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtOutput.KeyPress
        e.Handled = True
    End Sub

    Private Sub txtOutput_KeyDown(sender As Object, e As KeyEventArgs) Handles txtOutput.KeyDown
        If e.KeyCode = System.Windows.Forms.Keys.KeyCode.Return Then
            Dim lineNumber As Int16 = txtOutput.GetLineFromCharIndex(txtOutput.SelectionStart)
            'System.Diagnostics.Debug.Print("Line:" & lineNumber)
            Dim line As String = txtOutput.Lines(lineNumber).Trim()
            'System.Diagnostics.Debug.Print(txtOutput.Lines(lineNumber - 1))
            'System.Diagnostics.Debug.Print(txtOutput.Lines(lineNumber + 1))

            'Dim s As String = ""
            'For i As Integer = 0 To txtOutput.Lines.Length - 1
            ' s = s & i & "=" & Trim(txtOutput.Lines(i)) & vbNewLine
            ' Next i
            'MessageBox.Show(Me, s)

            If line.StartsWith("LINK") Then
                'Link to follow!
                Dim chars() As Char = {"(", ")"}
                Dim bits() As String = line.Split(chars)
                Dim indexS As String = bits(1)
                Dim node As mshtml.IHTMLDOMNode = Contents.Item(Integer.Parse(indexS))
                If node.nodeName = "A" Then
                    Dim a As mshtml.IHTMLElement = CType(node, mshtml.IHTMLElement)
                    Call a.click()
                ElseIf node.nodeName = "INPUT" Then
                End If
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub txtAddress_Enter(sender As Object, e As EventArgs) Handles txtAddress.Enter
        Call txtAddress.SelectAll()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        webMain.ScriptErrorsSuppressed = True
    End Sub
End Class
