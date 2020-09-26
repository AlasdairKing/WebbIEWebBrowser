Option Strict On
Option Explicit On
Friend Class frmGoogle
    Inherits System.Windows.Forms.Form

    '   This file is part of WebbIE.
    '
    '    WebbIE is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General Public License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.
    '
    '    WebbIE is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General Public License for more details.
    '
    '    You should have received a copy of the GNU General Public License
    '    along with WebbIE.  If not, see <http://www.gnu.org/licenses/>.


    Private ReadOnly mUrls(0 To MAX_NUMBER_LINKS_SUPPORTED - 1) As String
    Private mNextPageHRef As String
    Private mStartAt As Integer
    Private mXMLResultsPage As System.Text.StringBuilder
    Private Const NUMBER_RESULTS_ON_GOOGLE_PAGE As Integer = 10

    Private Sub cmdGo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGo.Click
        On Error Resume Next
        Call GoToResult()
    End Sub

    Private Sub GoToResult()
        Try
            If lstResults.SelectedIndex = lstResults.Items.Count - 1 And Len(mNextPageHRef) > 0 Then
                Call Me.mWebBrowser.Navigate(New System.Uri(mNextPageHRef))
            Else
                Call Me.Hide()
                Call frmMain.StartNavigating(mUrls(lstResults.SelectedIndex))
            End If
        Catch
        End Try
    End Sub

    Private Sub cmdSearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSearch.Click
        On Error Resume Next
        Call StartSearch()
    End Sub

    Private Sub frmGoogle_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            Me.Text = modI18N.GetText("Websearch")
            'Set focus to one of txtSearch or lstResults
            If Me.ActiveControl Is Nothing Then
                Call txtSearch.Focus()
            ElseIf Me.ActiveControl.Name = "txtSearch" Then
                Call txtSearch.Focus()
            ElseIf Me.ActiveControl.Name = "lstResults" Then
                Call lstResults.Focus()
            Else
                If Len(txtSearch.Text) = 0 Then
                    Call txtSearch.Focus()
                Else
                    Call lstResults.Focus()
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub frmGoogle_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
        Try
            If lstResults.Items.Count = 0 Then
                'no results. go to search box
                Call txtSearch.Focus()
            Else
                Call lstResults.Focus()
            End If
        Catch
        End Try
    End Sub

    Private Sub frmGoogle_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
        mWebBrowser.ScriptErrorsSuppressed = True
        mWebBrowser.Left = -100 - mWebBrowser.Width
        mWebBrowser.TabStop = False
        Me.lblSearch.TabIndex = 0
        Me.txtSearch.TabIndex = 1
        Me.lblResults.TabIndex = 2
        Me.lstResults.TabIndex = 3
        Me.lblSearch.TabIndex = 0
        Me.txtSearch.TabIndex = 1
        Me.lblResults.TabIndex = 2
        Me.lstResults.TabIndex = 3
    End Sub

    Private Sub frmGoogle_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error Resume Next
        tmrNavigating.Enabled = False
    End Sub

    Private Sub lstResults_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lstResults.KeyDown
        Try
            If eventArgs.KeyCode = System.Windows.Forms.Keys.Return And (lstResults.Items.Count > 0) Then
                Call GoToResult()
            ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Up And lstResults.SelectedIndex = 0 Then
                Call Beep()
            ElseIf eventArgs.KeyCode = System.Windows.Forms.Keys.Down And lstResults.SelectedIndex = lstResults.Items.Count - 1 Then
                Call Beep()
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrNavigating_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrNavigating.Tick
        '3.8.4 Prevent ticking if navigation has stopped for any reason.
        Try
            If mWebBrowser.IsBusy Then
                Call Debug.Print("Doing IsBusy")
                Call PlayWorkingSound()
            Else
                tmrNavigating.Enabled = False
            End If
        Catch
        End Try
    End Sub

    Private Sub txtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.TextChanged
        On Error Resume Next
        cmdSearch.Enabled = Len(txtSearch.Text) > 0
    End Sub

    Private Sub txtSearch_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.Enter
        On Error Resume Next
        txtSearch.SelectionStart = 0
        txtSearch.SelectionLength = Len(txtSearch.Text)
        cmdSearch.Enabled = Len(txtSearch.Text) > 0
    End Sub

    Private Sub txtSearch_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        On Error Resume Next
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then
            eventArgs.Handled = True
            Call StartSearch()
        End If
    End Sub

    Private Sub StartSearch()
        On Error Resume Next
        Dim searchString As String

        mStartAt = 0
        If txtSearch.Text.Trim <> "" Then
            searchString = modI18N.GetText("Search engine:")
            'Make sure default is specified.
            If searchString = "Search engine:" Then searchString = "http://www.google.com/search?q="
            searchString &= txtSearch.Text.Replace(" ", "+")
            Call mWebBrowser.Navigate(New System.Uri(searchString))
            tmrNavigating.Enabled = True
        End If
    End Sub

    Private Sub mWebBrowser_DocumentCompleted(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles mWebBrowser.DocumentCompleted
        On Error Resume Next
        Dim url As String = eventArgs.Url.ToString()

        Dim doc As mshtml.HTMLDocument
        Dim nodeForName As System.Xml.XmlNode
        Dim nodeForHREF As System.Xml.XmlNode
        Dim resultsDoc As System.Xml.XmlDocument
        Dim ns As System.Xml.XmlNodeList
        Dim n As System.Xml.XmlNode
        Dim linkCount As Integer

        If mWebBrowser.Document Is Nothing Then
            'Didn't get a document
        ElseIf eventArgs.Url.ToString = "http:///" Then
            'Ignore.
        ElseIf eventArgs.Url.ToString = "" Then
            'Ignore
        ElseIf eventArgs.Url.ToString = "about:blank" Then
            'Ignore
        Else
            'OK, process results.
            Call lstResults.Items.Clear()
            mNextPageHRef = ""
            mXMLResultsPage = New System.Text.StringBuilder("<results>")
            resultsDoc = New System.Xml.XmlDocument
            doc = CType(mWebBrowser.Document.DomDocument, mshtml.HTMLDocument)
            Call ParseToXML(CType(doc.body, mshtml.IHTMLDOMNode))
            mXMLResultsPage.Append("</results>")
            Call resultsDoc.LoadXml(mXMLResultsPage.ToString)
            'Did you mean? - no, we're not going to do "did you mean" any more, since Google doesn't seem to have those.
            ns = resultsDoc.DocumentElement.SelectNodes("//H3") ' [contains(@class, "" r "") or starts-with(@class, ""r "")]") ' or substring(@class, string-length(@class)-2, 2)="" r""]")
            If ns.Count > 0 Then
                For Each n In ns
                    nodeForName = n
                    nodeForHREF = n.ParentNode
                    If nodeForName Is Nothing Or nodeForHREF Is Nothing Then
                        'Whoops! 
                    ElseIf nodeForHREF.Attributes.GetNamedItem("href") Is Nothing Then
                    Else
                        Call lstResults.Items.Add(nodeForName.InnerText)
                        mUrls(linkCount) = nodeForHREF.Attributes.GetNamedItem("href").Value
                        linkCount += 1
                    End If
                Next n
            End If
            If lstResults.Items.Count = 0 Then
                Call lstResults.Items.Add(modI18N.GetText("No results found"))
            Else
                mStartAt += NUMBER_RESULTS_ON_GOOGLE_PAGE
                'Work out next-page link.
                Dim nextLink As System.Uri = New Uri(url)
                Dim args As String() = nextLink.Query.Split(Chr(38)) ' &
                For Each arg In args
                    If arg.StartsWith("start=") Then
                        arg = "start=" & mStartAt
                        Exit For
                    ElseIf arg.StartsWith("?start=") Then
                        arg = "?start=" & mStartAt
                        Exit For
                    End If
                Next arg
                mNextPageHRef = nextLink.GetLeftPart(UriPartial.Path) & String.Join("&", args)
                mNextPageHRef = url & "&start=" & mStartAt
                Call lstResults.Items.Add(modI18N.GetText("Next page of results"))
                'Update UI
                cmdGo.Enabled = True
            End If
            lstResults.SelectedIndex = 0
            If Not Me.Visible Then
                Call Me.Show()
            End If
            Call lstResults.Focus()
            ' Added API call for focus: fixes "if you hit next page of results, you get the focus lost in the
            '   webbrowser' bug. May 2016.
            Call NativeMethods.SetFocus(Me.lstResults.Handle)
            tmrNavigating.Enabled = False
        End If
    End Sub

    Private Sub ParseToXML(ByRef n As mshtml.IHTMLDOMNode)
        On Error Resume Next
        Dim childNodes As mshtml.IHTMLDOMChildrenCollection
        Dim childNode As mshtml.IHTMLDOMNode
        Dim childAttributes As mshtml.IHTMLAttributeCollection
        Dim childAttribute As mshtml.IHTMLDOMAttribute
        Dim doAttributes As Boolean
        Dim doChildren As Boolean
        Dim doClose As Boolean
        Dim attNodeName As String
        Dim nodeName As String = "NOT_A_NODE"
        Dim nodeType As Long

        nodeType = n.nodeType
        If nodeType = ELEMENT_NODE Then
            nodeName = n.nodeName
            Select Case nodeName
                Case "A", "LI", "H1", "H2", "H3", "H4"
                    mXMLResultsPage.Append("<" & nodeName & " ")
                    doAttributes = True
                    doChildren = True
                    doClose = True
                Case "SCRIPT", "HEAD", "RUBY", "RP", "RT", "BR", "NOSCRIPT", "AUDIO", "CANVAS", "OBJECT", "IMG", "VIDEO", "EMBED", "IFRAME",
                     "FRAME", "FRAMESET", "MAP", "NOFRAMES", "PARAM", "SOURCE", "STYLE"
                    doChildren = False
                Case Else
                    doChildren = True
            End Select
        ElseIf nodeType = TEXT_NODE Then
            If n.nodeValue Is Nothing Then
            Else
                mXMLResultsPage.Append(n.nodeValue.ToString.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;"))
            End If
        End If

        If doAttributes Then
            childAttributes = CType(n.attributes, mshtml.IHTMLAttributeCollection)
            If childAttributes Is Nothing Then
            Else
                For Each childAttribute In childAttributes
                    attNodeName = childAttribute.nodeName
                    If attNodeName = "class" Then
                        If childAttribute.nodeValue Is Nothing Then
                        Else
                            mXMLResultsPage.Append(" class=""" & childAttribute.nodeValue.ToString & """")
                        End If
                    ElseIf attNodeName = "id" Then
                        If childAttribute.nodeValue Is Nothing Then
                        Else
                            mXMLResultsPage.Append(" id=""" & childAttribute.nodeValue.ToString & """")
                        End If
                    ElseIf attNodeName = "href" Then
                        If childAttribute.nodeValue Is Nothing Then
                        Else
                            mXMLResultsPage.Append(" href=""" & childAttribute.nodeValue.ToString.Replace("&", "&amp;") & """")
                        End If
                    End If
                Next childAttribute
            End If
            mXMLResultsPage.Append(">")
        End If

        If doChildren And n.hasChildNodes Then
            childNodes = CType(n.childNodes, mshtml.IHTMLDOMChildrenCollection)
            For Each childNode In childNodes
                Call ParseToXML(childNode)
            Next childNode
        End If

        If doClose Then
            Call mXMLResultsPage.Append("</" & nodeName & ">")
        End If
    End Sub

    Public Sub StartSearch(ByRef terms As String)
        On Error Resume Next
        'starts a search from an external call

        txtSearch.Text = terms
        Call Me.Show(frmMain)
        Call StartSearch()
    End Sub

    Private Sub frmGoogle_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Call Me.Hide()
        End If
    End Sub

    Private Sub frmGoogle_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        On Error Resume Next
        If e.KeyChar = Microsoft.VisualBasic.ChrW(System.Windows.Forms.Keys.Escape) Then
            Call Me.Hide()
        End If
    End Sub

    Private Sub mWebBrowser_GotFocus(sender As Object, e As System.EventArgs) Handles mWebBrowser.GotFocus
        On Error Resume Next
        'You might add some logic like "put focus on the text search box if there are no search results" but
        'remember that WebBrowser objects pull focus all the time, so if you do that while the code is parsing
        'the Google results then you will end up with the focus on the text search box instead of the results
        'list.
        Call lstResults.Focus()
    End Sub

    Private Sub tmrControlFocus_Tick(sender As System.Object, e As System.EventArgs) Handles tmrControlFocus.Tick
        On Error Resume Next
        Dim setFocus As Boolean
        If Me.Visible Then
            If Me.ActiveControl Is Nothing Then
                setFocus = True
            ElseIf Me.ActiveControl.Name = "mWebBrowser" Then
                setFocus = True
            End If
            If setFocus Then
                Call Me.Focus()
                If lstResults.Items.Count > 0 Then
                    Call lstResults.Focus()
                Else
                    Call txtSearch.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub lstResults_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstResults.SelectedIndexChanged

    End Sub

    Private Sub txtSearch_KeyUp(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyUp
        On Error Resume Next
        If e.KeyCode = System.Windows.Forms.Keys.Return Then
            e.Handled = True
        End If
    End Sub
End Class