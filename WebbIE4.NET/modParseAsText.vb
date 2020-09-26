Option Strict On
Option Explicit On
Module modParseAsText
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
	
	
	Private gstrNewline As String
    Private gHadSomeNonLinkTextOnThePage As Boolean

    Public Function ParseDocs() As String
        Try
            Dim got() As String
            Dim keep() As Boolean

            modParseAsText.gHadSomeNonLinkTextOnThePage = False
            modParseAsText.gstrNewline = ""

            Try
                ParseDocs = modI18N.GetText("WEBPAGE:") & " " & modGlobals.gWebHost.webMain.DocumentTitle
                ParseDocs = ParseDocs & vbNewLine
            Catch
                ParseDocs = ""
            End Try

            Dim content As String = ""
            Try
                Dim doc As mshtml.HTMLDocument = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.HTMLDocument)
                content = ParseNodeText(CType(doc.documentElement, mshtml.IHTMLDOMNode))
            Catch
                content = ""
            End Try
            'now truncate ParseDocs to remove blank lines and links that aren't part
            'of text.
            'Unless, of course, ParseDocs didn't work.
            If content = "" Then
            Else
                ParseDocs = ParseDocs & content
                'First split into array by line
                got = Split(ParseDocs, vbNewLine)
                ReDim keep(got.Length)
                'Now process through array deciding what lines to keep
                For i = 0 To got.Length - 1
                    If got(i).Replace(" ", "").Trim().Length > 0 Then keep(i) = True
                Next i
                'Remove crap at top of page by removing lines with no ANSI text.
                For i As Integer = 0 To got.Length - 1
                    Dim found As String = got(i).Trim()
                    Dim lineOK As Boolean = False
                    For j As Integer = 0 To found.Length - 1
                        Dim charValue As Integer = AscW(found.Substring(j, 1))
                        'Debug.Print "Checking " & charValue
                        If ((charValue > 47 And charValue < 58) Or (charValue > 64 And charValue < 91) Or charValue > 96) And charValue <> 124 And charValue <> 183 Then
                            'a real character, OK: can keep line.
                            lineOK = True
                            Exit For
                        Else
                            'Debug.Print "Fake:[" & charValue & "]"
                        End If
                    Next j
                    If lineOK Then
                        'we've found a valid line, so quit the remove crap loop
                        'No, on second thoughts: strip out all the crap lines.
                        'Exit For
                    Else
                        'line is crap
                        keep(i) = False
                    End If
                Next i
                'now take out the remains of navigation links at the top of the page by taking
                'out single-word lines
                For i = 0 To got.Length - 1
                    If keep(i) Then
                        'okay, this line contains valid stuff: is it more than one word long?
                        Dim words() As String = got(i).Split(" ".ToCharArray())
                        If words.Length > 0 Then
                            'okay, we've found the start of the page content.
                            Exit For
                        Else
                            'nope, only one word: not the start of the page
                            'content yet.
                            keep(i) = False
                        End If
                    End If
                Next i
                ParseDocs = ""
                For i = 0 To got.Length - 1
                    If keep(i) Then ParseDocs = ParseDocs & got(i) & vbNewLine
                Next i
                'take out errant spaces
                ParseDocs = Replace(ParseDocs, " ,", ",")
                ParseDocs = Replace(ParseDocs, " .", ".")
            End If
        Catch
            Return ""
        End Try
    End Function


    Public Function ParseNodeText(ByRef node As mshtml.IHTMLDOMNode) As String
        Try
            'processes a document html node and all of its children and siblings (by calling itself.
            'Outputs results to output, which it returns
            ' gstrnewline is either Empty or vbNewline, and should be set accordingly


            Dim tagname As String ' the HTML node name
            'Dim label As String
            Dim output As String = "" ' the result of the parsing
            Dim nodeIterator As mshtml.IHTMLDOMNode ' used to iterate through any further node collections
            'Dim nodeIterator2 As mshtml.IHTMLDOMNode ' used to iterate again
            Dim i As Integer ' counter
            Dim trimmedText As String ' used to strip spare spaces out of text
            'Dim revText As String ' reversed text
            'Dim parentnode As mshtml.IHTMLElement
            Dim childNodes As mshtml.IHTMLDOMChildrenCollection
            Dim element2 As mshtml.IHTMLElement2
            Dim element As mshtml.IHTMLElement
            Dim direction As String

            tagname = node.nodeName
            'attempt to speed things up
            'Debug.Print "Nodetype:" & node.nodeType
            If node.nodeType = TEXT_NODE Then
                'get rid of whitespace at end/start of line, INCLUDING non-breaking whitespace (Unicode
                '   value 160, NBSP)
                trimmedText = node.nodeValue.ToString.Replace(Chr(160), " ").Replace(vbTab, " ").Trim
                output = Replace(trimmedText, vbCr, vbNewLine) & " "
                If trimmedText.Length > 25 Then gHadSomeNonLinkTextOnThePage = True
                gstrNewline = vbNewLine
            ElseIf node.nodeType = ELEMENT_NODE Then
                element = CType(node, mshtml.IHTMLElement)
                element2 = CType(node, mshtml.IHTMLElement2)
                If element2.currentStyle.display = "none" Then  'check displayed
                    'don't display this element at all! And don't do kids!
                    'TODO: really? Is this consistent with our main parsing?
                ElseIf tagname = "A" Then
                    'Don't do link content until we have some non-link text
                    If gHadSomeNonLinkTextOnThePage Then
                        'Sadly, I don't understand this check so commenting out because it makes the
                        'headline on http://www.voxeu.org/columns/reads disappear because they are
                        'A elements inside H2 elements, so they don't render.
                        ''okay, this is a link: does it make up all the content of the containing
                        ''element, and if so, don't display it.
                        'If Len(element.parentElement.innerText) <= Len(element.innerText) + 5 Then
                        '    'don't display! Probably a link in a list
                        '    'Debug.Print "Nope"
                        'Else
                        '    'okay, go ahead and display
                        'End If
                        childNodes = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                        For Each nodeIterator In childNodes
                            output = output & ParseNodeText(nodeIterator)
                        Next nodeIterator
                    Else
                        'We're processing through the page from the top but we still haven't set the 
                        'gHadSomeNonLinkTextOnThePage flag because, well, we haven't. Until that is
                        'set we don't show links.
                    End If
                ElseIf tagname = "IMG" Then
                    'Don't show images
                ElseIf tagname = "SELECT" Then
                    'don't show select
                ElseIf tagname = "INPUT" Then
                    'Don't show input
                ElseIf tagname = "BUTTON" Then
                    'Don't show button
                ElseIf tagname = "TEXTAREA" Then
                    'Don't show textarea
                ElseIf tagname = "HEAD" Then
                    'Don't do HEAD
                ElseIf tagname = "OL" Then
                    'an ordered list
                    i = 1
                    If node.hasChildNodes Then
                        childNodes = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                        For Each nodeIterator In childNodes
                            output = output & i & " "
                            gstrNewline = vbNewLine
                            output = output & ParseNodeText(nodeIterator)
                            i = i + 1
                        Next nodeIterator
                    End If
                ElseIf tagname = "HR" Then
                    'a horizontal rule: don't think I'll do this.
                    '        Set frmDummy.Font = frmMain.txtText.Font
                    '        For i = 1 To (frmMain.txtText.width / frmDummy.TextWidth("_") - 4)
                    '            output = output & "_"
                    '        Next i
                    '        output = output & vbNewLine
                    '        gstrNewline = ""
                ElseIf tagname = "AREA" Then
                    'Don't show
                ElseIf tagname = "BLOCKQUOTE" Then
                    'a quotation
                    output = output & gstrNewline & modI18N.GetText("[quotation marks]")
                    gstrNewline = vbNewLine
                    If node.hasChildNodes Then
                        childNodes = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                        For Each nodeIterator In childNodes
                            output = output & ParseNodeText(nodeIterator)
                        Next nodeIterator
                    End If
                    output = RTrim(output)
                    Try
                        While Right(output, 2) = vbNewLine
                            output = Left(output, Len(output) - 2)
                        End While
                    Catch
                    End Try

                    output = output & modI18N.GetText("[quotation marks]") & vbNewLine
                    gstrNewline = ""
                ElseIf tagname = "OBJECT" Or tagname = "APPLET" Then
                        'Don't show objects
                ElseIf tagname = "STYLE" Then
                        'content for browser - not intended for user to see
                ElseIf tagname = "DEL" Then
                        'deleted content - don't do any of it or it's child nodes!
                ElseIf tagname = "LABEL" Then
                        'Don't show labels
                ElseIf tagname = "IFRAME" Then
                        'internal frame - used for non-frame-supporting browsers - WebbIE does support frames,
                        'so don't show any content (see HTML4.0 spec)
                ElseIf tagname = "BDO" Then
                        'set-direction text - right-to-left or left-to-right
                        If element.getAttribute("dir") Is Nothing Then
                            direction = "ltr"
                        Else
                            direction = element.getAttribute("dir").ToString
                        End If
                        If direction = "ltr" Then
                            '"normal" left-to-right direction - process as "normal" text
                            If node.hasChildNodes Then
                                childNodes = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                                For Each nodeIterator In childNodes
                                    output = output & ParseNodeText(nodeIterator)
                                Next nodeIterator
                            End If
                        Else
                            'reversed right-to-left direction
                            Dim rtlContent As String = ""
                            If node.hasChildNodes Then
                                childNodes = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                                For Each nodeIterator In childNodes
                                    rtlContent = rtlContent & ParseNodeText(nodeIterator)
                                Next nodeIterator
                            End If
                            i = Len(rtlContent)
                            While i > 0
                                If Mid(rtlContent, i, 1) = vbLf Then
                                    'we've found a newline (CR + LF)
                                    output = output & vbNewLine
                                    gstrNewline = ""
                                    i = i - 2
                                Else
                                    'normal character - add to output
                                    output = output & Mid(rtlContent, i, 1)
                                    gstrNewline = vbNewLine
                                    i = i - 1
                                End If
                            End While
                        End If
                        'ElseIf tagname = "FORM" Then
                        'don't do form
                        'No, not doing forms doesn't work, because discussion pages (e.g. blogs) can encapsulate all their content in FORM
                        'nodes. 3.8.0 Jan 2009.
                        'don't do the sub-kids of an A node, an INPUT node, a FRAME node, an IMG node...
                Else
                        If tagname = "Q" Then
                            output = output & modI18N.GetText("[quotation marks]")
                            gstrNewline = vbNewLine
                        ElseIf tagname = "TABLE" Then
                            If element.getAttribute("summary") Is Nothing Then
                            Else
                                If element.getAttribute("summary").ToString <> "" Then
                                    output = output & gstrNewline & ID_TABLE & ": " & element.getAttribute("summary").ToString & vbNewLine
                                    gstrNewline = ""
                                End If
                            End If
                        ElseIf tagname = "H1" Then
                            output = output & gstrNewline & SECTION_MARKER_H1 & ": "
                            gstrNewline = vbNewLine
                        ElseIf tagname = "H2" Then
                            output = output & gstrNewline & SECTION_MARKER_H2 & ": "
                            gstrNewline = vbNewLine
                        ElseIf tagname = "H3" Then
                            output = output & gstrNewline & SECTION_MARKER_H3 & ": "
                            gstrNewline = vbNewLine
                        ElseIf tagname = "H4" Then
                            output = output & gstrNewline & SECTION_MARKER_H4 & ": "
                            gstrNewline = vbNewLine
                        ElseIf tagname = "H5" Then
                            output = output & gstrNewline & SECTION_MARKER_H5 & ": "
                            gstrNewline = vbNewLine
                        ElseIf tagname = "H6" Then
                            output = output & gstrNewline & SECTION_MARKER_H6 & ": "
                            gstrNewline = vbNewLine
                        End If

                        If node.hasChildNodes Then
                            childNodes = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                            For Each nodeIterator In childNodes
                                output = output & ParseNodeText(nodeIterator)
                            Next nodeIterator
                        End If
                        'UPGRADE_NOTE: Object nodeIterator may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        nodeIterator = Nothing
                        'do newline if the node is one requiring a newline
                        If (tagname = "BR" Or tagname = "P" Or tagname = "H1" Or tagname = "H2" Or tagname = "H3" Or tagname = "H4" Or tagname = "H5" Or tagname = "H6" Or tagname = "ADDRESS" Or tagname = "CENTER" Or tagname = "PRE" Or tagname = "TABLE" Or tagname = "TR" Or tagname = "CAPTION" Or tagname = "DIV") And gstrNewline = vbNewLine Then
                            output = output & vbNewLine
                            gstrNewline = ""
                        ElseIf tagname = "TD" Or tagname = "TH" Then
                            output = output & " "
                            gstrNewline = vbNewLine
                        ElseIf tagname = "Q" Then
                            If Right(output, 1) = " " Then output = Left(output, Len(output) - 1)
                            output = output & modI18N.GetText("[quotation marks]")
                            gstrNewline = vbNewLine
                        End If
                End If
            End If
            Return output
        Catch
            Return ""
        End Try
    End Function
End Module