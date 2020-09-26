Option Strict On
Option Explicit On
Module modParseA
	
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
	
	'modParseA
	'Parses an Anchor element, identifying text and image content and getting alt text where necessary
	'Also parses Button nodes, which can also contain content.
	
    Private mAnchorDatabase As New System.Collections.Generic.Dictionary(Of String, String)
    Private mToGetList As New System.Collections.Generic.Dictionary(Of String, String)
	
    Public Function Parse(ByRef anyNode As mshtml.IHTMLDOMNode) As String
        On Error Resume Next
        Parse = Trim(ParseA(anyNode))
    End Function

    Public Function ParseAnchor(ByRef anchorNode As mshtml.IHTMLDOMNode) As String
        On Error Resume Next
        ParseAnchor = ParseA(anchorNode).Trim()
    End Function

    Public Function ParseForText(ByRef node As mshtml.IHTMLDOMNode) As String
        Try
            'Takes the given DOM Node and parses it and its children for the maximum
            'well-formatted text possible, including alt attributes and titles and 
            'so on. This can be used for parsing images, labels, and anything else
            'where you want to extract "some meaningful text" because it's marked
            'up as an important blob - like link content or LABEL content.
            'It's really a generalised case of ParseA, so TODO merge them together
            'so I'm not duplicating functionality.
            'Caveat: note that it doesn't extract the CONTENTS of input elements.
            'This is so that if it's a LABEL element with an INPUT or TEXTVIEW
            'then you don't get the INPUT or TEXTVIEW content in the label!
            ParseForText = ""

            If node.nodeName = "#text" Then
                ParseForText = node.nodeValue.ToString
                'Convert whitespace to " " - including Chr(160), NBSP.
                ParseForText = ParseForText.Replace(Chr(160), " ").Replace(vbNewLine, " ").Replace(vbCr, " ").Replace(vbLf, " ")
            ElseIf node.nodeName = "IMG" Then
                Dim imgE = CType(node, mshtml.IHTMLImgElement)

                If imgE.alt.Length = 0 Then
                    'No alt tag!
                    Dim e As mshtml.IHTMLElement = CType(node, mshtml.IHTMLElement)
                    If e.title.Length = 0 Then
                        'No title either.
                    Else
                        ParseForText = " " & e.title.Trim & " "
                    End If
                Else
                    'We put the " " in because people tend to just do "text<img>" - BBC News
                    'does this.
                    ParseForText = " " & imgE.alt.Trim & " "
                End If
            End If
            If node.hasChildNodes Then
                Dim childNode As mshtml.IHTMLDOMNode
                Dim childNodes As mshtml.IHTMLDOMChildrenCollection = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                For Each childNode In childNodes
                    'Don't want to do inputs, because you use ParseForText in situations like:
                    '<label>Name: <input type='text' value='Enter name!'></label>
                    Select Case childNode.nodeName
                        Case "INPUT"
                        Case "TEXTAREA"
                        Case "BUTTON"
                        Case "SELECT"
                        Case Else
                            ParseForText = ParseForText & ParseForText(childNode)
                    End Select
                Next childNode
            End If
        Catch
            Return vbNullString
        End Try
    End Function


    Private Function ParseA(ByRef anchorNode As mshtml.IHTMLDOMNode) As String
        Try
            Dim childNode As mshtml.IHTMLDOMNode
            Dim NBSP As String = Chr(160)
            Dim text As String = ""
            Dim childNodes As mshtml.IHTMLDOMChildrenCollection
            Dim htmlE As mshtml.IHTMLElement

            If anchorNode.nodeName = "#text" Then
                If anchorNode.nodeValue Is Nothing Then
                    text = ""
                Else
                    text = anchorNode.nodeValue.ToString
                    text = Replace(text, NBSP, " ")
                    text = Replace(text, vbNewLine, "")
                    text = Replace(text, vbCr, "")
                    text = Replace(text, vbLf, "")
                    text = text & " " ' Because some sites - like FaceBook 
                    ' - stick nodes together with no rendered whitespace.
                    'Debug.Print "Text obtained via ParseA: " & ParseA
                End If
                Return text
            ElseIf anchorNode.nodeType <> ELEMENT_NODE Then
                'OK, it's a non-HTML-Element, like a comment - no good to us.
                Return ""
            Else
                'It's an HTML element.
                If anchorNode.nodeName = "IMG" Then
                    htmlE = CType(anchorNode, mshtml.IHTMLElement)
                    If htmlE.getAttribute("alt") Is Nothing Then
                        'No alt tag!
                    Else
                        'We put the " " in because people tend to just do "text<img>" - BBC News
                        'does this.
                        text = " " & htmlE.getAttribute("alt").ToString & " "
                    End If
                End If
                If text.Trim().Length = 0 Then
                    'Try aria-label
                    ' For reasons I do not understand, getAttribute("aria-label") does not return the aria
                    ' attribute (maybe the hyphen?). However, iterating through the attributes does, though
                    ' I'm going to assume it's really slow and do a check on the html first.
                    htmlE = CType(anchorNode, mshtml.IHTMLElement)
                    If htmlE.outerHTML.Contains("aria-label") Then
                        ' There is an aria-label or aria-labelledby attribute, possibly.
                        Dim ac As mshtml.IHTMLAttributeCollection = CType(anchorNode.attributes, mshtml.IHTMLAttributeCollection)
                        For i As Integer = 0 To ac.length - 1
                            Dim da As mshtml.IHTMLDOMAttribute = CType(ac.item(CType(i, Object)), mshtml.IHTMLDOMAttribute)
                            If da.nodeName.ToLowerInvariant() = "aria-label" Then
                                text = da.nodeValue.ToString()
                            ElseIf da.nodeName.ToLowerInvariant() = "aria-labelledby" Then
                                If da.nodeValue Is Nothing Then
                                Else
                                    Dim id As String = da.nodeValue.ToString()
                                    Dim el As mshtml.IHTMLElement = CType(htmlE.document, mshtml.IHTMLDocument3).getElementById(id)
                                    If el Is Nothing Then
                                    Else
                                        text = TrivialParse(el)
                                    End If
                                    If text.Trim() <> "" Then
                                        Exit For ' Prioritise aria-labelledby over aria-label, according to the WAI spec.
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                'If no text obtained: first try children.
                If text.Trim().Length = 0 Then
                    If anchorNode.hasChildNodes Then
                        childNodes = CType(anchorNode.childNodes, mshtml.IHTMLDOMChildrenCollection)
                        For Each childNode In childNodes
                            text = text & ParseA(childNode)
                        Next childNode
                    End If
                End If
                'If no text obtained: fall back to A element's title or alt attribute. Title first.
                'But might be a non-A element with some information itself, or it might be
                'an A element with a title attribute - see http://www.unco.edu/, which uses
                'links with &nbsp; as the link content, fills in the content with an image
                'as the background in CSS, and puts the text content in the title attribute
                If text.Trim().Length = 0 Then
                    Try
                        htmlE = CType(anchorNode, mshtml.IHTMLElement)
                        If htmlE.getAttribute("title") Is Nothing Then
                        Else
                            text = htmlE.getAttribute("title").ToString & " "
                        End If
                    Catch ex As InvalidCastException ' Get a COM Access security error sometimes. OK, 
                        'when you go to slashdot.org, and get javascript error messages, then turn off
                        'messages in Options, then hit refresh on slashdot.org.
                        '"Unable to cast COM object of type 'System.__ComObject' to interface type 'mshtml.IHTMLElement'. This operation failed because the QueryInterface call on the COM component for the interface with IID '{3050F1FF-98B5-11CF-BB82-00AA00BDCE0B}' failed due to the following error: No such interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE))."
                        text = ""
                    End Try
                End If
            End If
            Return text
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Parse the element and return the text contents. This is used instead of .innerText because it adds in 
    ''' a space between child text elements, which might produce extra spaces but means things like Facebook
    ''' look okay when parsed - otherwise you get "0REquests" and "3Messages" and suchlike. Alasdair, June 2013.
    ''' </summary>
    ''' <param name="element"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TrivialParse(element As mshtml.IHTMLElement) As String
        Dim node As mshtml.IHTMLDOMNode = CType(element, mshtml.IHTMLDOMNode)
        Dim contents As String = ""
        For Each childNode As mshtml.IHTMLDOMNode In CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
            contents = contents & TrivialParseNode(childNode)
        Next childNode
        Return contents
    End Function

    Private Function TrivialParseNode(node As mshtml.IHTMLDOMNode) As String
        If node.nodeName = "#text" Then
            Return node.nodeValue.ToString() & " "
        ElseIf node.nodeType = ELEMENT_NODE Then
            Dim contents As String = ""
            For Each childNode As mshtml.IHTMLDOMNode In CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                contents = contents & TrivialParseNode(childNode)
            Next childNode
            Return contents
        Else
            Return " "
        End If
    End Function

    Public Sub AddLocationTitle(ByRef url As String, ByRef title As String)
        Try
            If mAnchorDatabase.ContainsKey(url) Then
                mAnchorDatabase(url) = title
            Else
                Call mAnchorDatabase.Add(url, title)
            End If
        Catch
        End Try
    End Sub

    Public Function GetLocationTitle(ByRef url As String) As String
        Try
            If mAnchorDatabase.ContainsKey(url) Then
                GetLocationTitle = mAnchorDatabase.Item(url).ToString
            Else
                Call AddLocationToGet(url)
                GetLocationTitle = ""
            End If
        Catch
            Return vbNullString
        End Try
    End Function

    Private Sub AddLocationToGet(ByRef url As String)
        Try
            If InStr(1, url, "javascript", CompareMethod.Text) = 1 Or InStr(1, url, "doubleclick", CompareMethod.Text) > 0 Then
                'Don't get, crap.
            ElseIf Len(Trim(url)) = 0 Then
                'Don't get, invalid
            Else
                'Debug.Print "Added location to get: " & url
                If Not mToGetList.ContainsKey(url) Then
                    Call mToGetList.Add(url, "")
                End If
            End If
        Catch
        End Try
    End Sub

    Public Function NextLocationToGet() As String
        Try
            Dim toGet As String = ""

            If mToGetList.Count > 0 Then
                NextLocationToGet = ""
                For Each toGet In mToGetList.Keys
                    NextLocationToGet = toGet
                    Exit For
                Next toGet
                If toGet <> "" Then
                    Call mToGetList.Remove(toGet)
                End If
            Else
                Return ""
            End If
        Catch
            Return vbNullString
        End Try
    End Function
End Module