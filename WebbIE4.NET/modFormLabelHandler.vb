Option Strict On
Option Explicit On
Module modFormLabelHandler
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

    'modFormLabelHandler
    'Handles elements of type LABEL. These are descriptive labels for FORM elements,
    'but they can appear anywhere in the HTML, so we can't just populate as we
    'iterate through the page.

    Private ReadOnly mLabels As New System.Collections.Generic.Dictionary(Of String, String)

    Public Sub Clear()
		'clear the current label look-up
        On Error Resume Next
        Call mLabels.Clear()
    End Sub

    ''' <summary>
    ''' Find any label elements in the page and populate the dictionary
    ''' </summary>
    ''' <param name="page"></param>
    ''' <remarks></remarks>
    Public Sub ProcessWebpageForLabels(ByRef page As mshtml.HTMLDocument)
        Try
            Dim label As mshtml.IHTMLElement
            Dim labelElement As mshtml.IHTMLLabelElement
            Dim descriptiveText As String

            'check for labels
            'DEV: alas, we can't just check for forms: controls (and labels)
            'can exist outside of forms
            For Each label In page.getElementsByTagName("LABEL")
                'found a label!
                labelElement = CType(label, mshtml.IHTMLLabelElement)
                'find which form element it applies to
                If labelElement.htmlFor Is Nothing Then
                    'See below...
                ElseIf labelElement.htmlFor.Length = 0 Then
                    'nope, no applied form element: it must apply to the contained
                    'control element. We'll handle this in the parsing mechanism when
                    'we have a chance to extract the label contents. See ParseNode.
                Else
                    'got a target and applied form element
                    'work out the descriptive text
                    descriptiveText = ""
                    If label.innerText <> "" Then
                        descriptiveText = label.innerText
                        descriptiveText = descriptiveText.Replace(vbCr, "").Replace(vbLf, " ").Trim ' remove any newlines, replace with spaces
                    End If
                    If Len(descriptiveText) = 0 Then
                        'Use our parse-and-get-some-text code.
                        descriptiveText = ParseForText(CType(label, mshtml.IHTMLDOMNode))
                    Else
                        'okay, we have our descriptive text we can use.
                    End If
                    'add to the label look-up if we have managed to find some descriptive text
                    If Len(descriptiveText) > 0 Then
                        '3.6.9 Trim the descriptive text. IE will dump a huge LABEL element on you if the site
                        'miscodes.
                        'In theory a label should not exist when it gets here - if it does, ignore the
                        'latest one.
                        If mLabels.ContainsKey(labelElement.htmlFor) Then
                            Debug.Print("Warning: page has multiple LABEL elements for one id. Id=" & labelElement.htmlFor)
                        Else
                            'Was fifty, increased to 512 after someone complained their labels were getting truncated.
                            Call mLabels.Add(labelElement.htmlFor, Left(descriptiveText, 512))
                        End If
                    End If
                End If
            Next label
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Returns the descriptive text for a control.
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    ''' <remarks>Strictly, this should only use the id for a control. But, of course, people
    ''' won't code pages like that, so fall back to the name if no id. 
    '''returns empty string if not found
    '''</remarks>
    Public Function GetDescriptiveText(ByRef control As mshtml.IHTMLElement) As String
        Try
            Dim controlID As String = ""
            If control.id Is Nothing Then
            ElseIf control.id.Length = 0 Then
            Else
                controlID = control.id.ToString()
            End If
            If controlID = "" Then
                If control.getAttribute("name") Is Nothing Then
                Else
                    controlID = control.getAttribute("name").ToString()
                End If
            End If
            If controlID = "" Then
                controlID = control.GetHashCode().ToString()
                'System.Diagnostics.Debug.Print("HASH ?:" & controlID)
            End If
            If mLabels.ContainsKey(controlID) Then
                GetDescriptiveText = mLabels.Item(controlID).ToString.Trim
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try
    End Function

    'This function allows parsing to handle situations where a
    'label is not explicitly associated with a form element by id but contains the
    'form element it labels, e.g. <label>Name <input type="text" id="name"></label>
    Public Sub AddLabel(ByRef labelText As String, ByRef labelControl As mshtml.IHTMLElement)
        Try

            Dim oldContent As String
            Dim element2 As mshtml.IHTMLElement2

            element2 = CType(labelControl, mshtml.IHTMLElement2)
            If labelControl Is Nothing Then
                'whoops, not got the control to be labelled. Do nothing.
            ElseIf element2.currentStyle.display = "none" Then
                'Whoops, this element should not be rendered. Do nothing.
            Else
                'okay, already got control for which this is the label
                'What's the ID?
                Dim id As String = ""
                Try
                    If Len(labelControl.getAttribute("id")) > 0 Then
                        id = labelControl.getAttribute("id").ToString()
                    End If
                Catch
                    'DOM error - maybe missing id attribute? Don't know.
                    'Fall back to hash (below)
                End Try
                If id = "" Then
                    'whoops, no id on the control. We could give up, but it's entirely 
                    'likely that we'll find controls without ids inside labels in the
                    'wild - working with event handlers, or array indices. So have
                    'to label, even though we don't have an id. What do we have?
                    'Let's use the Hash Code?
                    id = labelControl.GetHashCode().ToString()
                    System.Diagnostics.Debug.Print("HASH +:" & id)
                End If
                If mLabels.ContainsKey(id) Then
                    'Already got label: concatenate with existing contents. This is
                    'completely valid HTML, even if no-one else supports it.
                    oldContent = mLabels.Item(id)
                    If Len(oldContent) > 0 Then
                        If InStr(1, ".!?:", Right(oldContent, 1)) = 0 Then
                            oldContent &= "."
                        End If
                    End If
                    labelText = oldContent & " " & labelText
                    Call mLabels.Remove(id)
                End If
                Call mLabels.Add(id, labelText)
            End If
        Catch
        End Try
    End Sub
End Module