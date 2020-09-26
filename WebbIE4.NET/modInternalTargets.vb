Option Strict On
Option Explicit On
Module modInternalTargets
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
	
	'This module contains code to handle Fragment Identifiers or internal
	'targets, such as mypage.html#internaltarget
	
	'How to use
	'1 Call DetermineInternalTarget with the webbrowser and url from beginNavigation
	'  This sets mstrInternalTarget and mTargetObject
	'  We can now see if there is an internal target by checking the length
	'  of mstrInternalTarget
	
	
    Private mInternalTarget As String
    Private mMarkerNode As mshtml.IHTMLDOMNode

    Public Function DetermineInternalTarget(ByRef webbrowserObject As Object, ByRef url As String) As Boolean
        Try
            'checks that the url sent does not indicate an internal target to be used
            mInternalTarget = ExtractInternalTargetFromURL(url)
            If Len(mInternalTarget) > 0 Then
                'got some kind of target: remember this web browser object
                DetermineInternalTarget = True
            Else
                'nope, forget about it
                DetermineInternalTarget = False
            End If
            Debug.Print("Internal Target: " & mInternalTarget)
        Catch
            Return False
        End Try
    End Function

    Public Function ExtractNonInternalTargetFromURL(ByVal url As String, ByRef doc As mshtml.IHTMLDocument2) As String
        Try
            'Does the opposite of ExtractInternalTargetFromURL: returns the part of a url
            'that is NOT the optional HTML target

            Dim a As mshtml.IHTMLAnchorElement
            Dim internalTarget As String

            doc = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.IHTMLDocument2)
            If doc Is Nothing Then
                'whoops, can't work it out! Do it by hand
                internalTarget = ManuallyExtractInternalTargetFromURL(url)
                ExtractNonInternalTargetFromURL = url
                If Len(internalTarget) > 0 Then
                    ExtractNonInternalTargetFromURL = Replace(ExtractNonInternalTargetFromURL, "#" & internalTarget, "")
                End If
            Else
                a = CType(doc.createElement("A"), mshtml.IHTMLAnchorElement)
                a.href = url
                ExtractNonInternalTargetFromURL = Replace(url, a.hash, "")
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Function ExtractInternalTargetFromURL(ByVal url As String) As String
        Try
            Dim a As mshtml.IHTMLAnchorElement
            Dim doc As mshtml.IHTMLDocument2

            doc = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.IHTMLDocument2)
            If doc Is Nothing Then
                'need to do by hand
                ExtractInternalTargetFromURL = ManuallyExtractInternalTargetFromURL(url)
            Else
                'got an anchor element, can use that
                a = CType(doc.createElement("A"), mshtml.IHTMLAnchorElement)
                a.href = url
                ExtractInternalTargetFromURL = Replace(a.hash, "#", "")
            End If
        Catch
            Return ""
        End Try
    End Function

    Private Function ManuallyExtractInternalTargetFromURL(ByVal url As String) As String
        Try
            'returns the optional html target e.g. myfile.htm#target3
            'DEV: why was this ByVal? No idea. Not working properly without it. Something
            'to do with URL being a Variant type?


            Dim counter As Integer
            Dim removedQuery As String

            'okay, I used to remove any query string (after "?") but this is incorrect
            'according to http://www.ietf.org/rfc/rfc2396.txt
            '    'truncate any argument starting ?
            '    counter = InStrRev(url, "?")
            '    If counter = 0 Then counter = Len(url)
            '    removedQuery = Left(url, counter)
            removedQuery = url
            'replace any "%20" with spaces: fragment identifiers (internal targets)
            'should not have them, but people will use them (e.g. Lloyds UK)
            removedQuery = Replace(removedQuery, "%20", " ")
            'get any part past a # character indicating an internal link
            counter = InStrRev(removedQuery, "#")
            If counter = 0 Then
                ManuallyExtractInternalTargetFromURL = ""
            Else
                ManuallyExtractInternalTargetFromURL = Right(removedQuery, Len(removedQuery) - counter)
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Sub LabelInternalTarget(ByRef Document As mshtml.HTMLDocument)
        Try
            'labels the appropriate element in document with an identifying
            'text node child that will show up in the parsed document. A
            'reference is kept to this child - mMarkerNode - so it can be excised
            'later.

            Dim targetNode As mshtml.IHTMLDOMNode

            'first check we don't have any already-labelled links
            If InStr(1, Document.body.innerText, TARGET_MARKER, CompareMethod.Binary) > 0 Then
                'oh, we've already labelled a target node: have to take it off!

            End If
            'Debug.Print "Looking for: " & mstrInternalTarget
            If Len(mInternalTarget) > 0 Then
                Debug.Print("Looking for: " & mInternalTarget)
                'okay, we have to find the target node
                targetNode = CType(Document.getElementById(mInternalTarget), mshtml.IHTMLDOMNode)
                If Not (targetNode Is Nothing) Then
                    'got it!
                    'create marker node as a child
                    mMarkerNode = Document.createTextNode(TARGET_MARKER)
                    'note insertBefore and appendChild won't work if you have not
                    'generated the node in the same document as you are inserting it.
                    If targetNode.hasChildNodes Then
                        'put marker node before all children
                        Call targetNode.insertBefore(mMarkerNode, targetNode.firstChild)
                    Else
                        'marker node is only child
                        'targetNode.innerText = targetNode.innerText & TARGET_MARKER 'don't know why I wanted to
                        'do this: it seems to leave extra text on the screen (as
                        'would be expected)
                        Call targetNode.appendChild(mMarkerNode)
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    'Public Sub MoveCursorToInternalLink(ByRef myForm As frmMain)
    '    'Moves the cursor to any incidence of the target marker and deletes
    '    'the marker from the text display and the DOM (mMarkerNode)
    '    'Operates on frmMain.txtText

    '    Dim selectStart As Integer
    '    Dim parentNode As mshtml.IHTMLDOMNode

    '    Debug.Print("Moving to: " & mstrInternalTarget)
    '    If Len(mstrInternalTarget) > 0 Then
    '        'okay, we've checked we're following an internal target
    '        selectStart = InStr(1, myForm.txtText.Text, TARGET_MARKER, CompareMethod.Text) - 1
    '        If selectStart > -1 Then ' check we found the internal link: if not, don't do anything but clear module variables
    '            'move cursor to correct position and remove marker from body text
    '            'I did use .SelText = "", but this didn't seem to make any
    '            'changes to the text passed through, so using Replace instead
    '            myForm.txtText.Text = Replace(myForm.txtText.Text, TARGET_MARKER, "")
    '            myForm.txtText.SelectionStart = selectStart
    '            'now remove marker node from the DOM
    '            If Not (mMarkerNode Is Nothing) Then
    '                parentNode = mMarkerNode.parentNode
    '                Call parentNode.removeChild(mMarkerNode)
    '                Debug.Print(parentNode.nodeName)
    '            End If
    '        End If
    '        'finished: clear everything
    '        'UPGRADE_NOTE: Object mMarkerNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '        mMarkerNode = Nothing
    '        mstrInternalTarget = ""
    '    End If
    'End Sub
End Module