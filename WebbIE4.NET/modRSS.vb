Option Strict On
Option Explicit On
Module modRSS
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
	
	'modRSS
	'Code for supporting integrated RSS use, such as RSS autodiscovery:
	'http://diveintomark.org/archives/2002/05/30/rss_autodiscovery
	
	
    Public Function CheckForRSSAlternate(ByRef doc As mshtml.IHTMLDocument3) As String
        Try
            'returns the url of an rss feed indicated in the HEAD element
            'of the doc Document as being an alternate to this page.

            Dim metaElements As mshtml.IHTMLElementCollection
            Dim metaIterator As mshtml.IHTMLElement
            Dim needToMakeAbsolute As Boolean
            Dim Path As String
            Dim doc2 As mshtml.IHTMLDocument2

            If doc Is Nothing Then
                Return ""
            Else
                CheckForRSSAlternate = ""
                metaElements = doc.getElementsByTagName("link")
                'Debug.Print "Got meta elements: " & metaElements.length
                For Each metaIterator In metaElements
                    If metaIterator.getAttribute("rel") Is Nothing Then
                    Else
                        If metaIterator.getAttribute("rel").ToString = "alternate" Then
                            If metaIterator.getAttribute("type") Is Nothing Then
                            Else
                                If metaIterator.getAttribute("type").ToString = "application/rss+xml" Then
                                    'found it!
                                    CheckForRSSAlternate = metaIterator.getAttribute("href").ToString
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next metaIterator
                'did we get one?
                If Len(CheckForRSSAlternate) > 0 Then
                    'we did! Is it local or absolute: that is,
                    'check this is fully-qualified, e.g. "/rss.xml" becomes "http://www.site.com/rss.xml"
                    'Debug.Print "Got RSS: " & CheckForRSSAlternate
                    needToMakeAbsolute = False
                    If Len(CheckForRSSAlternate) < Len("http://") Then
                        'obviously not...
                        needToMakeAbsolute = True
                    Else
                        'might be!
                        If Left(CheckForRSSAlternate, Len("http://")) = "http://" Then
                            'okay, no problem, already absolute:
                            needToMakeAbsolute = False
                        Else
                            'nope!
                            needToMakeAbsolute = True
                        End If
                    End If
                    If needToMakeAbsolute Then
                        'yes, we have something like "rss.xml" and we have to make "http://www.mysite.com/rss.xml"
                        doc2 = CType(doc, mshtml.IHTMLDocument2)
                        Path = doc2.location.pathname
                        Path = Left(Path, InStrRev(Path, "/", Len(Path)) - 1)
                        'UPGRADE_WARNING: Couldn't resolve default property of object doc.location. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        CheckForRSSAlternate = doc2.location.protocol & "//" & doc2.location.hostname & Path & "/" & CheckForRSSAlternate
                        '            path = doc.location.pathname
                        '            If Right(path, 1) = "/" Then
                        '                'all ready to have filename added
                        '                CheckForRSSAlternate = path & CheckForRSSAlternate
                        '            Else
                        '                'is there a path (www.site.com/lsdkfjk) or not (www.site.com)?
                        '                If InStr(Len("http://") + 1, path, "/") > 0 Then
                        '                    'yes, there is!
                        '                    CheckForRSSAlternate = Left(path, InStrRev(path, "/", Len(path))) & path
                        '                Else
                        '                    'nope, add straight to path
                        '                    CheckForRSSAlternate = path & "/" & CheckForRSSAlternate
                        '                End If
                        '            End If
                    End If
                End If
            End If
        Catch
            Return ""
        End Try
    End Function
End Module