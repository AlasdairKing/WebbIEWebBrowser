Option Strict On
Option Explicit On
Module modZoom
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
	
	'Handles magnification of the IE view
	
	
	Private mZoom As Double ' the zoom value for the web page in Internet Explorer
	
    Public Sub ZoomIE(ByRef domDocument As Object, ByRef direction As Integer)
        'zooms IE view by 0.05 (e.g. 5% on first zoom...) for IE view: passed mAlldocs
        'from frmMain to work on.

        'TODO Process all child Documents of this one. 

        Dim htmlElement As mshtml.IHTMLElement
        Dim styleObj As mshtml.IHTMLStyle3
        Dim changedZoom As Boolean
        Dim es As mshtml.IHTMLElementCollection
        Dim doc As mshtml.HTMLDocument

        doc = CType(domDocument, mshtml.HTMLDocument)
        If mZoom = 0 Then mZoom = 1.0#

        'first increase mZoom: this should have been reset to 1.0 by the ClearPageData
        'sub in frmMain, so we can always increase it.
        If direction < 0 Then
            mZoom = mZoom - 0.05
            If mZoom < 0.1 Then
                mZoom = 0.1
            Else
                changedZoom = True
            End If
        ElseIf direction = 0 Then
            'Just apply current level.
        ElseIf direction > 0 Then
            mZoom = mZoom + 0.05 ' 0.05 obtained by quick testing, no other rationale
            'now apply this zoom to every element
            changedZoom = True
        End If
        'Apply zoom if required - that is, not x1, or changed to x1.
        'Hmm, might want to not do zoom on some elements to prevent sideways scrolling.
        If (mZoom < 0.99 Or mZoom > 1.01) Or changedZoom Then
            es = CType(doc.body.all, mshtml.IHTMLElementCollection)
            For Each htmlElement In es
                styleObj = CType(htmlElement.style, mshtml.IHTMLStyle3)
                styleObj.zoom = mZoom
            Next htmlElement
        End If
    End Sub
End Module