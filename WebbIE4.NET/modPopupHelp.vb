Option Strict On
Option Explicit On
Module modPopupHelp
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
	
	'This module handles pop-up or context help, obtained by hitting
	'f1 while in a control
	
	
    Private popupHelpContents As New System.Collections.Generic.Dictionary(Of String, String)
	
	Public Sub popupHelp(ByRef formName As String, ByRef aControl As System.Windows.Forms.Control)
        'popup help called. controlName is the name of the current control
        'TODO 
    End Sub
	
	'Private Sub Initialise()
	''load the help file for the correct language
    '    
	'    Dim helpDoc As New MSXML2.DOMDocument30
	'    Dim nodeIterator As IXMLDOMNode
	'    Dim contentNode As IXMLDOMNode
	'
	'    helpDoc.async = False
	'    Call helpDoc.Load(App.Path & "\language.xml")
	'    For Each nodeIterator In helpDoc.documentElement.selectSingleNode("popupHelp").childNodes
	'        Set contentNode = nodeIterator.selectSingleNode("content[@language='" & modI18N.language & "']")
	'        Call popupHelpContents.Add(nodeIterator.selectSingleNode("key").Text, contentNode.Text)
	'    Next nodeIterator
	'    'popupHelpContents now contains control names indexing
	'    'help comments.
	'    initialised = True
	'End Sub
End Module