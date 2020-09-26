Option Strict On
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmParseHTMLHelp
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
	
	'frmParseHTMLHelp
	
	
	Private Const Document As Integer = 0
	Private mTitle As String
	
	Public Function ConvertHTMLHelp(ByRef Path As String) As String
        Try
            Dim folderPath As String
            Dim i As Integer
            Dim newIndex As String
            Dim doc As mshtml.HTMLDocument
            Dim found As Boolean

            ConvertHTMLHelp = "" ' Will return empty string if any problems.
            'create output folder
            folderPath = System.IO.Path.GetTempPath & "\" & System.IO.Path.GetTempFileName
            Call System.IO.Directory.CreateDirectory(folderPath)
            Dim decompileProcess As System.Diagnostics.Process
            decompileProcess = System.Diagnostics.Process.Start("hh.exe", "-decompile """ & folderPath & """ """ & Path & """")
            decompileProcess.WaitForExit()
            'Call Shell(runLine, vbNormalFocus)
            For i = 0 To 100
                System.Windows.Forms.Application.DoEvents()
            Next i
            'now try to parse index.hhc (or similar) to get help file index
            For Each fi As String In System.IO.Directory.EnumerateFiles(folderPath)
                If System.IO.Path.GetExtension(fi).ToLowerInvariant = "hhc" Then
                    'found an index file, process it for contents
                    'get the new index file ready
                    mTitle = ""
                    newIndex = "<html><head><title>TITLEHERE</title></head><body><h1>TITLEHERE</h1>" & vbNewLine
                    'UPGRADE_WARNING: Navigate2 was upgraded to Navigate and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                    Call mWebBrowser.Navigate(New System.Uri(folderPath & "\" & fi))
                    While mWebBrowser.ReadyState <> System.Windows.Forms.WebBrowserReadyState.Complete
                        System.Windows.Forms.Application.DoEvents()
                    End While
                    doc = CType(mWebBrowser.Document.DomDocument, mshtml.HTMLDocument)
                    newIndex &= ParseHelp(doc.body)
                    newIndex &= "</html>"
                    newIndex = Replace(newIndex, "TITLEHERE", mTitle)
                    Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(folderPath & "\helpindex.htm", False, System.Text.Encoding.UTF8)
                    Call sw.Write(newIndex)
                    Call sw.Close()
                    ConvertHTMLHelp = folderPath & "\helpindex.htm"
                    found = True
                    Exit For
                End If
            Next fi
        Catch
            Return ""
        End Try
    End Function
	
    Private Function ParseHelp(ByRef element As mshtml.IHTMLElement) As String
        Try
            'parses a node in the help file into html
            Dim childIterator As mshtml.IHTMLElement
            Dim childNode As mshtml.IHTMLDOMNode
            Dim childParam As mshtml.IHTMLParamElement

            Dim name_Renamed As String = ""
            Dim ci As mshtml.IHTMLDOMChildrenCollection
            Dim url As String = ""
            Dim node As mshtml.IHTMLDOMNode

            ParseHelp = ""

            node = CType(element, mshtml.IHTMLDOMNode)
            Select Case element.tagName
                Case "BODY"
                    ParseHelp = "<body>" & vbNewLine
                    ci = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                    For Each childIterator In ci
                        ParseHelp = ParseHelp & ParseHelp(childIterator) & vbNewLine
                    Next childIterator
                    ParseHelp = ParseHelp & "</body>" & vbNewLine
                Case "UL"
                    ParseHelp = "<ul>" & vbNewLine
                    ci = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                    For Each childIterator In ci
                        ParseHelp = ParseHelp & ParseHelp(childIterator) & vbNewLine
                    Next childIterator
                    ParseHelp = ParseHelp & "</ul>" & vbNewLine
                Case "LI"
                    ParseHelp = "<li>" & vbNewLine
                    ci = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                    For Each childIterator In ci
                        ParseHelp = ParseHelp & ParseHelp(childIterator) & vbNewLine
                    Next childIterator
                    ParseHelp = ParseHelp & "</li>" & vbNewLine
                Case "OBJECT"
                    If InStr(1, element.innerHTML, "name=""Local""", CompareMethod.Text) > 0 Then
                        'link to local file
                        ci = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                        For Each childNode In ci
                            If childNode.nodeName = "PARAM" Then
                                childParam = CType(childNode, mshtml.IHTMLParamElement)
                                If childParam.name = "Name" Then
                                    'got name
                                    name_Renamed = childParam.value
                                    If Len(mTitle) = 0 Then mTitle = name_Renamed
                                ElseIf childParam.name = "Local" Then
                                    'got url
                                    url = childParam.value
                                End If
                            End If
                        Next childNode
                        ParseHelp = "<a href=""" & url & """>" & name_Renamed & "</a>"
                    ElseIf InStr(1, element.innerHTML, "name=""ImageNumber""") > 0 Then
                        'section heading
                        ci = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                        For Each childNode In ci
                            If childNode.nodeName = "PARAM" Then
                                childParam = CType(childNode, mshtml.IHTMLParamElement)
                                If childParam.name = "Name" Then
                                    'got name
                                    ParseHelp = "<h2>" & name_Renamed & "</h2>" & vbNewLine
                                    Exit For
                                End If
                            End If
                        Next childNode
                    End If
            End Select
        Catch
            Return ""
        End Try
    End Function

    Private Sub frmParseHTMLHelp_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub
End Class