Option Strict On
Option Explicit On
Module modAccesskeys
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

    'This module handles the use of accesskeys, key presses that trigger particular links

    Private ReadOnly accessKeys As New System.Collections.Generic.Dictionary(Of String, String)
    Private ReadOnly validKeys As String

    Public Sub ClearAccessKeys()
		'clears any existing access keys
        On Error Resume Next
        Call accessKeys.Clear()
	End Sub
	
    ''' <summary>
    ''' Adds an access key to the current collection.
    ''' </summary>
    ''' <param name="anchor"></param>
    ''' <remarks></remarks>
    Public Sub AddAccessKey(ByRef anchor As mshtml.IHTMLAnchorElement)
        Try
            'remove it if it exists already - in case the site has assigned it twice
            If accessKeys.ContainsKey(anchor.accessKey) Then
                Call accessKeys.Remove(anchor.accessKey)
            End If
            Call accessKeys.Add(anchor.accessKey, anchor.href)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Returns the url for an access key, if any. Case insensitive.
    ''' </summary>
    ''' <param name="accessKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetURL(ByRef accessKey As Integer) As String
        Try

            Dim uCaseKey As String
            Dim lCaseKey As String

            uCaseKey = UCase(ChrW(accessKey))
            lCaseKey = LCase(ChrW(accessKey))
            'check we have any accesskeys
            If Not (accessKeys Is Nothing) Then
                If accessKeys.ContainsKey(lCaseKey) Then
                    'yep, we have this
                    GetURL = accessKeys(lCaseKey)
                ElseIf accessKeys.ContainsKey(uCaseKey) Then
                    GetURL = accessKeys(uCaseKey)
                Else
                    'nope, no href
                    GetURL = ""
                End If
            Else
                GetURL = ""
            End If
        Catch
            GetURL = ""
        End Try
    End Function
End Module