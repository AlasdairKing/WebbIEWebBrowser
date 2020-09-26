Option Strict On
Option Explicit On
Module modParseHTML
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
	
	
    Public Function Render(ByRef html As String) As String
        Try
            Dim i As Integer
            Dim output As String = ""
            Dim startAlt As Integer
            Dim endTag As Integer
            Dim startHREF As Integer
            Dim attr As String

            i = 1
            html = Trim(html)
            html = Replace(html, vbNewLine, " ")
            html = Replace(html, "  ", " ")
            While i <= Len(html)
                If Mid(html, i, 1) = "<" Then
                    'got a tag! Is it an image?
                    If LCase(Mid(html, i, 5)) = "<img " Then
                        'it's an image. Try to extract the alt.
                        startAlt = InStr(i, html, "alt=", CompareMethod.Text)
                        endTag = InStr(i, html, ">")
                        If startAlt > 0 And startAlt < endTag Then
                            startAlt = startAlt + Len("alt=")
                            attr = GetAttributeValue(Right(html, Len(html) - startAlt + 1))
                            output = output & attr
                        End If
                        i = endTag + 1
                    ElseIf (Mid(html, i, 3)) = "<a " Then
                        'Link
                        endTag = InStr(i + 1, html, ">")
                        If endTag > 0 Then
                            startHREF = InStr(i + 1, html, "href=", CompareMethod.Text)
                            If startHREF > 0 And startHREF < endTag Then
                                startHREF = startHREF + Len("href=")
                                attr = GetAttributeValue(Right(html, Len(html) - startHREF + 1))
                                gNumLinks += 1
                                If gNumLinks <= MAX_NUMBER_LINKS_SUPPORTED Then
                                    gLinks(gNumLinks).address = attr
                                    If Len(Trim(output)) > 0 Then
                                        output = output & vbNewLine
                                    End If
                                    output = output & ID_LINK & " " & gNumLinks & ": "
                                End If
                            Else
                                'Malformed, or no href
                            End If
                            i = endTag + 1
                        Else
                            'No closing >
                            i = Len(html) + 1
                        End If
                    Else
                        i = InStr(i + 1, html, ">")
                        If i = 0 Then i = Len(html)
                        i = i + 1
                    End If
                    'If right(output, 1) <> " " Then output = output & " "
                Else
                    'got a character
                    output = output & Mid(html, i, 1)
                    i = i + 1
                End If
            End While
            Render = StripTerminalWhitespace(output)
        Catch
            Return ""
        End Try
    End Function
	
    Private Function StripTerminalWhitespace(ByRef s As String) As String
        Try
            'removes whitespace at the start or end of s
            Dim removed As Boolean
            Dim l As String

            removed = True
            While removed
                removed = False
                'spaces at start/end
                l = CStr(Len(s))
                s = Trim(s)
                If CDbl(l) <> Len(s) Then removed = True
                'newlines at start
                If Left(s, Len(vbNewLine)) = vbNewLine Then
                    s = Right(s, Len(s) - Len(vbNewLine))
                    removed = True
                End If
                'newlines at end
                If Right(s, Len(vbNewLine)) = vbNewLine Then
                    s = Left(s, Len(s) - Len(vbNewLine))
                    removed = True
                End If
                'tabs at start
                If Left(s, Len(vbTab)) = vbTab Then
                    s = Right(s, Len(s) - Len(vbTab))
                    removed = True
                End If
                'tabs at end
                If Right(s, Len(vbTab)) = vbTab Then
                    s = Left(s, Len(s) - Len(vbTab))
                    removed = True
                End If
            End While
            StripTerminalWhitespace = s
        Catch
            Return s
        End Try
    End Function
	
    Private Function GetAttributeValue(ByRef html As String) As String
        Try
            'Parses some html, starting with the bit after the = in an element, and returns the content of the attribute.
            Dim startAttributeValue As Integer
            Dim endAttributeValue As Integer
            Dim endTag As Integer

            endTag = InStr(1, html, ">")
            If endTag = 0 Then endTag = Len(html)
            If Left(html, 1) = """" Then
                startAttributeValue = 2
                endAttributeValue = InStr(2, html, """")
            ElseIf Left(html, 1) = "'" Then
                startAttributeValue = 2
                endAttributeValue = InStr(2, html, "'")
            Else
                startAttributeValue = 1
                endAttributeValue = InStr(1, html, " ")
            End If
            If endAttributeValue > endTag Or endAttributeValue = 0 Then
                endAttributeValue = endTag
            End If
            GetAttributeValue = Mid(html, startAttributeValue, endAttributeValue - startAttributeValue)
        Catch
            Return ""
        End Try
    End Function
End Module