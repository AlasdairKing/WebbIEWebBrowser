Option Strict On
Option Explicit On
Module modAPIFunctions
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

    Public Sub Scroll(ByRef editControl As System.Windows.Forms.Control, ByRef lines As Integer)
		'scrolls the editControl textbox by lines lines
		' editControl - the current control
		' lines - the number of lines to scroll
        On Error Resume Next
        Call NativeMethods.SendMessageW(editControl.Handle, NativeMethods.EM_LINESCROLL, IntPtr.Zero, CType(lines, IntPtr))
    End Sub
	

    Public Function GetCurrentLine(ByRef editControl As System.Windows.Forms.RichTextBox) As String
        On Error Resume Next
        Dim startOfLine As Integer = editControl.GetFirstCharIndexOfCurrentLine
        Return RestOfLine(editControl, startOfLine)
    End Function

    'Public Function GetCurrentLine(ByRef editControl As System.Windows.Forms.RichTextBox) As String
    '    'returns the text on the current line from the current control
    '    '   editControl - the current control
    '    '   charPos - the position of the cursor in the control
    '    Dim lineNumber As IntPtr ' the line number of the current line
    '    'get the line number: arguments of -1 mean "line with caret", 0 is not used.
    '    lineNumber = SendMessage(editControl.Handle, EM_LINEFROMCHAR, New IntPtr(-1), New IntPtr(0))
    '    'We can get a "line number" from either SendMessage (the old VB6 way) or from the actual
    '    'edit control, and they agree. However, they are both wrong: if I is the caret:
    '    ' Line 1<newline>
    '    ' LineI 2           Correctly returns 1, and editControl.Lines(1) is "Line 2" Great!
    '    '
    '    'But now what if the text wraps? Then it all goes to crap.
    '    ' Line
    '    ' 1<newline>
    '    ' LineI                 Now returns 2, which is fine - it's the third line of text. BUT
    '    ' 2                     .lines() hasn't changed, so you get a blank piece of text - you 
    '    '                       still need to get .lines(1)!
    '    'So you can't do this:
    '    '        Return editControl.Lines(editControl.GetLineFromCharIndex(editControl.SelectionStart))
    '    'Instead you must still use the API, which I've done in GetLine.

    '    Return GetLine(editControl, CInt(lineNumber))

    'End Function

    Public Function GetLine(ByRef editControl As System.Windows.Forms.RichTextBox, ByVal ZeroBasedLinetoRead As Integer) As String
        On Error Resume Next
        Dim startOfLine As Integer = editControl.GetFirstCharIndexFromLine(ZeroBasedLinetoRead)
        Return RestOfLine(editControl, startOfLine)
    End Function

    Private Function RestOfLine(editControl As System.Windows.Forms.RichTextBox, startOfLine As Integer) As String
        'It appears to be vbCr for TextBox controls, and vbLf for RichTextBox controls.
        Dim endOfLine As Integer
        Try
            endOfLine = editControl.Text.IndexOf(vbLf, startOfLine)
        Catch ex As ArgumentOutOfRangeException
            Return ""
        End Try
        Try
            If endOfLine = -1 Then endOfLine = editControl.Text.Length - 1
            Return editControl.Text.Substring(startOfLine, endOfLine - startOfLine)
        Catch ex As ArgumentOutOfRangeException
            Return ""
        End Try
    End Function

    'Public Function GetLine(ByRef editControl As System.Windows.Forms.RichTextBox, ByVal ZeroBasedLinetoRead As Integer) As String
    '    'First, get the length of the line.
    '    Dim lineIndex As Integer = 
    '    Dim lineLen As Integer = CInt(SendMessage(editControl.Handle, EM_LINEFROMCHAR, New IntPtr(ZeroBasedLinetoRead), New IntPtr(0)))
    '    If lineLen = 0 Then
    '        Return String.Empty
    '    Else
    '        'Now construct a buffer to get the text content.
    '        Dim byte1 As String = Chr(lineLen And &HFF)
    '        Dim byte2 As String = Chr(CInt(lineLen / &H100))
    '        Dim bBuffer As String = byte1 & byte2 & New String(ChrW(32), lineLen)
    '        'OK, try to get back.
    '        Dim gotChars As Integer = CInt(SendMessage(editControl.Handle, EM_GETLINE, New IntPtr(ZeroBasedLinetoRead), bBuffer))
    '        If Err.LastDllError <> 0 Then
    '            Debug.Print("DLL Error: " & Err.LastDllError)
    '        End If
    '        If gotChars > 0 Then
    '            Return bBuffer.Substring(0, gotChars)
    '        Else
    '            Return String.Empty
    '        End If
    '    End If

    'End Function

    Public Function GetCurrentLineIndex(ByRef editControl As System.Windows.Forms.RichTextBox) As Integer
        On Error Resume Next
        Return editControl.GetLineFromCharIndex(editControl.SelectionStart)
    End Function


    Public Sub SetCurrentLineIndex(ByRef editControl As System.Windows.Forms.RichTextBox, ByRef line As Integer)
        On Error Resume Next
        editControl.SelectionStart = editControl.GetFirstCharIndexFromLine(line)
    End Sub

    ''' <summary>
    ''' Returns the number of lines in the control.
    ''' </summary>
    ''' <param name="editControl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNumberOfLines(ByRef editControl As System.Windows.Forms.RichTextBox) As Integer
        On Error Resume Next

        'This gives the actual number of lines as currently displayed, taking into account word wrap.
        'So while editControl.Lines has the number of [text chunks with newlines] and never changes 
        'even as you resize the control, GetNumberOfLines gives the number of [lines actually shown
        'in the control] which will be the same or greater.
        'GetNumberOfLines = modAPIDeclarations.SendMessage(editControl.Handle.ToInt32, EM_GETLINECOUNT, 0, 0)
        Return editControl.GetLineFromCharIndex(editControl.Text.Length - 1)
    End Function

    ''' <summary>
    ''' Returns the number of characters in the control up to the numbered line
    ''' </summary>
    ''' <param name="editControl"></param>
    ''' <param name="lineNumber"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCharacterIndexOfLine(ByRef editControl As System.Windows.Forms.RichTextBox, ByRef lineNumber As Integer) As Integer
        On Error Resume Next
        'GetCharacterIndexOfLine = modAPIDeclarations.SendMessage(editControl.Handle.ToInt32, EM_LINEINDEX, lineNumber, 0)
        Return editControl.GetFirstCharIndexFromLine(lineNumber)
    End Function

End Module