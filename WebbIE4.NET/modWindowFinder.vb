Option Strict Off
Option Explicit On
Module modWindowFinder
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
	
	'This module checks for the existence of an open window with a given name.
	'Use it to check to see if a command-line-started executable has finished
	
	
	Public EnumCAddress As Integer
	Public WindowNames As String
	
	Public Declare Function FindWindowEx Lib "user32"  Alias "FindWindowExA"(ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
	Public Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	
	Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Integer, ByVal lpEnumFunc As Integer, ByVal lParam As Integer) As Integer
	Declare Function EnumWindows Lib "user32" (ByVal lpfn As Integer, ByVal lParam As Integer) As Boolean
	Declare Function GetClassName Lib "user32"  Alias "GetClassNameA"(ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Public Const WM_GETTEXT As Integer = &HD
	Public Const LB_FINDSTRING As Integer = &H18F
	Public Const WM_CLOSE As Integer = &H10
	
	Public Function GetProcAddress(ByVal lAddress As Integer) As Integer
		On Error Resume Next
		GetProcAddress = lAddress
	End Function
	
	Function EnumWindowsProc(ByVal wHandle As Integer, ByVal lParam As Integer) As Integer
		On Error Resume Next
		Dim a As String
		Dim b As String
		Dim l As Integer
		Dim display As String
		
		a = Space(128)
		l = GetClassName(wHandle, a, 128)
		b = CStr(wHandle) & vbTab & Left(a, l) & vbTab
		l = SendMessage(wHandle, WM_GETTEXT, 128, a)
		'If frmMain.chkHide.Value = vbChecked Then
		'        display = Trim(Left$(a, l))
		'        'a = Trim(a)
		'        'Debug.Print "A:" & a & Len(a)
		'        If Len(display) > 0 Then
		'            WindowNames = WindowNames & display & vbNewLine
		'            'WindowNames = WindowNames & "Parent: " & b & Left$(a, l) & vbNewLine
		'        End If
		'    Else
		WindowNames = WindowNames & "Parent: " & b & Left(a, l) & vbNewLine
		'Debug.Print "Parent: " & b & Left$(a, l) & vbNewLine
		'    End If
		'    WindowNames = WindowNames & "Parent: " & Left$(a, l) & vbNewLine
		'If Not (frmMain.chkChildren.Value = vbChecked) Then
		EnumChildWindows(wHandle, EnumCAddress, lParam)
		'End If
		EnumWindowsProc = True
	End Function
	
	
	Function EnumChildWindowsProc(ByVal wHandle As Integer, ByVal lParam As Integer) As Integer
		On Error Resume Next
		Dim a As String
		Dim b As String
		Dim l As Integer
		Dim display As String
		
		a = Space(128)
		l = GetClassName(wHandle, a, 128)
		b = CStr(wHandle) & vbTab & "c" & vbTab & Left(a, l) & vbTab
		l = SendMessage(wHandle, WM_GETTEXT, 128, a)
		' If frmMain.chkHide.Value = vbChecked Then
		'        'Debug.Print "a [" & a & "]"
		'        'Debug.Print "trim a [" & Trim(a) & "]"
		'        display = Trim(Left$(a, l))
		'        'a = Trim(a)
		'        'Debug.Print "A:" & a & " " & Len(a)
		'        If Len(display) > 0 Then
		'            WindowNames = WindowNames & display & vbNewLine
		'            'WindowNames = WindowNames & "Child: " & b & Left$(a, l) & vbNewLine
		'        End If
		'    Else
		'check to see if we've found this already
		If InStr(1, WindowNames, "Child: " & b & Left(a, l)) > 0 Then
			'hey, we've found this already! Stop.
		Else
			'keep going
			WindowNames = WindowNames & "Child: " & b & Left(a, l) & vbNewLine
			'Debug.Print "Child: " & b & Left$(a, l) & vbNewLine
		End If
		'    WindowNames = WindowNames & "Child: " & Left$(a, l) & vbNewLine
		' If Not (frmMain.chkChildren.Value = vbChecked) Then
		Call EnumChildWindows(wHandle, EnumCAddress, lParam)
		'End If
		EnumChildWindowsProc = True
	End Function
End Module