Option Explicit On
Module modCommandProcessor
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
	
	'Does the processing of run-time created menu items for the bookmarks menu.
	Public targetForm As frmMain
	
	Public Function MenuProcessor(ByVal hwnd As Integer, ByVal uMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
		On Error GoTo errorHandler
		'process user commands: in other words, pass all the user commands/messages to the
		'normal WindowProc with the exception of those corresponding to the bookmark menu.
		'This has been constructed at run-time, so we have to build the processing here.
		'These clicks can be spotted by uMsg = WM_COMMAND and the low word of the wParam Long
		'corresponding to the range for bookmark menu items indicated by being greater than
		'BOOKMARK_ITEM_MARKER.  So, if there are more than BIM (1024) non-bookmark
		'menu items, this will break - those menu items will not work.
		If uMsg = WM_COMMAND And Loword(wParam) > BOOKMARK_ITEM_MARKER Then
			'process my stuff
			Call targetForm.StartNavigating(userBookmarks(Loword(wParam) - BOOKMARK_ITEM_MARKER))
			
			'frmMain.cboAddress.Text = userBookmarks(Loword(wParam) - BOOKMARK_ITEM_MARKER)
			'Call frmMain.cmdGo_Click
			MenuProcessor = 0 ' don't do any more processing
		Else
			'otherwise, use the "normal" response routine
			MenuProcessor = CallWindowProc(hNormalProc, hwnd, uMsg, wParam, lParam)
		End If
		Exit Function
errorHandler: 
		Debug.Print("ERR: MenuProcessor: " & Err.Description)
		Resume Next
	End Function
	
	Public Function Loword(ByRef inputNumber As Integer) As Integer

		'return the low word of the long inputNumber
		Loword = inputNumber And &HFFFF
	End Function
End Module