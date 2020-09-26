Option Strict On
Option Explicit On
Module modAPIDeclarations
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
	
	'I've put all of the constants and API function calls in this module - Alasdair - 20 September 2002
	
	
	'CONSTANTS
	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_USERS As Integer = &H80000003
	Public Const HKEY_CURRENT_CONFIG As Integer = &H80000005
	Public Const HKEY_DYN_DATA As Integer = &H80000006
	Public Const REG_SZ As Short = 1 'Unicode nul terminated string
	Public Const REG_BINARY As Short = 3 'Free form binary
	Public Const REG_DWORD As Short = 4 '32-bit number
	Public Const ERROR_SUCCESS As Short = 0
	Public Const RESERVED_NULL As Short = 0
	Public Const INTERNET_CONNECTION_MODEM As Integer = &H1
	Public Const MIIM_SUBMENU As Integer = &H4
	Public Const APINULL As Short = 0
	Public Const EM_SCROLLCARET As Integer = &HB7
	Public Const EM_LINEINDEX As Integer = &HBB
	Public Const EM_GETLINECOUNT As Integer = &HBA
	Public Const EM_LINELENGTH As Integer = &HC1
	Public Const EM_GETFIRSTVISIBLELINE As Integer = &HCE
	Public Const EM_LINEFROMCHAR As Integer = &HC9
	Public Const EM_GETLINE As Integer = &HC4
	Public Const EM_LINESCROLL As Integer = &HB6
	Public Const EM_SETSEL As Integer = &HB1
	Public Const KEY_READ As Integer = &H20019 '-- Permission for general read access.
	Public Const KEY_QUERY_VALUE As Integer = &H1
	Public Const KEY_SET_VALUE As Integer = &H2
	Public Const KEY_ALL_ACCESS As Integer = &H3F
	
	Public Const MF_MENUBREAK As Integer = &H40 'indicates a vertical break
	Public Const MIIM_STATE As Integer = &H1
	Public Const MIIM_ID As Integer = &H2
	Public Const MIIM_CHECKMARKS As Integer = &H8
	Public Const MIIM_TYPE As Integer = &H10
	Public Const MIIM_DATA As Integer = &H20
	Public Const MIIM_STRING As Integer = &H40
	Public Const MIIM_BITMAP As Integer = &H80
	Public Const MIIM_FTYPE As Integer = &H100
	Public Const FBYPOSITION_POSITION As Boolean = True
	Public Const FBYPOSTION_IDENTIFIER As Boolean = False
	Public Const MF_STRING As Integer = &H0
	Public Const GWL_WNDPROC As Integer = (-4)
	Public Const WM_COMMAND As Integer = &H111 ' indicates that a command has been intercepted by the app
	Public Const MF_BYPOSITION As Integer = &H400 ' indicates a menu item by position, not by name, for RemoveMenu
	Public Const MF_DISABLED As Integer = &H2 ' used in setting menu items (InsertMenuItem etc) to indicate that an item is greyed
	Public Const MF_GRAYED As Integer = &H1 ' allegedly does the same thing as MF_DISABLED
	
	'TYPES
	Public Structure DllVersionInfo
		Dim cbSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
	End Structure
	
	Public Structure MENUITEMINFO
		Dim cbSize As Integer
		Dim fMask As Integer
		Dim fType As Integer
		Dim fState As Integer
		Dim wID As Integer
		Dim hSubMenu As Integer
		Dim hbmpChecked As Integer
		Dim hbmpUnchecked As Integer
		Dim dwItemData As Integer
		Dim dwTypeData As String
		Dim cch As Integer
	End Structure
	
	
	'METHODS
	
	'URLDownloadToFile
	'Downloads a specified url to a local file
	Public Declare Function URLDownloadToFile Lib "urlmon"  Alias "URLDownloadToFileA"(ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer
	
	'For removing a menu item
	Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
	
	'For finding the system time - used for timing
	Public Declare Function GetTickCount Lib "kernel32" () As Integer
	Public Declare Function TimeGetTime Lib "winmm.dll" () As Integer
	
	'For finding IE version
	'UPGRADE_WARNING: Structure DllVersionInfo may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function DllGetVersion Lib "Shlwapi.dll" (ByRef dwVersion As DllVersionInfo) As Integer
	'for finding IE version
	
    Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Integer, ByRef lpszConnectionName As Integer, ByRef dwNameLen As Integer, ByVal dwReserved As Integer) As Integer
	
	'This is declaring a Function called
	'"PlaySound" and we tell VB that it's located in the
	'Library called "winmm.dll". The word 'Alias' means that "PlaySound"
	'is actually stored in the DLL as "PlaysoundA", and written in C++ but
	'we can use the function in VB as "PlaySound".
	'lpszname = file path, hmodule = 0 and dwflags = (Synchonous = 0/Asynchronously = 1)
	Public Declare Function PlaySound Lib "winmm.dll"  Alias "PlaySoundA"(ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer
	'returns a long, but 0 = successful play


    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
	Public Declare Function SendMessageW Lib "user32" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	
	Public Declare Function FolderRegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	' Note that if you declare the lpData parameter as String, you must pass it
	' By Value.
	
	'Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
	'must be called before FolderRegQueryEx
	'DEV RegOpenKey is deprecated for RegOpenKeyEx, and I think it's giving me grief with UAC in Vista.
	
	'for constructing the Favorites menu
	Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Integer) As Integer
	
	Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
	
	'for processing the menu favorites
	Public Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	
	Public Declare Function CallWindowProc Lib "user32.dll"  Alias "CallWindowProcA"(ByVal lpPrevWndFunc As Integer, ByVal hwnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	
	'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function InsertMenuItem Lib "user32"  Alias "InsertMenuItemA"(ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Integer
	
	'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function SetMenuItemInfo Lib "user32"  Alias "SetMenuItemInfoA"(ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Integer
	
	'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetMenuItemInfo Lib "user32"  Alias "GetMenuItemInfoA"(ByVal hMenu As Integer, ByVal un As Integer, ByVal b As Integer, ByRef lpMenuItemInfo As MENUITEMINFO) As Integer
	
	Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
	
	Public Declare Function CreateMenu Lib "user32" () As Integer
	
	Public Declare Function CreatePopupMenu Lib "user32" () As Integer
	
	Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Integer) As Integer
	
	Public Declare Function GetModuleHandle Lib "kernel32"  Alias "GetModuleHandleA"(ByVal lpModuleName As String) As Integer
	
	'Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
	
	'for getting the locale ID for the current machine/user
	Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Integer
	
	'Used to determine the (window handle of) the current active window in Windows.
	Public Declare Function GetForegroundWindow Lib "user32" () As Integer
End Module