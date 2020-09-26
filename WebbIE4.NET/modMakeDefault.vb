Option Strict On
Option Explicit On
Module modMakeDefault
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
	
	'This module should provide function to let you register
	'your application as the default handler for any mime
	'type or file extension.
	
	
	'To alert Windows to update start menu
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
	Private Const WM_SETTINGCHANGE As Integer = &H1A
	Private Const HWND_BROADCAST As Integer = &HFFFF
	
	'   *****************REGISTRY CODE********************
	' This allows programs to write to and read from the Windows registry.
	' Code taken from Microsoft HOWTO 145679
	
	'types of data returned
	Private Const REG_SZ As Integer = 1
	Private Const REG_DWORD As Integer = 4
	
	'the four root classes available
	Private Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Private Const HKEY_CURRENT_USER As Integer = &H80000001
	Private Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Private Const HKEY_USERS As Integer = &H80000003
	
	'the error messages: see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/debug/base/system_error_codes__0-499_.asp
	Private Const ERROR_NONE As Short = 0
	Private Const ERROR_BADDB As Short = 1
	Private Const ERROR_BADKEY As Short = 2
	Private Const ERROR_CANTOPEN As Short = 3
	Private Const ERROR_CANTREAD As Short = 4
	Private Const ERROR_CANTWRITE As Short = 5
	Private Const ERROR_OUTOFMEMORY As Short = 6
	Private Const ERROR_ARENA_TRASHED As Short = 7
	Private Const ERROR_ACCESS_DENIED As Short = 8
	Private Const ERROR_INVALID_PARAMETERS As Short = 87
	Private Const ERROR_MORE_DATA As Short = 234
	Private Const ERROR_NO_MORE_ITEMS As Short = 259
	
	'access types
	Private Const KEY_QUERY_VALUE As Integer = &H1
	Private Const KEY_SET_VALUE As Integer = &H2
	Private Const KEY_ALL_ACCESS As Integer = &H3F
	
	Private Const REG_OPTION_NON_VOLATILE As Short = 0
	
	Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	'close key after use
	
	Private Declare Function RegCreateKeyEx Lib "advapi32.dll"  Alias "RegCreateKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, ByVal lpSecurityAttributes As Integer, ByRef phkResult As Integer, ByRef lpdwDisposition As Integer) As Integer
	'create a key
	
	Private Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	'open a key for reading or writing
	
	Private Declare Function RegQueryValueExString Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	'get the contents of a string key
	
	Private Declare Function RegQueryValueExLong Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer
	'get the contents of a long key (DWORD)
	
	'Get child keys of main key
	'UPGRADE_WARNING: Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function RegEnumKeyEx Lib "advapi32.dll"  Alias "RegEnumKeyExA"(ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpName As String, ByRef lpcbName As Integer, ByVal lpReserved As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByRef lpftLastWriteTime As FILETIME) As Integer
	
	Private Declare Function RegQueryValueExNULL Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As Integer, ByRef lpcbData As Integer) As Integer
	
	Private Declare Function RegSetValueExString Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal reserved As Integer, ByVal dwType As Integer, ByVal lpValue As String, ByVal cbData As Integer) As Integer
	
	Private Declare Function RegSetValueExLong Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal reserved As Integer, ByVal dwType As Integer, ByRef lpValue As Integer, ByVal cbData As Integer) As Integer
	
	Private Declare Function RegDeleteKey Lib "advapi32.dll"  Alias "RegDeleteKeyA"(ByVal lngRootKey As Integer, ByVal lpSubKey As String) As Integer
	
	Private Structure FILETIME 'ft
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure
	
	Private Sub RegisterMimeAndExtension(ByRef progID As String, ByRef exePath As String, ByRef canonicalExeName As String, ByRef canonicalName As String, ByRef regProgPath As String, ByRef CLSID As String, ByRef extension As String, ByRef mimeType As String)
		'make the ActiveX object you provide the handler for the extension
		'and mimeType specified

		'ASSUME this is done by the Installer!
        ''1 Apply mimetype and extension
        ''mimetype:
        'If Not KeyExists(HKEY_CLASSES_ROOT, "MIME\DataBase\Content Type\" & mimeType) Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "MIME\DataBase\Content Type\" & mimeType)
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "MIME\DataBase\Content Type\" & mimeType, "Extension", extension, REG_SZ)
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "MIME\DataBase\Content Type\" & mimeType, "CLSID", CLSID, REG_SZ)
        ''And on ActiveX object:
        'If Not KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\EnableFullPage\" & extension) Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\EnableFullPage\" & extension)
        'End If
        ''And finally on extension:
        'If Not KeyExists(HKEY_CLASSES_ROOT, extension) Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, extension)
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, extension, "", progID, REG_SZ)
        'result = SetKeyValue(HKEY_CLASSES_ROOT, extension, "ContentType", mimeType, REG_SZ)
	End Sub
	
	Private Sub PrepareActiveXObject(ByRef progID As String, ByRef exePath As String, ByRef CLSID As String)

		'1 Prep our object for use.
        'Done by installer!
        ''CLSID
        'If Not KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & CLSID) Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" & CLSID)
        'End If
        'If Not KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\InProcServer32") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\InProcServer32")
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\InProcServer32", "", exePath, REG_SZ)
        'If Not KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\ProgID") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\ProgID")
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\ProgID", "", progID, REG_SZ)
        'If Not KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\Control") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\Control")
        'End If
        'If Not KeyExists(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\EnableFullPage") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" & CLSID & "\EnableFullPage")
        'End If
		'Extensions go here!
		
		'ProgID
        'If Not KeyExists(HKEY_CLASSES_ROOT, progID) Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, progID)
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, progID, "", progID, REG_SZ)
        'If Not KeyExists(HKEY_CLASSES_ROOT, progID & "\DefaultIcon") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, progID & "\DefaultIcon")
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, progID & "\DefaultIcon", "", exePath & ",0", REG_SZ)
        'If Not KeyExists(HKEY_CLASSES_ROOT, progID & "\Clsid") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, progID & "\Clsid")
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, progID & "\Clsid", "", CLSID, REG_SZ)
        'If Not KeyExists(HKEY_CLASSES_ROOT, progID & "\Shell") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, progID & "\Shell")
        'End If
        'If Not KeyExists(HKEY_CLASSES_ROOT, progID & "\Shell\Open") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, progID & "\Shell\Open")
        'End If
        'If Not KeyExists(HKEY_CLASSES_ROOT, progID & "\Shell\Open\Command") Then
        '	result = CreateNewKey(HKEY_CLASSES_ROOT, progID & "\Shell\Open\Command")
        'End If
        'result = SetKeyValue(HKEY_CLASSES_ROOT, progID & "\Shell\Open\Command", "", """" & exePath & " ""%1""", REG_SZ)
	End Sub
	
	
	'   *****************REGISTRATION CODE ********************
	Private Sub RegisterAsDefaultBrowser(ByRef progID As String, ByRef exePath As String, ByRef canonicalExeName As String, ByRef canonicalName As String, ByRef regProgPath As String, ByRef CLSID As String, ByRef extension As String)
        'TODO for 4
        'See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/shell_adv/registeringapps.asp
		'Prerequisites:
		'    ProgID is already registered in registry.
		'    The 0th Icon of the exe file for the browser must be the program icon
		'Arguments:
		'    progID = program ID, e.g. AccessibleBrowser.CBrowser
		'    exePath = path to main executable, e.g. C:\Program Files\AB\AccessibleBrowser.exe
		'              No quotation marks.
		'    canonicalExeName = name of executable, e.g. AccessibleBrowser.exe
		'    canonicalName = name of program, e.g. "Accessible Browser"
		'    regProgPath = path to an exe file that can do the registration/unregistration process
		'                  i.e. call this module.
        '		On Error GoTo localHandler
        '		Dim result As Integer

        '		'Make sure recipient of this status is all set up for it.
        '		Call PrepareActiveXObject(progID, exePath, CLSID)

        '		'1 Register program in Web Browser list in Windows
        '		'See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/programmersguide/shell_adv/registeringapps.asp
        '		'Add canonical name and display name
        '		result = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName)
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName, "", canonicalName, REG_SZ)
        '		'3 Add icon to AB client entry
        '		result = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\DefaultIcon")
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\DefaultIcon", "", exePath & ",0", REG_SZ)
        '		'Add shell open verb
        '		result = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\shell")
        '		result = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\shell\open")
        '		result = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\shell\open\command")
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\shell\open\command", "", """" & exePath & """ ""%1""", REG_SZ)
        '		'Now InstallInfo
        '		result = CreateNewKey(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\InstallInfo")
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\InstallInfo", "ReinstallCommand", """" & regProgPath & """ /reinstall", REG_SZ)
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\InstallInfo", "HideIconsCommand", """" & regProgPath & """ /hideicons", REG_SZ)
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\InstallInfo", "ShowIconsCommand", """" & regProgPath & """ /showicons", REG_SZ)
        '		result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Clients\StartMenuInternet\" & canonicalExeName & "\InstallInfo", "IconsVisible", "1", REG_DWORD)
        '		Call ChangeHTTPHandler(progID, exePath)
        '		Exit Sub
        'localHandler: 
        '		'MsgBox Err.Number & " " & Err.description, vbCritical
	End Sub
	
	'
	'Public Sub ChangeExtensionHandler(extension As String, newProgramID As String)
	''change extension (e.g. "htm")
	'    On Error GoTo localHandler:
	'    Dim currentDefault As String
	'    Dim regResult As Long
	'
	'    'check for extension
	'    If Left(extension, 1) <> "." Then extension = "." & extension
	'    'check it already exists
	'    currentDefault = QueryValue(HKEY_CLASSES_ROOT, extension, "")
	'    If Len(currentDefault) > 0 Then
	'        'okay, that's fine
	'        regResult = SetKeyValue(HKEY_CLASSES_ROOT, extension, "", newProgramID, REG_SZ)
	'        If regResult <> 0 Then MsgBox "Error returned from registry write, code: " & regResult, vbOKOnly, "Error in SetKeyValue"
	'    End If
	'    Exit Sub
	'localHandler:
	'    MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error in registry access"
	'    Exit Sub
	'End Sub
	
	Private Sub ChangeHTTPHandler(ByRef progID As String, ByRef exePath As String)
        'Setup program as default Internet browser
        'TODO
        '		On Error GoTo localHandler
        '		'@ value of HTTP
        '		Call SetKeyValue(HKEY_CLASSES_ROOT, "HTTP", "", progID, REG_SZ)
        '		'Default icon
        '		Call SetKeyValue(HKEY_CLASSES_ROOT, "HTTP\DefaultIcon", "", exePath & ",0", REG_SZ)
        '		'And the actual important bit, shell
        '		Call SetKeyValue(HKEY_CLASSES_ROOT, "HTTP\shell\open\command", "", """" & exePath & """ %1", REG_SZ)
        '		Exit Sub
        'localHandler: 
        '		'MsgBox Err.Number & " " & Err.Description, vbOKOnly, "Error in setting default HTTP handler"
        '		Exit Sub
	End Sub
	
	Private Sub ChangeStartMenuButton(ByRef canonicalExeName As String, ByRef internationalResourcePath As String, ByRef internationalResourceIndex As Integer, ByRef exePath As String)
		'Puts your application in as the default web browser.f
        'See http://support.microsoft.com/?kbid=297878
        'TODO
        '		On Error GoTo localHandler
        '		Dim lRetVal As Integer
        '		Dim hKey As Integer

        '		lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\StartMenuInternet", 0, KEY_ALL_ACCESS, hKey)
        '		Call RegCloseKey(hKey)
        '		If lRetVal = ERROR_NONE Then
        '			'okay, we can manipulate the HKLM hive, use that:
        '			'Delete HKCU key
        '			lRetVal = DeleteKeyTree(HKEY_CURRENT_USER, "SOFTWARE\Clients\StartMenuInternet", False)
        '			lRetVal = DeleteKeyTree(HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\StartMenuInternet", False)
        '			'Commented out because some non-defined variables.
        '			'Call AddStartMenuEntries(HKEY_LOCAL_MACHINE, canonicalExeName, internationalResourcePath, internationalResourceIndex, iconResourcePath, iconIndex, exePath)
        '		Else
        '			'failed, use HKCU key instead
        '			lRetVal = SetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\Clients\StartMenuInternet", "", canonicalExeName, REG_SZ)
        '			If lRetVal = ERROR_NONE Then
        '				lRetVal = DeleteKeyTree(HKEY_CURRENT_USER, "SOFTWARE\Clients\StartMenuInternet", False)
        '				'Call AddStartMenuEntries(HKEY_CURRENT_USER, canonicalExeName, internationalResourcePath, internationalResourceIndex, iconResourcePath, iconIndex, exePath)
        '			End If
        '		End If
        '		Call modAPIDeclarations.SendMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Software\Clients\StartMenuInternet" & Chr(0))
        '		Exit Sub
        'localHandler: 
        '		'MsgBox Err.Number & " " & Err.description, vbOKOnly, "Error in changing menu button"
        '		Exit Sub
	End Sub
	
	Private Sub AddStartMenuEntries(ByRef lRootClass As Integer, ByRef canonicalExeName As String, ByRef internationalResourcePath As String, ByRef internationalResourceIndex As Integer, ByRef exePath As String)
		'adds necessary start menu entries into specified key (lRootClass)
		'Prerequisites: existing start menu entries have been removed (by DeleteKeyTree)
        'TODO
        'Call CreateNewKey(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName)
        'Call SetKeyValue(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName, "", internationalResourcePath & "," & internationalResourceIndex, REG_SZ)
        'Call CreateNewKey(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName & "\DefaultIcon")
        'Call SetKeyValue(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName & "\DefaultIcon", "", exePath & ",0", REG_SZ)
        'Call CreateNewKey(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName & "\Shell")
        'Call CreateNewKey(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName & "\Shell\Open")
        'Call CreateNewKey(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName & "\Shell\Open\Command")
        'Call SetKeyValue(lRootClass, "SOFTWARE\Clients\StartMenuInternet\" & canonicalExeName & "\Shell\Open\Command", "", exePath, REG_SZ)
	End Sub
	
	Private Function DeleteKeyTree(ByVal lRootClass As Integer, ByRef sKeyName As String, ByRef deleteThisKeyToo As Boolean) As Integer
        On Error Resume Next
        'deletes every descendant key of the key specified and (optionally)
        'this key too
        'by sKeyName or lParentKey

        Dim lRetVal As Integer ' results of API call: error_none=success
        Dim hKey As Integer ' handle to the key
        Dim ft As FILETIME
        Dim childName As String
        Dim className As String
        Dim lenChildName As Integer
        Dim lenClassName As Integer

        'See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcedata5/html/wce50lrfregopenkeyex.asp
        '1 Open this key
        lRetVal = RegOpenKeyEx(lRootClass, sKeyName, 0, KEY_ALL_ACCESS, hKey)
        If lRetVal = ERROR_NONE Then
            '2 Go through any children.
            Debug.Print(vbNewLine & "Opened key for deletion:" & sKeyName)
            childName = Space(255)
            className = Space(255)
            lenChildName = Len(childName)
            lenClassName = Len(className)
            lRetVal = RegEnumKeyEx(hKey, 0, childName, lenChildName, 0, className, lenClassName, ft)
            Debug.Print("GetChildResult=" & lRetVal)
            childName = Left(childName, lenChildName)
            On Error GoTo exitWhile
            While lRetVal = ERROR_NONE
                'okay, we have a child (because when we don't we get ERROR_NO_MORE_ITEMS
                'childName = Left(childName, Len(childName))
                Debug.Print("Child:[" & childName & "]")
                Call DeleteKeyTree(lRootClass, sKeyName & "\" & childName, True)
                'get the next key
                childName = Space(255)
                className = Space(255)
                lRetVal = RegEnumKeyEx(hKey, 0, childName, lenChildName, 0, className, Len(className), ft)
                childName = Left(childName, lenChildName)
                Debug.Print("Enumresult:" & lRetVal)
            End While
            DeleteKeyTree = lRetVal
exitWhile:

            '3 Delete this key
            If deleteThisKeyToo Then
                lRetVal = RegDeleteKey(lRootClass, sKeyName)
            End If
            Debug.Print("Delete result:" & lRetVal)
            '4 Close this key
            lRetVal = RegCloseKey(hKey)
        Else
            Debug.Print("Failed to open key for deletion: " & sKeyName & vbNewLine & "Result: " & lRetVal)
        End If
    End Function
	
	Private Function KeyExists(ByRef lRootClass As Integer, ByRef sKeyName As String) As Boolean
        On Error Resume Next
		Dim lRetVal As Integer
		Dim hKey As Integer
		
		lRetVal = RegOpenKeyEx(lRootClass, sKeyName, 0, KEY_QUERY_VALUE, hKey)
		KeyExists = (lRetVal = ERROR_NONE)
		lRetVal = RegCloseKey(hKey)
	End Function
	
	Public Sub MakeDefaultWebBrowser()

		'Makes WebbIE the default web browser
        'TODO
		'1 Register as default handler for text/html mime type.
        'result = SetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\MIME\Database\Content Type\text/html", "CLSID", "{EF9C0B4D-A14A-45f5-8DD3-BD49D5FC6D1C}", REG_SZ)
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "MIME\Database\Content Type\text/html", "CLSID", "{EF9C0B4D-A14A-45f5-8DD3-BD49D5FC6D1C}", REG_SZ)
        ''http
        ''UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "HTTP\DefaultIcon", "", My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe,0", REG_SZ)
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "HTTP\shell", "", "open", REG_SZ)
        ''UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "HTTP\shell\open\command", "", My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & " ""%1""", REG_SZ)
        'result = RegDeleteKey(HKEY_CLASSES_ROOT, "HTTP\shell\open\ddeexec")
        ''https
        ''UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "HTTPS\DefaultIcon", "", My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe,0", REG_SZ)
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "HTTPS\shell", "", "open", REG_SZ)
        ''UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'result = SetKeyValue(HKEY_CLASSES_ROOT, "HTTPS\shell\open\command", "", My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & " ""%1""", REG_SZ)
        'result = RegDeleteKey(HKEY_CLASSES_ROOT, "HTTPS\shell\open\ddeexec")
        ''2 Make it default handler for htm/html file type.
        'result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Classes\.html", "", "WebbIE.Application", REG_SZ)
        'result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Classes\.html\PersistentHandler", "", "{EF9C0B4D-A14A-45f5-8DD3-BD49D5FC6D1C}", REG_SZ)
        'result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Classes\.htm", "", "WebbIE.Application", REG_SZ)
        'result = SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Classes\.htm\PersistentHandler", "", "{EF9C0B4D-A14A-45f5-8DD3-BD49D5FC6D1C}", REG_SZ)
        ''3 Register for Program Access and Defaults.
        ''Try for Local Machine:
        'result = SetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\StartMenuInternet", "", "WebbIE3.exe", REG_SZ)
        ''If this didn't work, do Current User:
        'If result <> ERROR_NONE Then
        '	result = SetKeyValue(HKEY_CURRENT_USER, "SOFTWARE\Clients\StartMenuInternet", "", "WebbIE3.exe", REG_SZ)
        'End If
	End Sub
End Module