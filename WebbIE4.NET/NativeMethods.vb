Friend Class NativeMethods

    Friend Declare Auto Function SetCursorPos Lib "User32.dll" (ByVal X As Integer, ByVal Y As Integer) As Long
    Friend Declare Auto Function GetCursorPos Lib "User32.dll" (ByRef lpPoint As Point) As Long
    Friend Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As UInteger, ByVal dx As UInteger, ByVal dy As UInteger, ByVal cButtons As UInteger, ByVal dwExtraInfo As IntPtr)
    Friend Const MOUSEEVENTF_LEFTDOWN As UInteger = &H2 ' left button down
    Friend Const MOUSEEVENTF_LEFTUP As UInteger = &H4 ' left button up
    Friend Const MOUSEEVENTF_MIDDLEDOWN As UInteger = &H20 ' middle button down
    Friend Const MOUSEEVENTF_MIDDLEUP As UInteger = &H40 ' middle button up
    Friend Const MOUSEEVENTF_RIGHTDOWN As UInteger = &H8 ' right button down
    Friend Const MOUSEEVENTF_RIGHTUP As UInteger = &H10 ' right button up

    Friend Const EM_SETMODIFY As Integer = &HB9
    Friend Declare Function SetFocus Lib "user32" (ByVal hWnd As IntPtr) As IntPtr
    ''' <summary>
    ''' Used to work out if a key is pressed or not.
    ''' </summary>
    ''' <param name="vKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto, CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall)>
    Friend Shared Function GetKeyState(ByVal vKey As Integer) As Short
    End Function

    'URLDownloadToFile
    'Downloads a specified url to a local file
    Friend Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Integer, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Integer, ByVal lpfnCB As Integer) As Integer

    'For removing a menu item
    Friend Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer

    'For finding the system time - used for timing
    Friend Declare Function GetTickCount Lib "kernel32" () As Integer

    'For finding IE version
    'UPGRADE_WARNING: Structure DllVersionInfo may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Friend Declare Function DllGetVersion Lib "Shlwapi.dll" (ByRef dwVersion As DllVersionInfo) As Integer
    'for finding IE version

    Friend Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Integer, ByRef lpszConnectionName As Integer, ByRef dwNameLen As Integer, ByVal dwReserved As Integer) As Integer

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Friend Shared Function SendMessageW(ByVal hwnd As IntPtr, ByVal wMsg As UInt32, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Friend Declare Function FolderRegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
    ' Note that if you declare the lpData parameter as String, you must pass it
    ' By Value.

    'Friend Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    'must be called before FolderRegQueryEx
    'DEV RegOpenKey is deprecated for RegOpenKeyEx, and I think it's giving me grief with UAC in Vista.

    'for processing the menu favorites
    Friend Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer

    Friend Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Integer, ByVal hwnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Friend Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Integer

    'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Friend Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Integer

    'UPGRADE_WARNING: Structure MENUITEMINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Friend Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Integer, ByVal un As Integer, ByVal b As Integer, ByRef lpMenuItemInfo As MENUITEMINFO) As Integer

    Friend Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer

    Friend Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Integer

    'Friend Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

    'for getting the locale ID for the current machine/user
    Friend Declare Function GetUserDefaultLCID Lib "kernel32" () As Integer

    'Used to determine the (window handle of) the current active window in Windows.
    Friend Declare Function GetForegroundWindow Lib "user32" () As Integer
    'CONSTANTS
    Friend Const HKEY_CLASSES_ROOT As Integer = &H80000000
    Friend Const HKEY_CURRENT_USER As Integer = &H80000001
    Friend Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Friend Const HKEY_USERS As Integer = &H80000003
    Friend Const HKEY_CURRENT_CONFIG As Integer = &H80000005
    Friend Const HKEY_DYN_DATA As Integer = &H80000006
    Friend Const REG_SZ As Short = 1 'Unicode nul terminated string
    Friend Const REG_BINARY As Short = 3 'Free form binary
    Friend Const REG_DWORD As Short = 4 '32-bit number
    Friend Const ERROR_SUCCESS As Short = 0
    Friend Const RESERVED_NULL As Short = 0
    Friend Const INTERNET_CONNECTION_MODEM As Integer = &H1
    Friend Const MIIM_SUBMENU As Integer = &H4
    Friend Const APINULL As Short = 0
    Friend Const EM_SCROLLCARET As Integer = &HB7
    Friend Const EM_LINEINDEX As Integer = &HBB
    Friend Const EM_GETLINECOUNT As Integer = &HBA
    Friend Const EM_LINELENGTH As Integer = &HC1
    Friend Const EM_GETFIRSTVISIBLELINE As Integer = &HCE
    Friend Const EM_LINEFROMCHAR As Integer = &HC9
    Friend Const EM_GETLINE As Integer = &HC4
    Friend Const EM_LINESCROLL As Integer = &HB6
    Friend Const EM_SETSEL As Integer = &HB1
    Friend Const KEY_READ As Integer = &H20019 '-- Permission for general read access.
    Friend Const KEY_QUERY_VALUE As Integer = &H1
    Friend Const KEY_SET_VALUE As Integer = &H2
    Friend Const KEY_ALL_ACCESS As Integer = &H3F

    Friend Const MF_MENUBREAK As Integer = &H40 'indicates a vertical break
    Friend Const MIIM_STATE As Integer = &H1
    Friend Const MIIM_ID As Integer = &H2
    Friend Const MIIM_CHECKMARKS As Integer = &H8
    Friend Const MIIM_TYPE As Integer = &H10
    Friend Const MIIM_DATA As Integer = &H20
    Friend Const MIIM_STRING As Integer = &H40
    Friend Const MIIM_BITMAP As Integer = &H80
    Friend Const MIIM_FTYPE As Integer = &H100
    Friend Const FBYPOSITION_POSITION As Boolean = True
    Friend Const FBYPOSTION_IDENTIFIER As Boolean = False
    Friend Const MF_STRING As Integer = &H0
    Friend Const GWL_WNDPROC As Integer = (-4)
    Friend Const WM_COMMAND As Integer = &H111 ' indicates that a command has been intercepted by the app
    Friend Const MF_BYPOSITION As Integer = &H400 ' indicates a menu item by position, not by name, for RemoveMenu
    Friend Const MF_DISABLED As Integer = &H2 ' used in setting menu items (InsertMenuItem etc) to indicate that an item is greyed
    Friend Const MF_GRAYED As Integer = &H1 ' allegedly does the same thing as MF_DISABLED

    'TYPES
    Friend Structure DllVersionInfo
        Dim cbSize As Integer
        Dim dwMajorVersion As Integer
        Dim dwMinorVersion As Integer
        Dim dwBuildNumber As Integer
        Dim dwPlatformId As Integer
    End Structure

    Friend Structure MENUITEMINFO
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

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
    Friend Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Friend Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Friend Const SW_SHOWNORMAL As Integer = 1

    Friend Declare Ansi Function GetPrivateProfileString _
      Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal lpDefault As String,
      ByVal lpReturnedString As System.Text.StringBuilder,
      ByVal nSize As Integer, ByVal lpFileName As String) _
      As Integer

    Friend Declare Ansi Function WritePrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal lpString As String,
      ByVal lpFileName As String) As Integer
    Friend Declare Ansi Function GetPrivateProfileInt _
      Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal nDefault As Integer,
      ByVal lpFileName As String) As Integer
    Friend Declare Ansi Function FlushPrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer,
      ByVal lpKeyName As Integer, ByVal lpString As Integer,
      ByVal lpFileName As String) As Integer


    Friend Const VK_CAPITAL As Integer = &H14
    Friend Const KEYEVENTF_EXTENDEDKEY As Integer = &H1
    Friend Const KEYEVENTF_KEYUP As Integer = &H2
    Friend Const VER_PLATFORM_WIN32_NT As Short = 2
    Friend Const VER_PLATFORM_WIN32_WINDOWS As Short = 1

    'SHGetSpecialFolderLocation
    'Returns the Folder ID of the user's My Documents folder (or another folder indicated
    'by CSIDL)
    Friend Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwnd As Integer, ByVal nFolder As Integer, ByRef ppidl As Integer) As Integer
    'SHGetPathFromIDList
    'Returns the path (string) from the folder ID obtained by SHGetSpecialFolderLocation
    Friend Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal Pidl As Integer, ByVal pszPath As String) As Integer

    ' constants for Shell.NameSpace method -- these are the "special folders"
    'contained in the Windows shell
    Friend Const CSIDL_DESKTOP As Integer = &H0 ' Desktop
    Friend Const CSIDL_INTERNET As Integer = &H1 ' The internet
    Friend Const CSIDL_PROGRAMS As Integer = &H2 ' Shortcuts in the Programs menu
    Friend Const CSIDL_CONTROLS As Integer = &H3 ' Control Panel
    Friend Const CSIDL_PRINTERS As Integer = &H4 ' Printers
    Friend Const CSIDL_PERSONAL As Integer = &H5 ' Shortcuts to Personal files
    Friend Const CSIDL_FAVORITES As Integer = &H6 ' Shortcuts to favorite folders
    Friend Const CSIDL_STARTUP As Integer = &H7 ' Shortcuts to apps that start at boot Time
    Friend Const CSIDL_RECENT As Integer = &H8 ' Shortcuts to recently used docs
    Friend Const CSIDL_SENDTO As Integer = &H9 ' Shortcuts for the SendTo menu
    Friend Const CSIDL_BITBUCKET As Integer = &HA ' Recycle Bin
    Friend Const CSIDL_STARTMENU As Integer = &HB ' User-defined items in Start Menu
    Friend Const CSIDL_DESKTOPDIRECTORY As Integer = &H10 ' Directory with all the desktop shortcuts
    Friend Const CSIDL_DRIVES As Integer = &H11 ' My Computer
    Friend Const CSIDL_NETWORK As Integer = &H12 ' Network Neighborhood virtual folder
    Friend Const CSIDL_NETHOOD As Integer = &H13 ' Directory containing objects in the network neighborhood
    Friend Const CSIDL_FONTS As Integer = &H14 ' Installed fonts
    Friend Const CSIDL_TEMPLATES As Integer = &H15 ' Shortcuts to document templates
    Friend Const CSIDL_COMMON_STARTMENU As Integer = &H16 ' Directory with items in the Start menu for all users
    Friend Const CSIDL_COMMON_PROGRAMS As Integer = &H17 ' Directory with items in the Programs menu for all users
    Friend Const CSIDL_COMMON_STARTUP As Integer = &H18 ' Directory with items in the StartUp submenu for all users
    Friend Const CSIDL_COMMON_DESKTOPDIRECTORY As Integer = &H19 ' Directory with items on the desktop of all users
    Friend Const CSIDL_APPDATA As Integer = &H1A ' Folder for application-specific data
    Friend Const CSIDL_PRINTHOOD As Integer = &H1B ' Directory with references to printer links
    Friend Const CSIDL_LOCAL_APPDATA As Integer = &H1C '{user name}\Local Settings\Application Data (non roaming)
    Friend Const CSIDL_ALTSTARTUP As Integer = &H1D ' (DBCS) Directory corresponding to user 's nonlocalized Startup program group
    Friend Const CSIDL_COMMON_ALTSTARTUP As Integer = &H1E ' (DBCS) Directory with Startup items for all users
    Friend Const CSIDL_COMMON_FAVORITES As Integer = &H1F ' Directory with all user's favorit items
    Friend Const CSIDL_INTERNET_CACHE As Integer = &H20 ' Directory for temporary internet Files
    Friend Const CSIDL_COOKIES As Integer = &H21 ' Directory for Internet cookies
    Friend Const CSIDL_HISTORY As Integer = &H22 ' Directory for Internet history items
    Friend Const CSIDL_COMMON_APPDATA As Integer = &H23 'All Users\Application Data
    Friend Const CSIDL_WINDOWS As Integer = &H24 'GetWindowsDirectory()
    Friend Const CSIDL_SYSTEM As Integer = &H25 'GetSystemDirectory()
    Friend Const CSIDL_PROGRAM_FILES As Integer = &H26 'C:\Program Files
    Friend Const CSIDL_MYPICTURES As Integer = &H27 'C:\Program Files\My Pictures
    Friend Const CSIDL_PROFILE As Integer = &H28 'USERPROFILE
    Friend Const CSIDL_SYSTEMX86 As Integer = &H29 'x86 system directory on RISC
    Friend Const CSIDL_PROGRAM_FILESX86 As Integer = &H2A 'x86 C:\Program Files on RISC
    Friend Const CSIDL_PROGRAM_FILES_COMMON As Integer = &H2B 'C:\Program Files\Common
    Friend Const CSIDL_PROGRAM_FILES_COMMONX86 As Integer = &H2C 'x86 Program Files\Common on RISC
    Friend Const CSIDL_COMMON_TEMPLATES As Integer = &H2D 'All Users\Templates
    Friend Const CSIDL_COMMON_DOCUMENTS As Integer = &H2E 'All Users\Documents
    Friend Const CSIDL_COMMON_ADMINTOOLS As Integer = &H2F 'All Users\Start Menu\Programs\Administrative Tools
    Friend Const CSIDL_ADMINTOOLS As Integer = &H30 '{user name}\Start Menu\Programs\Administrative Tools

    Friend Const CSIDL_FLAG_CREATE As Integer = &H8000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
    Friend Const CSIDL_FLAG_DONT_VERIFY As Integer = &H4000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
    Friend Const CSIDL_FLAG_MASK As Integer = &HFF00 'mask for all possible flag values

    Friend Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    'open a key for reading or writing

    <System.Runtime.InteropServices.DllImport("user32.dll", CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall,
           CharSet:=System.Runtime.InteropServices.CharSet.Unicode, EntryPoint:="keybd_event",
           ExactSpelling:=True, SetLastError:=True)>
    Friend Shared Function keybd_event(ByVal bVk As Byte, ByVal bScan As Byte,
                              ByVal dwFlags As UInt32, ByVal dwExtraInfo As IntPtr) As Boolean
    End Function

    Friend Const VK_BACK As Int32 = &H8
    Friend Const VK_TAB As Int32 = &H9
    ' &HA-&HB are unassigned, according to http://blogs.msdn.com/michkap/archive/2006/03/23/558658.aspx
    Friend Const VK_CLEAR As Int32 = &HC
    Friend Const VK_RETURN As Int32 = &HD
    ' &HE and &F - don't know!
    Friend Const VK_SHIFT As Int32 = &H10
    Friend Const VK_CONTROL As Int32 = &H11
    Friend Const VK_ALTER As Int32 = &H12
    Friend Const VK_Pause As Int32 = &H13
    Friend Const VK_CapsLock As Int32 = &H14
    Friend Const VK_KANA As Int32 = &H15
    Friend Const VK_HANGEUL As Int32 = &H15
    Friend Const VK_JUNJA As Int32 = &H17 ' 22
    Friend Const VK_FINAL As Int32 = &H18 ' 23
    Friend Const VK_HANJA As Int32 = &H19 ' 24
    Friend Const VK_KANJI As Int32 = &H19 ' 25
    '&H1A - don't know. 26
    Friend Const VK_Escape As Int32 = &H1B ' 27
    Friend Const VK_CONVERT As Int32 = &H1C ' 28
    Friend Const VK_NONCONVERT As Int32 = &H1D ' 29
    Friend Const VK_ACCEPT As Int32 = &H1E ' 30
    Friend Const VK_MODECHANGE As Int32 = &H1F ' 31
    Friend Const VK_Space As Int32 = &H20 ' 32
    Friend Const VK_PRIOR As Int32 = &H21 'Keys.PageUp
    Friend Const VK_NEXT As Int32 = &H22 '  Keys.PageDown
    Friend Const VK_End As Int32 = &H23
    Friend Const VK_Home As Int32 = &H24
    Friend Const VK_Left As Int32 = &H25
    Friend Const VK_Up As Int32 = &H26
    Friend Const VK_RIGHT As Int32 = &H27
    Friend Const VK_Down As Int32 = &H28
    Friend Const VK_SELECT As Int32 = &H29
    Friend Const VK_PRINT As Int32 = &H2A
    Friend Const VK_EXECUTE As Int32 = &H2B
    Friend Const VK_SNAPSHOT As Int32 = &H2C ' Keys.PrintScreen
    Friend Const VK_Insert As Int32 = &H2D
    Friend Const VK_Delete As Int32 = &H2E
    Friend Const VK_HELP As Int32 = &H2F
    Friend Const VK_0 As Int32 = &H30
    Friend Const VK_1 As Int32 = &H31
    Friend Const VK_2 As Int32 = &H32
    Friend Const VK_3 As Int32 = &H33
    Friend Const VK_4 As Int32 = &H34
    Friend Const VK_5 As Int32 = &H35
    Friend Const VK_6 As Int32 = &H36
    Friend Const VK_7 As Int32 = &H37
    Friend Const VK_8 As Int32 = &H38
    Friend Const VK_9 As Int32 = &H39
    ' &H40 : unassigned
    Friend Const VK_A As Int32 = &H41
    Friend Const VK_B As Int32 = &H42
    Friend Const VK_C As Int32 = &H43
    Friend Const VK_D As Int32 = &H44
    Friend Const VK_E As Int32 = &H45
    Friend Const VK_F As Int32 = &H46
    Friend Const VK_G As Int32 = &H47
    Friend Const VK_H As Int32 = &H48
    Friend Const VK_I As Int32 = &H49
    Friend Const VK_J As Int32 = &H4A
    Friend Const VK_K As Int32 = &H4B
    Friend Const VK_L As Int32 = &H4C
    Friend Const VK_M As Int32 = &H4D
    Friend Const VK_N As Int32 = &H4E
    Friend Const VK_O As Int32 = &H4F
    Friend Const VK_P As Int32 = &H50
    Friend Const VK_Q As Int32 = &H51
    Friend Const VK_R As Int32 = &H52
    Friend Const VK_S As Int32 = &H53
    Friend Const VK_T As Int32 = &H54
    Friend Const VK_U As Int32 = &H55
    Friend Const VK_V As Int32 = &H56
    Friend Const VK_W As Int32 = &H57
    Friend Const VK_X As Int32 = &H58
    Friend Const VK_Y As Int32 = &H59
    Friend Const VK_Z As Int32 = &H5A
    Friend Const VK_LWIN As Int32 = &H5B
    Friend Const VK_RWIN As Int32 = &H5C
    Friend Const VK_APPS As Int32 = &H5D
    ' &H5E : reserved
    Friend Const VK_SLEEP As Int32 = &H5 ' Keys.Sleep
    Friend Const VK_NUMPAD0 As Int32 = &H60
    Friend Const VK_NUMPAD1 As Int32 = &H61
    Friend Const VK_NUMPAD2 As Int32 = &H62
    Friend Const VK_NUMPAD3 As Int32 = &H63
    Friend Const VK_NUMPAD4 As Int32 = &H64
    Friend Const VK_NUMPAD5 As Int32 = &H65
    Friend Const VK_NUMPAD6 As Int32 = &H66
    Friend Const VK_NUMPAD7 As Int32 = &H67
    Friend Const VK_NUMPAD8 As Int32 = &H68
    Friend Const VK_NUMPAD9 As Int32 = &H69
    Friend Const VK_MULTIPLY As Int32 = &H6A '*
    Friend Const VK_ADD As Int32 = &H6B '+
    Friend Const VK_SEPARATOR As Int32 = &H6C
    Friend Const VK_SUBTRACT As Int32 = &H6D '-
    Friend Const VK_DECIMAL As Int32 = &H6E '.
    Friend Const VK_DIVIDE As Int32 = &H6F '/
    Friend Const VK_F1 As Int32 = &H70
    Friend Const VK_F2 As Int32 = &H71
    Friend Const VK_F3 As Int32 = &H72
    Friend Const VK_F4 As Int32 = &H73
    Friend Const VK_F5 As Int32 = &H74
    Friend Const VK_F6 As Int32 = &H75
    Friend Const VK_F7 As Int32 = &H76
    Friend Const VK_F8 As Int32 = &H77
    Friend Const VK_F9 As Int32 = &H78
    Friend Const VK_F10 As Int32 = &H79
    Friend Const VK_F11 As Int32 = &H7A
    Friend Const VK_F12 As Int32 = &H7B
    Friend Const VK_F13 As Int32 = &H7C
    Friend Const VK_F14 As Int32 = &H7D
    Friend Const VK_F15 As Int32 = &H7E
    Friend Const VK_F16 As Int32 = &H7F
    Friend Const VK_F17 As Int32 = &H80
    Friend Const VK_F18 As Int32 = &H81
    Friend Const VK_F19 As Int32 = &H82
    Friend Const VK_F20 As Int32 = &H83
    Friend Const VK_F21 As Int32 = &H84
    Friend Const VK_F22 As Int32 = &H85
    Friend Const VK_F23 As Int32 = &H86
    Friend Const VK_F24 As Int32 = &H87
    ' &H88 - &H8F : unassigned
    Friend Const VK_NUMLOCK As Int32 = &H90
    Friend Const VK_SCROLL As Int32 = &H91
    Friend Const VK_OEM_NEC_EQUAL As Int32 = &H92        ',                     ' &H92, NEC PC-9800 kbd definition
    Friend Const VK_OEM_FJ_JISHO As Int32 = &H92 ',                     ' &H92, Fujitsu/OASYS kbd definition
    Friend Const VK_OEM_FJ_MASSHOU As Int32 = &H93 ',                     ' &H93, Fujitsu/OASYS kbd definition
    Friend Const VK_OEM_FJ_TOUROKU As Int32 = &H94 ',                     ' &H94, Fujitsu/OASYS kbd definition
    Friend Const VK_OEM_FJ_LOYA As Int32 = &H95 ',                     ' &H95, Fujitsu/OASYS kbd definition
    Friend Const VK_OEM_FJ_ROYA As Int32 = &H96 ',                     ' &H96, Fujitsu/OASYS kbd definition
    ' &H97 - &H9F : unassigned
    Friend Const VK_LSHIFT As Int32 = &HA0 ' 160
    Friend Const VK_RSHIFT As Int32 = &HA1 ' 161
    Friend Const VK_LCONTROL As Int32 = &HA2 ' 162
    Friend Const VK_RCONTROL As Int32 = &HA3 ' 163
    Friend Const VK_LMENU As Int32 = &HA4 ' 164
    Friend Const VK_RMENU As Int32 = &HA5
    Friend Const VK_BROWSER_BACK As Int32 = &HA6
    Friend Const VK_BROWSER_FORWARD As Int32 = &HA7
    Friend Const VK_BROWSER_REFRESH As Int32 = &HA8
    Friend Const VK_BROWSER_STOP As Int32 = &HA9
    Friend Const VK_BROWSER_SEARCH As Int32 = &HAA
    Friend Const VK_BROWSER_FAVORITES As Int32 = &HAB
    Friend Const VK_BROWSER_HOME As Int32 = &HAC
    Friend Const VK_VOLUME_MUTE As Int32 = &HAD
    Friend Const VK_VOLUME_DOWN As Int32 = &HAE
    Friend Const VK_VOLUME_UP As Int32 = &HAF
    Friend Const VK_MEDIA_NEXT_TRACK As Int32 = &HB0
    Friend Const VK_MEDIA_PREV_TRACK As Int32 = &HB1
    Friend Const VK_MEDIA_STOP As Int32 = &HB2
    Friend Const VK_MEDIA_PLAY_PAUSE As Int32 = &HB3
    Friend Const VK_LAUNCH_MAIL As Int32 = &HB4
    Friend Const VK_LAUNCH_MEDIA_SELECT As Int32 = &HB5
    Friend Const VK_LAUNCH_APP1 As Int32 = &HB6
    Friend Const VK_LAUNCH_APP2 As Int32 = &HB7
    ' &HB8 - &HB9 : reserved
    Friend Const VK_OEM_1 As Int32 = &HBA ' Keys.Oem1
    Friend Const VK_OEM_PLUS As Int32 = &HBB
    Friend Const VK_OEM_COMMA As Int32 = &HBC
    Friend Const VK_OEM_MINUS As Int32 = &HBD
    Friend Const VK_OEM_PERIOD As Int32 = &HBE
    Friend Const VK_OEM_2 As Int32 = &HBF
    Friend Const VK_OEM_3 As Int32 = &HC0           ' Keys.Oem3
    ' &HC1 - &HD7 : reserved
    ' &HD8 - &HDA : unassigned
    Friend Const VK_OEM_4 As Int32 = &HDB
    Friend Const VK_OEM_5 As Int32 = &HDC
    Friend Const VK_OEM_6 As Int32 = &HDD
    Friend Const VK_OEM_7 As Int32 = &HDE
    Friend Const VK_OEM_8 As Int32 = &HDF
    '&HE0 : reserved
    Friend Const VK_OEM_AX As Int32 = &HE1               '                     &HE1 ' 'AX' key on Japanese AX kbd
    Friend Const VK_OEM_102 As Int32 = &HE2        ' Keys.Oem102
    Friend Const VK_ICO_HELP As Int32 = &HE3 ' Help key on ICO
    Friend Const VK_ICO_00 As Int32 = &HE4 ' 00 key on ICO
    Friend Const VK_PROCESSKEY As Int32 = &HE5
    Friend Const VK_ICO_CLEAR As Int32 = &HE6            ' &HE6
    Friend Const VK_PACKET As Int32 = &HE7 '                     &HE7 ' Keys.Packet
    ' &HE8 : unassigned
    Friend Const VK_OEM_RESET As Int32 = &HE9 ' Nokia/Ericsson definition
    Friend Const VK_OEM_JUMP As Int32 = &HEA ' Nokia/Ericsson definition
    Friend Const VK_OEM_PA1 As Int32 = &HEB ' Nokia/Ericsson definition
    Friend Const VK_OEM_PA2 As Int32 = &HEC ' Nokia/Ericsson definition
    Friend Const VK_OEM_PA3 As Int32 = &HED ' Nokia/Ericsson definition
    Friend Const VK_OEM_WSCTRL As Int32 = &HEE ' Nokia/Ericsson definition
    Friend Const VK_OEM_CUSEL As Int32 = &HEF ' Nokia/Ericsson definition
    Friend Const VK_OEM_ATTN As Int32 = &HF0 ' Nokia/Ericsson definition
    Friend Const VK_OEM_FINISH As Int32 = &HF1 ' Nokia/Ericsson definition
    Friend Const VK_OEM_COPY As Int32 = &HF2 ' Nokia/Ericsson definition
    Friend Const VK_OEM_AUTO As Int32 = &HF3 ' Nokia/Ericsson definition
    Friend Const VK_OEM_ENLW As Int32 = &HF4 ' Nokia/Ericsson definition
    Friend Const VK_OEM_BACKTAB As Int32 = &HF5 ' Nokia/Ericsson definition
    Friend Const VK_ATTN As Int32 = &HF6
    Friend Const VK_CRSEL As Int32 = &HF7
    Friend Const VK_EXSEL As Int32 = &HF8
    Friend Const VK_EREOF As Int32 = &HF9
    Friend Const VK_PLAY As Int32 = &HFA
    Friend Const VK_ZOOM As Int32 = &HFB
    Friend Const VK_NONAME As Int32 = &HFC
    Friend Const VK_PA1 As Int32 = &HFD
    Friend Const VK_OEM_CLEAR As Int32 = &HFE
    ' &HFF : reserved



End Class
