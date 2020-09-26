Imports System
Imports System.Collections.Generic
Imports System.Text

''' <summary>
''' Helper functions  for using the Microsoft WebBrowser control
''' </summary>
''' <remarks>Last update 3 March 2014 for IE11</remarks>
Public Class WebBrowserUtility

    ''' <summary>
    ''' Values for FEATURE_BROWSER_EMULATION
    ''' </summary>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ee330730(v=vs.85).aspx</remarks>
    Public Enum InternetExplorerEmulationMode
        IE11NoMatterWhat = 11001
        IE11IfPageHasCorrectDoctype = 11000
        IE10NoMatterWhat = 10001
        IE10IfPageHasCorrectDoctype = 10000
        IE9NoMatterWhat = 9999
        IE9IfPageHasCorrectDoctype = 9000
        IE8NoMatterWhat = 8888
        IE8IfPageHasCorrectDoctype = 8000
        IE7IfPageHasCorrectDoctypeDEFAULT = 7000
        BestIEForOS = 0
    End Enum

    Private Const MAJOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN As Integer = 6
    Private Const MINOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN As Integer = 2

    ''' <summary>
    ''' Writes the HKEY_CURRENT_USER registry key necessary for any WebBrowser components used in this
    ''' project to emulate a particular version of Internet Explorer. If you don't do this then any
    ''' WebBrowser objects will run as IE7. 
    ''' http://msdn.microsoft.com/en-us/library/ee330730(v=VS.85).aspx
    ''' http://msdn.microsoft.com/en-us/library/ee330730%28VS.85%29.aspx#browser_emulation
    ''' http://stackoverflow.com/questions/4612255/regarding-ie9-webbrowser-control
    ''' </summary>
    ''' <param name="ieMode">The IE mode you want. Note that "best" is "best at the time this code is written". So in January 2014
    ''' "best" is "IE11". When IE12 comes out, if this code hasn't been updated, then, you'll not get the latest IE.</param>
    ''' <param name="applyOnFutureVersions">If true then your IE mode will be applied, even if you are using a version of Windows
    ''' that was not extant when the code was written and the logic has therefore not been checked. You probably want this.</param>
    ''' <remarks></remarks>
    Public Shared Sub SetWebBrowserEmulation(Optional ieMode As InternetExplorerEmulationMode = InternetExplorerEmulationMode.BestIEForOS, Optional applyOnFutureVersions As Boolean = True)

        Dim majorVersion As Integer = System.Environment.OSVersion.Version.Major
        Dim minorVersion As Integer = System.Environment.OSVersion.Version.Minor
        ' Windows 2000 is version 5, Windows XP is version 5.1.
        ' Vista is 6, Windows 7 is 6.1, Windows 8 is 6.2.
        ' IE9 and IE10 are not available on Windows XP or earlier.
        ' IE10 is not available on Windows Vista or earlier
        ' If the version number is greater than X, where X = the versions released when this code was last
        ' updated, then the setting will only be set if applyOnFutureVersions is true.
        If ieMode = InternetExplorerEmulationMode.IE11IfPageHasCorrectDoctype Or ieMode = InternetExplorerEmulationMode.IE11NoMatterWhat Then
            If (majorVersion < 6 And minorVersion < 1) Then
                Throw New ArgumentException("Cannot request IE11 on Windows versions before Windows 7")
            End If
        ElseIf ieMode = InternetExplorerEmulationMode.IE10IfPageHasCorrectDoctype Or ieMode = InternetExplorerEmulationMode.IE10NoMatterWhat Then
            If (majorVersion < 6 And minorVersion < 1) Then
                Throw New ArgumentException("Cannot request IE10 on Windows versions before Windows 7")
            End If
        ElseIf ieMode = InternetExplorerEmulationMode.IE9IfPageHasCorrectDoctype Or ieMode = InternetExplorerEmulationMode.IE9NoMatterWhat Then
            If (majorVersion < 6) Then
                Throw New ArgumentException("Cannot request IE9 on Windows versions before Windows Vista")
            End If
        ElseIf (majorVersion > MAJOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN And Not applyOnFutureVersions) Then
            ' Do not set. Later version of Windows. 
        ElseIf majorVersion = MAJOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN And minorVersion > MINOR_WINDOWS_VERSION_WHEN_THIS_CODE_WAS_WRITTEN And Not applyOnFutureVersions Then
            ' Do not set. Later version of Windows.
        Else
            ' OK, checked versions, we're good!
            Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", True)
            If regKey Is Nothing Then
                Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl")
                regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION")
            End If
            Dim appExeName As String = AppDomain.CurrentDomain.FriendlyName

            ' What do we write? Depends on ieMode.
            If ieMode = InternetExplorerEmulationMode.BestIEForOS Then
                If (majorVersion < 6) Then
                    ' Windows XP or earlier: Maximum Internet Explorer 8
                    Call regKey.SetValue(appExeName, InternetExplorerEmulationMode.IE8NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord)
                    If appExeName.Contains(".vshost") Then
                        Call regKey.SetValue(appExeName.Replace(".vshost", ""), InternetExplorerEmulationMode.IE8NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                ElseIf (majorVersion = 6 And minorVersion < 2) Then
                    ' Windows Vista: Maximum Internet Explorer 9
                    Call regKey.SetValue(appExeName, InternetExplorerEmulationMode.IE9NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord)
                    If appExeName.Contains(".vshost") Then
                        Call regKey.SetValue(appExeName.Replace(".vshost", ""), InternetExplorerEmulationMode.IE9NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                Else
                    ' Windows 7 or later: at time of writing, maximum Internet Explorer 11
                    Call regKey.SetValue(appExeName, InternetExplorerEmulationMode.IE11NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord)
                    If appExeName.Contains(".vshost") Then
                        Call regKey.SetValue(appExeName.Replace(".vshost", ""), InternetExplorerEmulationMode.IE11NoMatterWhat, Microsoft.Win32.RegistryValueKind.DWord)
                    End If
                End If
            Else
                Call regKey.SetValue(appExeName, ieMode, Microsoft.Win32.RegistryValueKind.DWord)
                If appExeName.Contains(".vshost") Then
                    Call regKey.SetValue(appExeName.Replace(".vshost", ""), ieMode, Microsoft.Win32.RegistryValueKind.DWord)
                End If
            End If
            Call regKey.Close()
        End If
    End Sub
End Class
