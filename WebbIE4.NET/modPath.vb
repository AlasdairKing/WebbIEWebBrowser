Option Strict On
Option Explicit On
Module modPath
	
	'Copyright (c) 2007, Alasdair King
	'All rights reserved.
	'
	'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
	'
	'    * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
	'    * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
	'    * Neither the name of [Alasdair] nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
	'
	'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	
	'Changelog
	'       28 Aug 2009             Made GetAppPath return String explicitly.
	
	Public settingsPath As String ' the settings directory for the application
	Public runningLocal As Boolean ' whether we're running off the local folder
	'only
	Public commonSettingsPath As String ' the settings directory for the WebbIE applications
	Public nonRoamingSettingsPath As String ' the settings directory for non-roaming data
    '(e.g. Local Settings)


    Public Sub DetermineSettingsPath(Optional ByVal companyName As String = "", Optional ByVal productName As String = "", Optional ByVal version As String = "")
        Try
            'works out whether we are running on a memory stick/standalone or as an
            'installed application.
            Dim key As String
            Dim section As String
            Dim Path As String
            Dim got As String

            'Get any override values for company, version and application name from the program .ini file, if any.
            'Otherwise use the values passed.
            'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            companyName = modIniFile.GetString("ApplicationDataPath", "Company", companyName, GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini")
            If Len(companyName) = 0 Then
                companyName = My.Application.Info.CompanyName
            End If
            'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            productName = modIniFile.GetString("ApplicationDataPath", "ApplicationName", productName, GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini")
            If Len(productName) = 0 Then
                productName = My.Application.Info.ProductName
            End If
            'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            version = modIniFile.GetString("ApplicationDataPath", "Version", version, GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini")
            If Len(version) = 0 Then
                version = CStr(My.Application.Info.Version.Major)
            End If

            If System.IO.File.Exists(GetAppPath() & "\installed.ini") Then
                runningLocal = False
            Else
                'try checking for local INI file to indicate not running from stick
                key = "RunAsInstalled" & Chr(0)
                section = "Program" & Chr(0)
                'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                Path = GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini" & Chr(0)
                got = modIniFile.GetString(section, key, "0" & Chr(0), Path)
                If CBool(got) Then
                    'ini file indicates we should run as installed version
                    runningLocal = False
                Else
                    'running from a memory stick or other non-installed location
                    runningLocal = True
                End If
            End If
            If runningLocal Then
                'running from a memory stick or other non-installed location
                settingsPath = GetAppPath() & "\Settings"
                nonRoamingSettingsPath = settingsPath
                'need to create
                If Not System.IO.Directory.Exists(settingsPath) Then Call System.IO.Directory.CreateDirectory(settingsPath)
            Else
                'run as installed version
                settingsPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
                '3.12.2 Can't access Local Appdata - it's per-machine.
                nonRoamingSettingsPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            End If
            settingsPath = settingsPath & "\" & companyName
            If Not System.IO.Directory.Exists(settingsPath) Then Call System.IO.Directory.CreateDirectory(settingsPath)
            nonRoamingSettingsPath = nonRoamingSettingsPath & "\" & companyName
            If Not System.IO.Directory.Exists(nonRoamingSettingsPath) Then Call System.IO.Directory.CreateDirectory(nonRoamingSettingsPath)
            commonSettingsPath = settingsPath & "\Common"
            If Not System.IO.Directory.Exists(commonSettingsPath) Then Call System.IO.Directory.CreateDirectory(commonSettingsPath)
            settingsPath = settingsPath & "\" & productName
            If Not System.IO.Directory.Exists(settingsPath) Then Call System.IO.Directory.CreateDirectory(settingsPath)
            nonRoamingSettingsPath = nonRoamingSettingsPath & "\" & productName
            If Not System.IO.Directory.Exists(nonRoamingSettingsPath) Then Call System.IO.Directory.CreateDirectory(nonRoamingSettingsPath)
            settingsPath = settingsPath & "\" & version
            If Not System.IO.Directory.Exists(settingsPath) Then Call System.IO.Directory.CreateDirectory(settingsPath)
            nonRoamingSettingsPath = nonRoamingSettingsPath & "\" & version
            If Not System.IO.Directory.Exists(nonRoamingSettingsPath) Then Call System.IO.Directory.CreateDirectory(nonRoamingSettingsPath)
        Catch
        End Try
    End Sub

    Public Function GetAppPath() As String
        On Error Resume Next
        'work out some paths for use
        GetAppPath = My.Application.Info.DirectoryPath
        If Right(GetAppPath, 1) = "\" Then
            GetAppPath = Left(GetAppPath, Len(GetAppPath) - 1)
        End If
    End Function

    Public Function GetSettingIni(ByRef appName As String, ByRef section As String, ByRef key As String, ByRef default_Renamed As String) As String
        Try
            Dim value As String
            Dim nullTerminatedDefault As String
            Dim Path As String

            nullTerminatedDefault = default_Renamed & Chr(0)
            section = section & Chr(0)
            key = key & Chr(0)
            value = Space(256) & Chr(0)
            Path = settingsPath & "\" & appName & ".ini" & Chr(0)
            GetSettingIni = modIniFile.GetString(section, key, nullTerminatedDefault, Path)
        Catch
            Return ""
        End Try
    End Function

    Public Sub SaveSettingIni(ByRef appName As String, ByRef section As String, ByRef key As String, ByRef value As String)
        Try
            Dim Path As String

            section = section & Chr(0)
            key = key & Chr(0)
            value = value & Chr(0)
            Path = settingsPath & "\" & appName & ".ini"

            Call modIniFile.WriteString(section, key, value, Path)
        Catch
        End Try
    End Sub

    Public Function ReadAppEXEIni(ByRef section As String, ByRef key As String, ByRef default_Renamed As String) As String
        Try
            Dim iniFile As String
            Dim got As String

            'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            iniFile = GetAppPath() & "\" & My.Application.Info.AssemblyName & ".ini" & Chr(0)
            section = section & Chr(0)
            key = key & Chr(0)
            default_Renamed = default_Renamed & Chr(0)
            got = Space(255) & Chr(0)

            Return modIniFile.GetString(section, key, default_Renamed, iniFile)
        Catch
            Return default_Renamed
        End Try
    End Function
End Module