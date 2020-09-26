Option Explicit On
Module modExplorerVersion
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
	
	'Displays the version of IE used
	'IEVersion returns an integer number e.g 5 for IE 5
	'IEVersionString returns a string plus a build number e.g. Internet Explorer 5.0.1.0
	
	
	
	Public Function IEVersion() As Integer
        Try

            Dim VersionInfo As DllVersionInfo

            VersionInfo.cbSize = Len(VersionInfo)

            Call DllGetVersion(VersionInfo)

            IEVersion = VersionInfo.dwMajorVersion
        Catch
            Return 7 ' 7 is what WebBrowser behaves like unless I do something different.
        End Try
    End Function

    Public Function IEVersionString() As Object

        Dim VersionInfo As DllVersionInfo
        VersionInfo.cbSize = Len(VersionInfo)

        Call DllGetVersion(VersionInfo)

        'UPGRADE_WARNING: Couldn't resolve default property of object IEVersionString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        IEVersionString = VersionInfo.dwMajorVersion & "." & VersionInfo.dwMinorVersion & "." & VersionInfo.dwBuildNumber

    End Function
End Module