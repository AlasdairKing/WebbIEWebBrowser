Option Strict On
Option Explicit On
Module modChangehomepage
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
	
	'This will change the registry values for internet explorer
	'Allows you to change the homepage for IE
	'Use by typing: SetStartPage ("http://www.myaddress.com")
	
    ''' <summary>
    ''' Calls SaveString with 4 parameters.
    ''' </summary>
    ''' <param name="url"></param>
    ''' <remarks></remarks>
    Public Sub SetStartPage(ByRef url As String)
        Try
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Explorer\Main", True)
            Call regKey.SetValue("Start Page", url)
            Call regKey.Close()
        Catch
        End Try
    End Sub

    Public Function RetrieveStartPage() As String
        Try
            Dim homepage As String
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Explorer\Main", False)
            homepage = regKey.GetValue("Start Page").ToString
            Call regKey.Close()
            'return this URL
            Return homepage
        Catch
            Return ""
        End Try
    End Function
	
End Module