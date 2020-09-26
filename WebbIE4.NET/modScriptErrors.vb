Option Strict On
Option Explicit On
Module modScriptErrors
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
	
	'modScriptErrors
	'This module turns off IE script errors in the registry, and restores their state when the program
	'closes. Prevents all those annoying script errors popping up in WebbIE.
	
	
	Private mDisableScriptDebugger As String
	Private mDisableScriptDebuggerIE As String
	Private mErrorDlgDisplayedOnEveryError As String



    Public Sub DisableScriptErrors()
        Try
            'Turn off all script errors after recording their values.
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Explorer\Main", True)
            If regKey.GetValue("Disable Script Debugger") Is Nothing Then
                mDisableScriptDebugger = "yes" ' Disable debuggers by default.
            Else
                mDisableScriptDebugger = regKey.GetValue("Disable Script Debugger").ToString
            End If
            Call regKey.SetValue("Disable Script Debugger", "yes")
            If regKey.GetValue("DisableScriptDebuggerIE") Is Nothing Then
                mDisableScriptDebuggerIE = "yes" ' Disable debuggers by default.
            Else
                mDisableScriptDebuggerIE = regKey.GetValue("DisableScriptDebuggerIE").ToString
            End If
            Call regKey.SetValue("DisableScriptDebuggerIE", "yes")
            If regKey.GetValue("Error Dlg Displayed On Every Error") Is Nothing Then
                mErrorDlgDisplayedOnEveryError = "yes" ' Disable debuggers by default.
            Else
                mErrorDlgDisplayedOnEveryError = regKey.GetValue("Error Dlg Displayed On Every Error").ToString
            End If
            Call regKey.SetValue("Error Dlg Displayed On Every Error", "no")
            Call regKey.Close()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Restore script error settings to what they were when DisableScriptErrors was called.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RestoreScriptErrors()
        Try
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Explorer\Main", True)
            If Len(mDisableScriptDebugger) > 0 Then Call regKey.SetValue("Disable Script Debugger", mDisableScriptDebugger)
            If Len(mDisableScriptDebuggerIE) > 0 Then Call regKey.SetValue("DisableScriptDebuggerIE", mDisableScriptDebuggerIE)
            If Len(mErrorDlgDisplayedOnEveryError) > 0 Then Call regKey.SetValue("Error Dlg Displayed On Every Error", mErrorDlgDisplayedOnEveryError)
            regKey.Close()
        Catch
        End Try
    End Sub
	
End Module