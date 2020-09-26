Option Strict Off
Option Explicit On
Module modDialupManager
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
	
	
	'This module manages whether the computer is dialled-up to a modem connection
	'and hangs up if necessary. It supercedes EndCall.bas.
	
	Private Const RAS_MaxDeviceType As Short = 16
	Private Const RAS95_MaxDeviceName As Short = 128
	Private Const RAS95_MaxEntryName As Short = 256
	Private Structure RASCONN95
		'set dwsize to 412
		Dim dwSize As Integer
		Dim hRasConn As Integer
		<VBFixedArray(RAS95_MaxEntryName)> Dim szEntryName() As Byte
		<VBFixedArray(RAS_MaxDeviceType)> Dim szDeviceType() As Byte
		<VBFixedArray(RAS95_MaxDeviceName)> Dim szDeviceName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szEntryName(RAS95_MaxEntryName)
			ReDim szDeviceType(RAS_MaxDeviceType)
			ReDim szDeviceName(RAS95_MaxDeviceName)
		End Sub
	End Structure
	Private Structure RASENTRYNAME95
		'set dwsize to 264
		Dim dwSize As Integer
		<VBFixedArray(RAS95_MaxEntryName)> Dim szEntryName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szEntryName(RAS95_MaxEntryName)
		End Sub
	End Structure
	Private Structure RASDEVINFO
		Dim dwSize As Integer
		<VBFixedArray(RAS_MaxDeviceType)> Dim szDeviceType() As Byte
		<VBFixedArray(RAS95_MaxDeviceName)> Dim szDeviceName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szDeviceType(RAS_MaxDeviceType)
			ReDim szDeviceName(RAS95_MaxDeviceName)
		End Sub
	End Structure
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RasEnumDevices Lib "RasApi32.DLL"  Alias "RasEnumDevicesA"(ByRef lprasdevinfo As Any, ByRef lpcb As Integer, ByRef lpcConnections As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RasEnumConnections Lib "RasApi32.DLL"  Alias "RasEnumConnectionsA"(ByRef lprasconn As Any, ByRef lpcb As Integer, ByRef lpcConnections As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RasEnumEntries Lib "RasApi32.DLL"  Alias "RasEnumEntriesA"(ByVal reserved As String, ByVal lpszPhonebook As String, ByRef lprasentryname As Any, ByRef lpcb As Integer, ByRef lpcEntries As Integer) As Integer
	
	Public Function DialledUp() As String
		'returns whether the computer is dialled up to a modem connection or not.
		'if so, returns true.
		On Error Resume Next
		Dim connectionSize As Integer
		Dim lengthEnumeration As Integer
		Dim connectionName As String
		'Dim connections(255) As RASENTRYNAME95
		'UPGRADE_WARNING: Array activeConnections may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
		Dim activeConnections(255) As RASCONN95
		'Dim connectionNames() As String
		'Dim activeConnectionName As String
		
		'Dim b As String
		'Dim i As Long
		'Dim allConnections As String
		
		'    connections(0).dwSize = 264
		'    connectionSize = 256 * connections(0).dwSize
		'    Call RasEnumEntries(vbNullString, vbNullString, connections(0), connectionSize, lengthEnumeration)
		'    ReDim connectionNames(0 To lengthEnumeration)
		'    If lengthEnumeration = 0 Then
		'        'no dial-up connections, so can't be dialled up...
		'        DialledUp = False
		'    Else
		'        'at least one, let's check them
		'        For i = 0 To lengthEnumeration - 1
		'            connectionName = StrConv(connections(i).szEntryName(), vbUnicode)
		'            'a = connections(i).szEntryName
		'            connectionNames(i) = Left$(connectionName, InStr(connectionName, Chr$(0)) - 1)
		'            'allConnections = allConnections & Left$(connectionName, InStr(connectionName, Chr$(0)) - 1) & vbNewLine
		'            'List1.AddItem Left$(a$, InStr(a$, Chr$(0)) - 1)
		'        Next i
		'    End If
		
		activeConnections(0).dwSize = 412
		connectionSize = 256 * activeConnections(0).dwSize
		'UPGRADE_WARNING: Couldn't resolve default property of object activeConnections(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call RasEnumConnections(activeConnections(0), connectionSize, lengthEnumeration)
		If lengthEnumeration > 0 Then
			'there are some active connections!
			DialledUp = CStr(True)
		Else
			DialledUp = CStr(False)
		End If
		'        For i = 0 To lengthEnumeration - 1
		'            activeConnectionName = StrConv(connections(i).szEntryName(), vbUnicode)
		'
		'        Next i
		'    End If
		'    l = RasEnumConnections(r(0), s, ln)
		'    For l = 0 To ln - 1
		'        a$ = StrConv(r(l).szEntryName(), vbUnicode)
		'        a$ = Left$(a$, InStr(a$, Chr$(0)) - 1)
		'        If a$ = b$ Then MsgBox "Connected (or connecting)!": Exit Function
		'    Next
		'    MsgBox "Not Connected!"
	End Function
End Module