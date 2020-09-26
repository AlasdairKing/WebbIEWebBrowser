Option Strict Off
Option Explicit On
Module modEndCall
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
	
	'Terminates all active Modem/Terminal adapter connections
	'To call use HangUp()
	
	
	Private Const RAS_MAXENTRYNAME As Short = 256
	Private Const RAS_MaxDeviceType As Short = 16
	Private Const RAS_MAXDEVICENAME As Short = 128
	Private Const RAS_RASCONNSIZE As Short = 412
	
	
	Private Structure RasEntryName
		Dim dwSize As Integer
		<VBFixedArray(RAS_MAXENTRYNAME)> Dim szEntryName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szEntryName(RAS_MAXENTRYNAME)
		End Sub
	End Structure
	
	Private Structure RasConn
		Dim dwSize As Integer
		Dim hRasConn As Integer
		<VBFixedArray(RAS_MAXENTRYNAME)> Dim szEntryName() As Byte
		<VBFixedArray(RAS_MaxDeviceType)> Dim szDeviceType() As Byte
		<VBFixedArray(RAS_MAXDEVICENAME)> Dim szDeviceName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim szEntryName(RAS_MAXENTRYNAME)
			ReDim szDeviceType(RAS_MaxDeviceType)
			ReDim szDeviceName(RAS_MAXDEVICENAME)
		End Sub
	End Structure
	
	Private gstrISPName As String
	Private ReturnCode As Integer
	
	
	Public Sub HangUp()
		On Error Resume Next
		Dim i As Integer
		'UPGRADE_WARNING: Array lprasconn may need to have individual elements initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B97B714D-9338-48AC-B03F-345B617E2B02"'
		Dim lprasconn(255) As RasConn
		Dim lpcb As Integer
		Dim lpcConnections As Integer
		Dim hRasConn As Integer
		lprasconn(0).dwSize = RAS_RASCONNSIZE
		lpcb = RAS_MAXENTRYNAME * lprasconn(0).dwSize
		lpcConnections = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object lprasconn(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ReturnCode = RasEnumConnections(lprasconn(0), lpcb, lpcConnections)
		
		If ReturnCode = ERROR_SUCCESS Then 'if query was successful
			For i = 0 To lpcConnections - 1 'look at all connections
				'match the connections with the dialled one
				If Trim(ByteToString(lprasconn(i).szEntryName)) = Trim(gstrISPName) Then
					hRasConn = lprasconn(i).hRasConn
					ReturnCode = RasHangUp(hRasConn) 'hang up
				End If
			Next i
		End If
		
	End Sub
	
	Private Function ByteToString(ByRef bytString() As Byte) As String
		On Error Resume Next
		Dim i As Short
		ByteToString = "" 'if return is empty
		i = 0
		While bytString(i) = 0 'Error_success
			ByteToString = ByteToString & Chr(bytString(i))
			i = i + 1
		End While
	End Function
End Module