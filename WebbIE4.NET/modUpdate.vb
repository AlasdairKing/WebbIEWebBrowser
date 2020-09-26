Option Strict On
Option Explicit On
Module modUpdate
	
	'Do an update check for new versions

	Public Sub CheckForUpdates()
		Dim checkPath As String
		
		checkPath = GetAppPath
		checkPath = checkPath & "\CheckForWebbIEUpdates.exe"
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Len(Dir(checkPath)) > 0 Then
			Call Shell(checkPath, AppWinStyle.MinimizedNoFocus)
		End If
		Exit Sub
	End Sub
	
	'Private Function UpdateIntervalElapsed(days As Long, Path As String) As Boolean
	'    'Returns whether the number of days indicated by days has elapsed since the last update
	'    On Error GoTo Failed:
	'    Dim lastTime As String
	'    Dim lastDate As Date
	'    Dim dateNow As Date
	'
	'    lastTime = GetSetting(Path, "Update", "LastCheck")
	'    If lastTime = "" Then
	'        UpdateIntervalElapsed = True
	'    Else
	'        lastDate = CDate(lastTime)
	'        dateNow = Now
	'        If (lastDate < dateNow - days) Then
	'            UpdateIntervalElapsed = True
	'        End If
	'    End If
	'    Exit Function
	'Failed:
	'        UpdateIntervalElapsed = False
	'End Function
	'
	'Private Function GetSetting(iniFilePath As String, section As String, key As String) As String
    '    
	'    Dim value As String
	'    Dim size As Long
	'    Dim nullTerminatedDefault As String
	'    Dim Path As String
	'
	'    nullTerminatedDefault = "" & Chr(0)
	'    section = section & Chr(0)
	'    key = key & Chr(0)
	'    value = Space(256) & Chr(0)
	'    size = GetPrivateProfileString(section, key, nullTerminatedDefault, value, Len(value), iniFilePath)
	'    If size = 0 Then
	'        'nothing read
	'        GetSetting = ""
	'    Else
	'        'read okay, cut off unread section (including 0)
	'        GetSetting = left(value, size - 1)
	'    End If
	'End Function
End Module