Option Strict Off
Option Explicit On
Module modRememberPosition
	'Calendar
	'Copyright Alasdair King, 2010, http://www.alasdairking.me.uk
	'Released under the GNU Public Licence, Version 3.
	
	'Code to ensure your form is correctly positioned according to where you previously
	'put it.
	
	'23 Mar 2009
	'   Updated to prevent settings being saved when form exits minimized.
	
	
	Public Sub LoadPosition(ByRef f As System.Windows.Forms.Form)
		'Load position of f from ini file. Call in Form_Load.
		On Error Resume Next
		
		'Start maximised if that's how we left it.
		'Hey, no, don't make Maximised the default!
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		If CShort(modPath.GetSettingIni(My.Application.Info.AssemblyName, f.Name & "Position", "Maximised", CStr(vbNormal))) = System.Windows.Forms.FormWindowState.Maximized Then f.WindowState = System.Windows.Forms.FormWindowState.Maximized
		
		'Don't remember anything else. This is because users can accidentally make the
		'form too small, or off the screen, and screenreaders stop working.
		'    If f.WindowState = vbNormal Then
		'        f.left = modPath.GetSettingIni(App.EXEName, f.name & "Position", "Left", f.left)
		'        If f.left < 0 Then f.left = 0
		'        f.Top = modPath.GetSettingIni(App.EXEName, f.name & "Position", "Top", f.Top)
		'        If f.Top < 0 Then f.Top = 0
		'        f.width = modPath.GetSettingIni(App.EXEName, f.name & "Position", "Width", f.width)
		'        f.height = modPath.GetSettingIni(App.EXEName, f.name & "Position", "Height", f.height)
		'    End If
	End Sub
	
	Public Sub SavePosition(ByRef f As System.Windows.Forms.Form)
		'Save position of f to ini file. Call in Form_Unload.
		On Error Resume Next
		If f.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			'Don't save settings, we're just closed from start bar.
		Else
			'        Call modPath.SaveSettingIni(App.EXEName, f.name & "Position", "Left", f.left)
			'        Call modPath.SaveSettingIni(App.EXEName, f.name & "Position", "Top", f.Top)
			'        Call modPath.SaveSettingIni(App.EXEName, f.name & "Position", "Width", f.width)
			'        Call modPath.SaveSettingIni(App.EXEName, f.name & "Position", "Height", f.height)
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			Call modPath.SaveSettingIni(My.Application.Info.AssemblyName, f.Name & "Position", "Maximised", CStr(f.WindowState))
		End If
	End Sub
End Module