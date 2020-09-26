Option Explicit On
Module modCharacterSupport
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
	
	'This module allows WebbIE to cope with different character sets.  IE will pass
	' through a string identifyer for the document loaded.  This has to be matched
	' to a Windows font and charset that supports the display of the characters
	' in the web page.  The matching of IE ID with charset and font is done by
	' the array of languageCharsetMapping types called languageCharsetMappings
	
	
    Public Sub InitCharsetMappings()
        'TODO Charsets!
        '		'called when the program starts up: loads all the mappings for use
        '		'
        '		On Error GoTo errorHandler
        '		'setup all the entries for languageCharsetMappings
        '		'load this from a configuration file
        '		Dim fo As New Scripting.FileSystemObject ' access to the file system
        '		Dim ts As Scripting.TextStream 'access to the contents of a text file
        '		Dim i As Short ' counter
        '		Dim b() As Byte
        '		Dim got As String
        '		Dim lines() As String

        '		'UPGRADE_ISSUE: Global method LoadResData was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '		b = VB6.LoadResData("COMMON", "CHARSETS")
        '        got = System.Text.UnicodeEncoding.Unicode.GetString(b)
        '		lines = Split(got, vbNewLine)
        '		'get the number of mappings
        '		numberLanguageCharsetMappings = CShort(Trim(lines(LBound(lines))))
        '		'define the array
        '		'UPGRADE_WARNING: Lower bound of array languageCharsetMappings was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
        '		ReDim languageCharsetMappings(numberLanguageCharsetMappings)
        '		'Debug.Print "Got " & numberLanguageCharsetMappings
        '		For i = 1 To numberLanguageCharsetMappings
        '			''ts.ReadLine  'skip a comment line
        '			languageCharsetMappings(i).windowsCharset = CInt(lines(i * 3 - 2))
        '			'Debug.Print "LCM:" & languageCharsetMappings(i).windowsCharset
        '			languageCharsetMappings(i).localeID = CInt(lines(i * 3 - 1))
        '			languageCharsetMappings(i).ieDescription = CStr(Trim(lines(i * 3)))
        '			'Debug.Print "LCM.ieDesc:" & languageCharsetMappings(i).ieDescription
        '		Next i
        '		'Debug.Print languageCharsetMappings(numberLanguageCharsetMappings).ieDescription

        '		Exit Sub
        'errorHandler: 
        '		'problem reading language information: set the character set mappings to
        '		'void, and rely on ParseHTML to use the default set for every page.
        '		numberLanguageCharsetMappings = 0
        '		Exit Sub
    End Sub

    Public Sub InitSystemLocale()
        Try
            'sets up the lngSystemLocale variable, by getting the current locale and then
            'converting it into the root locale - so British English gets converted to US English,
            'for example.  This will then be used to determine whether a character set will be
            'displayed properly, which is the kind of coarse-grain determintation that doesn't
            'need the exact locale.

            lngSystemLocale = NativeMethods.GetUserDefaultLCID 'okay, now convert this into a compatible locale
            Select Case lngSystemLocale
                Case 1025 ' Arabic - Saudi Arabia
                    lngSystemLocale = LOCALE_ARABIC
                Case 1026 ' Bulgarian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1027 ' Catalan
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1028 ' Chinese - Taiwan
                    lngSystemLocale = LOCALE_CHINESE_SIMPLIFIED
                Case 1029 ' Czech
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1030 ' Danish
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1031 ' German - Germany
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1032 ' Greek
                    lngSystemLocale = LOCALE_GREEK
                Case 1033 ' English - United States
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1034 ' Spanish - Spain
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1035 ' Finnish
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1036 ' French - France
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1037 ' Hebrew
                    lngSystemLocale = LOCALE_HEBREW
                Case 1038 ' Hungarian
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1039 ' Icelandic
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1040 ' Italian - Italy
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1041 ' Japanese
                    lngSystemLocale = LOCALE_JAPANESE
                Case 1042 ' Korean
                    lngSystemLocale = LOCALE_KOREAN
                Case 1043 ' Dutch – The Netherlands
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1044 ' Norwegian - Bokmål
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1045 ' Polish
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1046 ' Portuguese - Brazil
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1047 ' Raeto-Romance
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1048 ' Romanian - Romania
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1049 ' Russian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1050 ' Croatian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1051 ' Slovak
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1052 ' Albanian
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1053 ' Swedish - Sweden
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1054 ' Thai
                    lngSystemLocale = LOCALE_THAI
                Case 1055 ' Turkish
                    lngSystemLocale = LOCALE_TURKISH
                Case 1056 ' Urdu
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1057 ' Indonesian
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1058 ' Ukrainian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1059 ' Belarusian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1060 ' Slovenian
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1061 ' Estonian
                    lngSystemLocale = LOCALE_BALTIC
                Case 1062 ' Latvian
                    lngSystemLocale = LOCALE_BALTIC
                Case 1063 ' Lithuanian
                    lngSystemLocale = LOCALE_BALTIC
                Case 1065 ' Farsi
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1066 ' Vietnamese
                    lngSystemLocale = LOCALE_VIETNAMESE
                Case 1067 ' Armenian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1068 ' Azeri – Latin
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1069 ' Basque
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1070 ' Sorbian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1071 ' FYRO Macedonian
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 1072 ' Sutu
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1073 ' Tsonga
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1074 ' Setsuana
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1076 ' Xhosa
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1077 ' Zulu
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1078 ' Afrikaans
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1080 ' Faroese
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1081 ' Hindi
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1082 ' Maltese
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1084 ' Gaelic - Scotland
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1085 ' Yiddish
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1086 ' Malay - Malaysia
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1089 ' Swahili
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1091 ' Uzbek – Latin
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1092 ' Tatar
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 1097 ' Tamil
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1102 ' Marathi
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 1103 ' Sanskrit
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2049 ' Arabic - Iraq
                    lngSystemLocale = LOCALE_ARABIC
                Case 2052 ' Chinese - China
                    lngSystemLocale = LOCALE_CHINESE_SIMPLIFIED
                Case 2055 ' German - Switzerland
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2057 ' English - United Kingdom
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2058 ' Spanish - Mexico
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2060 ' French - Belgium
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2064 ' Italian - Switzerland
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2067 ' Dutch - Belgium
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2068 ' Norwegian – Nynorsk
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2070 ' Portuguese - Portugal
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2072 ' Romanian - Moldova
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 2073 ' Russian - Moldova
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 2074 ' Serbian – Latin
                    lngSystemLocale = LOCALE_EASTERN_EUROPE
                Case 2077 ' Swedish - Finland
                    lngSystemLocale = LOCALE_BALTIC
                Case 2092 ' Azeri – Cyrillic
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 2108 ' Gaelic – Ireland
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2110 ' Malay – Brunei
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 2115 ' Uzbek – Cyrillic
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 3073 ' Arabic - Egypt
                    lngSystemLocale = LOCALE_ARABIC
                Case 3076 ' Chinese - Hong Kong S.A.R.
                    lngSystemLocale = LOCALE_CHINESE_SIMPLIFIED
                Case 3079 ' German - Austria
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 3081 ' English - Australia
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 3084 ' French - Canada
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 3098 ' Serbian - Cyrillic
                    lngSystemLocale = LOCALE_RUSSIAN
                Case 4097 ' Arabic - Libya
                    lngSystemLocale = LOCALE_ARABIC
                Case 4100 ' Chinese - Singapore
                    lngSystemLocale = LOCALE_CHINESE_SIMPLIFIED
                Case 4103 ' German - Luxembourg
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 4105 ' English - Canada
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 4106 ' Spanish - Guatemala
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 4108 ' French - Switzerland
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 5121 ' Arabic - Algeria
                    lngSystemLocale = LOCALE_ARABIC
                Case 5124 ' Chinese – Macau S.A.R
                    lngSystemLocale = LOCALE_CHINESE_SIMPLIFIED
                Case 5127 ' German - Liechtenstein
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 5129 ' English - New Zealand
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 5130 ' Spanish - Costa Rica
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 5132 ' French - Luxembourg
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 6145 ' Arabic - Morocco
                    lngSystemLocale = LOCALE_ARABIC
                Case 6153 ' English - Ireland
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 6154 ' Spanish - Panama
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 7169 ' Arabic - Tunisia
                    lngSystemLocale = LOCALE_ARABIC
                Case 7177 ' English - South Africa
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 7178 ' Spanish - Dominican Republic
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 8193 ' Arabic - Oman
                    lngSystemLocale = LOCALE_ARABIC
                Case 8201 ' English - Jamaica
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 8202 ' Spanish - Venezuela
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 9217 ' Arabic - Yemen
                    lngSystemLocale = LOCALE_ARABIC
                Case 9225 ' English – Carribbean
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 9226 ' Spanish - Colombia
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 10241 ' Arabic - Syria
                    lngSystemLocale = LOCALE_ARABIC
                Case 10249 ' English - Belize
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 10250 ' Spanish - Peru
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 11265 ' Arabic - Jordan
                    lngSystemLocale = LOCALE_ARABIC
                Case 11273 ' English - Trinidad
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 11274 ' Spanish - Argentina
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 12289 ' Arabic - Lebanon
                    lngSystemLocale = LOCALE_ARABIC
                Case 12298 ' Spanish - Ecuador
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 13313 ' Arabic - Kuwait
                    lngSystemLocale = LOCALE_ARABIC
                Case 13321 ' English – Phillippines
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 13322 ' Spanish - Chile
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 14337 ' Arabic – United Arab Emirates
                    lngSystemLocale = LOCALE_ARABIC
                Case 14346 ' Spanish - Uruguay
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 15361 ' Arabic - Bahrain
                    lngSystemLocale = LOCALE_ARABIC
                Case 15370 ' Spanish - Paraguay
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 16385 ' Arabic - Qatar
                    lngSystemLocale = LOCALE_ARABIC
                Case 16394 ' Spanish - Bolivia
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 17418 ' Spanish - El Salvador
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 18442 ' Spanish - Honduras
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 19466 ' Spanish - Nicaragua
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN
                Case 20490 ' Spanish - Puerto Rico
                    lngSystemLocale = LOCALE_WESTERN_EUROPEAN


            End Select
        Catch

        End Try
    End Sub
End Module