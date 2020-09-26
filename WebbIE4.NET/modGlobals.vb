Option Explicit On
Module modGlobals
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

    ''' <summary>
    ''' indicates that only form contents should be parsed and shown.
    ''' </summary>
    Public gShowFormsOnly As Boolean

    ''' <summary>
    ''' 3.7.4. Set this to true when a user action has occurred (e.g. a
    '''   form submit) that MUST result in a page navigation even though the resultant href indicates that
    '''   this is a intra-page (#) link. Try submitting a comment to MetaFilter: the target is .....php/#comment
    '''   which looks like an internal link.
    ''' </summary>
    Public gForceNavigation As Boolean

    ''' <summary>
    ''' How we try to work out if the page has loaded.
    ''' </summary>
    Public Enum PageLoadStrategies
        ''' <summary>
        ''' Try to use the JavaScript onload event to detect page loading.
        ''' </summary>
        UseJavaScriptOnLoad
        ''' <summary>
        ''' Use DocumentComplete event to detect page loading.
        ''' </summary>
        UseDocumentCompleteOnly
    End Enum
    ''' <summary>
    ''' If True, use JavaScript to work out when a page has loaded. If False, use DocumentComplete.
    ''' This is because I think the former results in security problems and the latter in pages
    ''' with rubbish on them. 
    ''' </summary>
    Public gUseJavascriptForPageLoad As PageLoadStrategies = PageLoadStrategies.UseJavaScriptOnLoad

    ''' <summary>
    ''' If set, then
    ''' the user has done something that should result in a
    ''' change to the page - like clicking a Submit button.
    ''' Sometimes, however, this doesn't actually trigger a
    ''' page navigation. If this is the case then you have to
    ''' do a refresh instead.
    ''' </summary>
    Public gNoPageActionSoRefresh As Boolean

    ''' <summary>
    ''' Set when we are closing.
    ''' </summary>
    Public gClosing As Boolean = False

    ''' <summary>
    ''' The main Windows Form hosting the WebBrowser that does all the work.
    ''' </summary>
    Public gWebHost As frmWeb

    ''' <summary>
    ''' The only time that frmWeb is allowed to get focus is when we are trying
    ''' to get a Youtube video to start by clicking on it.
    ''' </summary>
    ''' <remarks></remarks>
    Public gDisplayingYoutube As Boolean

    ''' <summary>
    ''' 'the url of the current web page's RSS feed, or "" if there is none.
    ''' </summary>
    Public gRSSFeedURL As String

    ''' <summary>
    ''' indicates we're quitting the application and event-based stuff we might have
    '''   started should terminate. Primarily the WinInet control. TODO What about gClosing?
    ''' </summary>
    Public gExiting As Boolean

    ''' <summary>
    ''' The number of characters allowed between links if they are to stay on the same line.
    ''' </summary>
    Public Const NUMBER_CHARS_ALLOWED_BETWEEN_LINKS As Integer = 5
    Public Const MAX_NUMBER_LINKS_SUPPORTED As Integer = 2000
	Public Const MAX_NUMBER_TABLES_SUPPORTED As Integer = 100
	Public Const MAX_NUMBER_TEXT_AREA_INPUTS_SUPPORTED As Integer = 100
	Public Const MAX_NUMBER_SELECTS_SUPPORTED As Integer = 100
    Public Const MAX_NUMBER_TEXT_INPUTS_SUPPORTED As Integer = 100
    Public Const MAX_NUMBER_EMAIL_INPUTS_SUPPORTED As Integer = 100
    Public Const MAX_NUMBER_RANGE_INPUTS_SUPPORTED As Integer = 10
    Public Const MAX_NUMBER_SEARCH_INPUTS_SUPPORTED As Integer = 10
	Public Const MAX_NUMBER_PASSWORD_INPUTS_SUPPORTED As Integer = 100
	Public Const MAX_NUMBER_SUBMITS_SUPPORTED As Integer = 100
	Public Const MAX_NUMBER_FILE_INPUTS_SUPPORTED As Integer = 10
	Public Const MAX_NUMBER_BUTTON_INPUTS_SUPPORTED As Integer = 100
    Public Const MAX_NUMBER_RESET_INPUTS_SUPPORTED As Integer = 10
    Public Const MAX_NUMBER_VIDEOS_SUPPORTED As Integer = 10
    Public Const MAX_NUMBER_AUDIOS_SUPPORTED As Integer = 10

	'Mime types
	Public Const MIMETYPE_TEXT As String = "Text Document"
	Public Const MIMETYPE_HTML As String = "HTML Document"
    Public Const MIMETYPE_ACROBAT As String = "Adobe Acrobat Document" ' This will only be true if
    '   Adobe Reader is installed: TODO make this more robust.

    Public Const SKIP_UP As Integer = -1
    Public Const SKIP_DOWN As Integer = 1
	
	Public Const LINK_SECTION As String = "WEBBIE_LINK_SECTION" ' indicates link section
	Public Const CONTENT_SECTION As String = "WEBBIE_CONTENT_SECTION" ' indicates content section
	Public SECTION_MARKER_H1 As String ' what you tell the user
	Public SECTION_MARKER_H2 As String
	Public SECTION_MARKER_H3 As String
	Public SECTION_MARKER_H4 As String
	Public SECTION_MARKER_H5 As String
	Public SECTION_MARKER_H6 As String
	Public SECTION_MARKER_COMMON As String
	
	Public Const TARGET_MARKER As String = "WEBBIE_TARGET_MARKER8766754765876657654"
	Public Const FLASH_INDICATOR As String = "download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version"
	Public Const AJAX_TARGET_MARKER As String = "AJAX_TARGET_MARKERioiouu097897y897687"

    ''' <summary>
    ''' Indicates whether the form with the WebBrowser control should be visible (requested by the user)
    ''' </summary>
    ''' <remarks></remarks>
    Public gShowingWebpage As Boolean = False

    ''' <summary>
    ''' Indicates that the focus should be put back on frmMain.
    ''' </summary>
    ''' <remarks></remarks>
    Public gFrmWebHasFocusAndItShouldNot As Boolean = False

    Structure linkStruct
        ''' <summary>
        ''' link description
        ''' </summary>
        Dim description As String
        ''' <summary>
        ''' link address e.g. www.orange.net
        ''' </summary>
        Dim address As String
        ''' <summary>
        ''' target of link
        ''' </summary>
        Dim target As String
        ''' <summary>
        ''' index of frame in which link resides
        ''' </summary>
        Dim frameIndex As Integer
        ''' <summary>
        ''' the link element
        ''' </summary>
        Dim element As mshtml.IHTMLElement
    End Structure

    Structure targetStruct
        ''' <summary>
        ''' the text following/contained in the target anchor
        ''' </summary>
        Dim idText As String
        ''' <summary>
        ''' the name Of the anchor
        ''' </summary>
        Dim name As String
    End Structure

    Structure selectStruct
        ''' <summary>
        ''' Name of menu
        ''' </summary>
        Dim name As String
        ''' <summary>
        ''' the options in the select/combo
        ''' </summary>
        <VBFixedArray(300)> Dim options() As String
        ''' <summary>
        ''' the number of the option which is selected
        ''' </summary>
        Dim selected As Integer
        ''' <summary>
        ''' number of options in menu
        ''' </summary>
        Dim size As Integer
        ''' <summary>
        ''' the node in the DOM
        ''' </summary>
        Dim node As mshtml.IHTMLSelectElement

        Public Sub Initialize()
            ReDim options(300)
        End Sub
    End Structure

    ''' <summary>
    ''' A frame
    ''' </summary>
    Structure frameStruct
        Dim name As String
        Dim url As String
        Dim Document As mshtml.IHTMLDocument2
        Dim target As Boolean
    End Structure

    ''' <summary>
    ''' the text searched for in a page by the user
    ''' </summary>
    Public gfindText As String

    'Nb: the indexing number (i) must match the input number
    Public textInput(MAX_NUMBER_TEXT_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement 'text input name & number
    Public numTextInputs As Integer 'number of text boxes on page

    Public emailInput(MAX_NUMBER_EMAIL_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement ' email input name and number
    Public numEmailInputs As Integer ' number of email inputs on page

    Public rangeInput(MAX_NUMBER_RANGE_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement ' range input
    Public numRangeInputs As Integer

    Public searchInput(MAX_NUMBER_SEARCH_INPUTS_SUPPORTED) As mshtml.IHTMLElement ' search textbox input from html5
    Public numSearchInputs As Integer ' number of search boxes on page.

    Public gImageHrefs As Collection ' contains the alt text for each image

    ''' <summary>
    ''' Indicates that the HTML5 "ARTICLE" element has been found on the page.
    ''' </summary>
    ''' <remarks></remarks>
    Public gPageHasAnArticle As Boolean

    ''' <summary>
    ''' Indicates that the HTML5 "MAIN" element has been found on the page.
    ''' </summary>
    ''' <remarks></remarks>
    Public gPageHasMain As Boolean

    Public passwordInput(0 To MAX_NUMBER_PASSWORD_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement 'password input name & number
    Public numPassInputs As Integer 'number of password boxes on page
	
    Public submitInput(0 To MAX_NUMBER_SUBMITS_SUPPORTED - 1) As mshtml.IHTMLElement 'submit input name & number
    Public numSubmitInputs As Integer 'number of submits on this page
	
    Public radioInput(0 To 499) As mshtml.IHTMLElement 'radio input name & number
    Public numRadioInputs As Integer
	
    Public videos(0 To MAX_NUMBER_VIDEOS_SUPPORTED - 1) As mshtml.IHTMLElement
    Public numVideos As Integer

    Public audios(0 To MAX_NUMBER_AUDIOS_SUPPORTED - 1) As mshtml.IHTMLElement
    Public numAudios As Integer

    Public checkboxInput(0 To 499) As mshtml.IHTMLInputElement 'checkbox input name & number
    Public numCheckboxInputs As Integer 'number of checkboxes on page
	
    Public resetInput(0 To MAX_NUMBER_RESET_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement 'reset button name and number
    Public numResetInputs As Integer 'number of resets on page
	
    Public buttonInput(0 To MAX_NUMBER_BUTTON_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement ' input button control
    Public numButtonInputs As Integer ' number of input buttons on page
	
    Public fileInput(0 To MAX_NUMBER_FILE_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement ' file selection control
    Public numFileInputs As Integer ' number of file selection controls on page
	
    Public textAreaInput(0 To MAX_NUMBER_TEXT_AREA_INPUTS_SUPPORTED - 1) As mshtml.IHTMLElement ' text area input
    Public numTextAreaInputs As Integer
	
    Public altText(0 To MAX_NUMBER_LINKS_SUPPORTED - 1) As String 'description for images
    Public numImages As Integer 'holds number of images
	
    Public selects(0 To MAX_NUMBER_SELECTS_SUPPORTED - 1) As selectStruct 'combo boxes
    Public numSelects As Integer
	
    Public gLinks(0 To MAX_NUMBER_LINKS_SUPPORTED - 1) As linkStruct
    Public gSortedLink(0 To MAX_NUMBER_LINKS_SUPPORTED - 1) As linkStruct 'sorted alphabetically
    Public gNumLinks As Integer
	
	Public targets(599) As targetStruct
    Public numTargets As Integer
	
    Public numForms As Integer
	

    Public tables(0 To MAX_NUMBER_TABLES_SUPPORTED) As mshtml.IHTMLElement
    Public numTables As Integer
	
    Public gCurrentHomepage As String 'the current IE homepage
	
	Public blnAutoskiplinks As Boolean ' whether the program automatically skips
	' to the first piece of text post-links

    Public lngPageLocale As Integer ' the locale to use for translating a page into ANSI
    Public lngSystemLocale As Integer ' the locale of the user's system
	Public blnAllowPopups As Boolean ' whether popups are allowed or suppressed.  They are prompted for in any case

	'bookmark stuff
    Public Const BOOKMARK_MENU_INDEX As Integer = 3 ' The location from left to right of the Bookmarks menu item, counting from 0
    Public Const NUMBER_ITEMS_ABOVE_BOOKMARK_LIST As Integer = 3
	Public Const INDEX_AS_POSITION As Boolean = True
    Public Const MAX_NUMBER_BOOKMARKS As Integer = 500 'maximum number of bookmarks supported by the app
	'UPGRADE_WARNING: Lower bound of array userBookmarks was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public userBookmarks(MAX_NUMBER_BOOKMARKS) As String ' the bookmarks and their destinations
    Public numberBookmarks As Integer ' the number of bookmarks
	
	' The registry titles for WebbIE
	Public Const USER_SETTINGS As String = "User Settings"

	'Character set constants: used to define the font character set for the main text
	'output so that is displays non-ansi characters correctly (e.g. Russian)
	Structure languageCharsetMapping
		Dim ieDescription As String
		Dim windowsCharset As Integer
		Dim localeID As Integer
	End Structure
	
    Public gNumberLanguageCharsetMappings As Integer
    Public gLanguageCharsetMappings() As languageCharsetMapping
	
    Public gCharacterMapper() As Integer
	Public Const NUMBER_UNICODE_CHARACTERS_SUPPORTED As Integer = 65536
	' the number of unicode characters supported for translation into ansi
	' (also the maximum array value of characterMappper + 1)
    Public Const NUMBER_CHARSETS_SUPPORTED As Integer = 2
	'the number of charsets that we've written the code information for.
	'UPGRADE_WARNING: Lower bound of array supportedCharsets was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
    Public gSupportedCharsets(NUMBER_CHARSETS_SUPPORTED) As Integer
	'the list of charsets supported, used to pick from an array of them in frmLanguage
	
	'The charsets for choosing Unicode conversions
    Public Const CHARSET_WESTERN_EUROPEAN As Integer = 0
    Public Const CHARSET_JAPANESE As Integer = 128 ' Japanese
    Public Const CHARSET_KOREAN As Integer = 129
    Public Const CHARSET_CHINESE_SIMPLIFIED As Integer = 134 'Simplified Chinese
    Public Const CHARSET_CHINESE_TRADITIONAL As Integer = 136
    Public Const CHARSET_GREEK As Integer = 161
    Public Const CHARSET_TURKISH As Integer = 162
    Public Const CHARSET_VIETNAMESE As Integer = 163
    Public Const CHARSET_HEBREW As Integer = 177
    Public Const CHARSET_ARABIC As Integer = 178
    Public Const CHARSET_BALTIC As Integer = 186
    Public Const CHARSET_RUSSIAN As Integer = 204
    Public Const CHARSET_THAI As Integer = 222
    Public Const CHARSET_EASTERN_EUROPE As Integer = 238
	
	'The locales for choosing Unicode conversions
    Public Const LOCALE_ARABIC As Integer = 1025
    Public Const LOCALE_CHINESE_TRADITIONAL As Integer = 1028
    Public Const LOCALE_EASTERN_EUROPE As Integer = 1029
    Public Const LOCALE_GREEK As Integer = 1032
    Public Const LOCALE_WESTERN_EUROPEAN As Integer = 1033 ' the locale needed to switch to 256-char
    Public Const LOCALE_HEBREW As Integer = 1037
    Public Const LOCALE_JAPANESE As Integer = 1041 ' Japanese
    Public Const LOCALE_KOREAN As Integer = 1042
    Public Const LOCALE_RUSSIAN As Integer = 1049
    Public Const LOCALE_THAI As Integer = 1054
    Public Const LOCALE_TURKISH As Integer = 1055
    Public Const LOCALE_BALTIC As Integer = 1061
    Public Const LOCALE_VIETNAMESE As Integer = 1066
    Public Const LOCALE_CHINESE_SIMPLIFIED As Integer = 2052 'Simplified Chinese
	
	'the W3C DOM node types,
    Public Const ELEMENT_NODE As Integer = 1
    Public Const ATTRIBUTE_NODE As Integer = 2
    Public Const TEXT_NODE As Integer = 3
    Public Const CDATA_SECTION_NODE As Integer = 4
    Public Const ENTITY_REFERENCE_NODE As Integer = 5
    Public Const ENTITY_NODE As Integer = 6
    Public Const PROCESSING_INSTRUCTION_NODE As Integer = 7
    Public Const COMMENT_NODE As Integer = 8
    Public Const DOCUMENT_NODE As Integer = 9
    Public Const DOCUMENT_TYPE_NODE As Integer = 10
    Public Const DOCUMENT_FRAGMENT_NODE As Integer = 11
    Public Const NOTATION_NODE As Integer = 12
	
	'the terms used to describe the items on the page e.g. LINK for an A element
	Public ID_LINK As String
	Public ID_SELECT As String
	Public ID_BUTTON As String
	Public ID_CHECKBOX As String
	Public ID_RADIO As String
    Public ID_TEXTBOX As String
    Public ID_EMAILINPUT As String
    Public ID_RANGEINPUT As String
	Public ID_PASSWORD As String
	Public ID_SUBMIT As String
	Public ID_FILE As String
	Public ID_RESET As String
	Public ID_TEXTAREA As String
	Public ID_TABLE As String
    Public ID_VIDEO As String
    Public ID_AUDIO As String

	Public ID_DISABLED As String
	Public ID_READONLY As String
	Public ID_CHECKED As String
	Public ID_NOTCHECKED As String
	Public ID_SELECTED As String
    Public ID_NOTSELECTED As String
    ''' <summary>
    ''' Set this if you want the webbrowser to navigate somewhere. frmWeb will notice and send it off. 
    ''' </summary>
    Friend gDesiredURL As String

    Public Sub Initialise()
        On Error Resume Next
        'Cute characters, but not good for screenreaders!
        'SECTION_MARKER_H1 = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") & " 1" ' what you tell the user
        'SECTION_MARKER_H2 = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") & " 2" ' what you tell the user
        'SECTION_MARKER_H3 = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") & " 3" ' what you tell the user
        'SECTION_MARKER_H4 = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") & " 4" ' what you tell the user
        'SECTION_MARKER_H5 = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") & " 5" ' what you tell the user
        'SECTION_MARKER_H6 = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") & " 6" ' what you tell the user
        'SECTION_MARKER_COMMON = ChrW(&H25BA) & " " & modI18N.GetText("PAGE HEADLINE") ' what you tell the user
        SECTION_MARKER_H1 = modI18N.GetText("PAGE HEADLINE") & " 1" ' what you tell the user
        SECTION_MARKER_H2 = modI18N.GetText("PAGE HEADLINE") & " 2" ' what you tell the user
        SECTION_MARKER_H3 = modI18N.GetText("PAGE HEADLINE") & " 3" ' what you tell the user
        SECTION_MARKER_H4 = modI18N.GetText("PAGE HEADLINE") & " 4" ' what you tell the user
        SECTION_MARKER_H5 = modI18N.GetText("PAGE HEADLINE") & " 5" ' what you tell the user
        SECTION_MARKER_H6 = modI18N.GetText("PAGE HEADLINE") & " 6" ' what you tell the user
        SECTION_MARKER_COMMON = modI18N.GetText("PAGE HEADLINE") ' what you tell the user

        'ID_LINK = ChrW(&H2192) & " " & modI18N.GetText("LINK")
        'ID_SELECT = ChrW(&H2206) & " " & modI18N.GetText("SELECT ITEM")
        'ID_BUTTON = ChrW(&H2206) & " " & modI18N.GetText("INPUT BUTTON")
        'ID_CHECKBOX = ChrW(&H2206) & " " & modI18N.GetText("CHECKBOX")
        'ID_RADIO = ChrW(&H2206) & " " & modI18N.GetText("RADIO BUTTON")
        'ID_TEXTBOX = ChrW(&H2206) & " " & modI18N.GetText("TEXT INPUT BOX")
        'ID_PASSWORD = ChrW(&H2206) & " " & modI18N.GetText("PASSWORD INPUT")
        'ID_SUBMIT = ChrW(&H2206) & " " & modI18N.GetText("SUBMIT BUTTON")
        'ID_FILE = ChrW(&H2206) & " " & modI18N.GetText("FILE SELECT")
        'ID_RESET = ChrW(&H2206) & " " & modI18N.GetText("RESET BUTTON")
        'ID_TEXTAREA = ChrW(&H2206) & " " & modI18N.GetText("TEXT AREA INPUT")
        ID_LINK = modI18N.GetText("LINK")
        ID_SELECT = modI18N.GetText("SELECT ITEM")
        ID_BUTTON = modI18N.GetText("INPUT BUTTON")
        ID_CHECKBOX = modI18N.GetText("CHECKBOX")
        ID_RADIO = modI18N.GetText("RADIO BUTTON")
        ID_TEXTBOX = modI18N.GetText("TEXT INPUT BOX")
        ID_EMAILINPUT = modI18N.GetText("EMAIL INPUT BOX")
        ID_RANGEINPUT = modI18N.GetText("RANGE INPUT")
        ID_PASSWORD = modI18N.GetText("PASSWORD INPUT")
        ID_SUBMIT = modI18N.GetText("SUBMIT BUTTON")
        ID_FILE = modI18N.GetText("FILE SELECT")
        ID_RESET = modI18N.GetText("RESET BUTTON")
        ID_TEXTAREA = modI18N.GetText("TEXT AREA INPUT")

        ID_TABLE = modI18N.GetText("TABLE")
        ID_DISABLED = modI18N.GetText("[DISABLED]") & " "
		ID_READONLY = modI18N.GetText("[READ-ONLY]") & " "
		ID_CHECKED = modI18N.GetText("[CHECKED]") & " "
		ID_NOTCHECKED = modI18N.GetText("[NOT CHECKED]") & " "
		ID_SELECTED = modI18N.GetText("[SELECTED]") & " "
        ID_NOTSELECTED = modI18N.GetText("[NOT SELECTED]") & " "
        'ID_VIDEO = ChrW(&H263A) & " " & modI18N.GetText("VIDEO")
        'ID_AUDIO = ChrW(&H266B) & " " & modI18N.GetText("AUDIO")
        ID_VIDEO = modI18N.GetText("VIDEO")
        ID_AUDIO = modI18N.GetText("AUDIO")
    End Sub
End Module