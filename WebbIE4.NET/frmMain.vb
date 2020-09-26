Option Strict On
Option Explicit On

'Compiler options:
'   MONITOR_FOCUS   If you add tmrFocus (or something) to the form and enable this then
'       you'll track what has focus, in case of the MSAA bug.



Public Class frmMain

    'WebbIE

    '.Net version of WebbIE

    'LICENCE
    'This program is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General Public License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.
    '
    '    This program is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General Public License for more details.
    '
    '    You should have received a copy of the GNU General Public License
    '    along with this program.  If not, see <http://www.gnu.or5g/licenses/>.

    '"WebbIE" is copyright 2002 Paul Blenkhorn and Gareth Evans, and is released by them under the GNU Public
    'Licence Version 3. A copy is available at http://www.gnu.org/licenses/gpl.txt
    'WebbIE includes code copyright 2002, 2005 Alasdair King, and where this is the case it is also licenced
    'under the GNU Public Licence Version 3.
    'Copyright 2007 Alasdair King.

    'Name: Alasdair King alasdair@webbie.org.uk

    '4.0 Conversion to .Net
    '       Removed:
    '           Ability to report errors.
    '           Bespoke font form
    '           Help forms.
    '           Table display.
    '           Bookmarks
    '           External browser
    '           Ability to start a dial-up connection automatically and monitor it.
    '           Display of online status
    '           Any kind of Ajax tracking.
    '           Readability
    '           Fetching PDF files from Google.
    '4.0.10 Jan 2013    Shipped as Beta
    '4.0.11 Jan 2013    Bugfixes (registry error preventing startup, -reinstall not working)
    '4.0.12 Jan 2013    Bugfixes to list of links and crop functions.
    '4.1.0 Feb 2013    Added back a way to see the web page.
    '                   Removed image save.
    '                   Reverted to just showing the browser home page.
    '4.2.0 Feb 2013     Fixed crash if homepage is "http://"
    '                   Added Favorites window on B (or Ctrl+B) that shows all your favourites in a 
    '                   treeview window. There are now three ways to get at your favourites: the menu,
    '                   the home page, and the popup window. I like the last best, but there's no reason
    '                   to drop the others as yet. I should get rid of the home page, but someone wrote
    '                   to say they love it, so I don't want to!
    '                   Made Print print the web page, not the text area.
    '                   Added a keyhook so that Escape reliably closes the frmWeb view.
    '                   Made the frmWeb view fullscreen.
    '                   Now a choice for home page: the IE home page, or the WebbIE home page (default)
    '                   Lots of work to suppress navigation errors.
    '                   Should now more accurately stop navigation sounds when page finishes loading.
    '4.3.0  26 Feb 2013 Fixed uncrop bug - crash when failing to place caret.
    '                   Made crop use MAIN element, if found.
    '                   Added input type="email" support (HTML5)
    '                   Added input type="range" support (HTML5)
    '                   Added progress element support (HTML5)
    '                   Improved labelling of controls: many more controls will have working labels,
    '                   and there will be less duplication of label text.
    '                   Fixed bug with toolbar option not hiding/showing toolbar.
    '                   Added back in the ability to number links in WebbIE (off by default)
    '4.3.1  27 Mar 2013 Made loading smoother when loading Favorites menu - better feel.
    '                   Made loading much faster.
    '                   Fixed toolbar display bug. 
    '4.3.3  7 May 2013  Put in comprehensive error handling. This means that users should never
    '                   see any error messages unless there is a comprehensive Windows system
    '                   crash. Instead, things will just not work or work oddly. This is to
    '                   increase confidence for users: ideally they get an error message which
    '                   would help them to fix it, or report it to me to get fixed. However,
    '                   this generally doesn't work as a process. 
    '                   Made "show IE homepage" the default again because of negative feedback.
    '                   Made "localhost" work as URL.
    '                   Added support for aria-label and aria-labelledby (found on Facebook)
    '4.3.4  7 June 2013 Made about:blank load faster.
    '                   Fixed textarea input so it works(!)
    '                   Made textarea not label with Name, because that's code and sucks.
    '                   Made text notruntogether on Facebook.
    '4.3.5 14 June 2013 Settings now update from one version to the next.
    '                   Defaults to Maximimised window state.
    '                   Won't check for updates more than once a day.
    '4.3.6 23 Feb 2014  Fixed a link text bug, should label links better.
    '                   Fixed a focus problem with youtube.com.
    '                   Fixed a "finished navigation" problem with youtube.com.
    '                   No longer lets you change disabled text inputs etc.
    '                   Made manufacturer "Accessible and WebbIE"
    '                   Added online activation
    '                   Updated updater DLL
    '4.4.0  2 Mar 2014  Fixed a bug with making WebbIE default browser: should now work.
    '                   Fixed font and appearance for text view, select forms.
    '                   Added colour dialog option to set foreground/background colours.
    '                   Now uses IE11 rendering.
    '                   Made most forms maximised.
    '4.4.1  6 June 2014 Fixed many input buttons not working.
    '4.5.0  22 Dec 2014 Fixed case-sensitive urls not working when typed into address bar.
    '                   Added ability to download/open VIDEO and AUDIO HTML5 elements directly in your
    '                   default media player: just hit Open. Doesn't work with embedded data URI elements.
    '                   Shortcut keys for media play: Ctrl+P play, Ctrl+O open, Space stop.
    '                   If you hold Control while clicking Refresh, or Shift while selecting Ctrl+R, 
    '                   then you get a full page refresh in the web browser. (Previously you couldn't
    '                   get a full browser refresh, just WebbIE re-displaying)
    '                   Can now open saved MHT files from File > Open.
    '                   Added TeamViewer download link.
    '4.5.1              Fixed focus going to webbrowser, not text area (e.g. if you go to google.com)
    '                   Took out display of frames that are hard to get at (security exceptions) which
    '                   will break some pages but fixes lots of navigation complexity and loads faster. 
    '                   ... And they are probably all ads anyway. 
    '                   Fixed activation DLL counting every use as an activation.
    '4.5.2              Fixed focus getting lost on Google search form when Next Page Of Results used. May 2016
    '                       (At least, I think I have fixed this, I can't test it because I'm in a hospital with 
    '                       no wi-fi.) TODO Test focus on Google Search!
    '4.5.3              Updated frame handling a bit, will now detect and refresh when an IFRAME loads.
    '                   Links which are also headings now displayed correctly.
    '4.5.4, 12 Mar 2017
    '                   Fixed a bug With the password input Not applying the password If you pressed Return instead
    '                   of clicking. 
    '                   Added a check for navigation not terminating and forcing page rendering.
    '                   Removed WebbIE update call if WebbIEUpdater.dll not found - for Windows Store.
    '                   Took out activation code, since I've never used it. 
    '5.0.0 22 Dec 2018  
    '                   Fixed web search. 


    'TODO
    '   Highlighting for headings.
    '   Spellcheck on text areas.
    '   ARIA hidden attribute (see Facebook front page)
    '   Tab through links.
    '   Escape/stop/back interrupt navigation.
    '   ARIA role main support
    '   RSS link when we have finished RSS Reader
    '   Hide addressbar and statusbar options
    '   Back button shouldn't try to go back to pages that autoforwarded (e.g. Google results)
    '
    'Little stuff:
    '   RSS button should only enable when RSS
    '

    ''' <summary>
    ''' Indicates that there should be a blank line at this point in the output. 
    ''' </summary>
    Private Const BLANK_LINE_MARKER As String = "jweofijweoifj"

#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
    'For capturing focus changes that break MSAA.
    Private focusHandler As System.Windows.Automation.AutomationFocusChangedEventHandler = Nothing
    Private mControlName As String ' The result of UIA telling me the name of the current element.
    Private mControlType As String ' the result of UIA telling me the type of the current element, e.g. "menu item"
#End If
    Public Enum IEVisible
        Toggle
        MakeIEVisible
        MakeTextVisible
    End Enum

    ''' <summary>
    ''' Indicates that the page has had links etc. stripped from it when set.
    ''' </summary>
    ''' <remarks></remarks>
    Private mCropped As Boolean

    'Parsing State Variables
    Private mOutput As System.Text.StringBuilder
    Private Const NUMBER_CHARS_AFTER_LINK_PERMITTED As Integer = 500 ' the number of chars permitted

    Private mForceFollowAddress As Boolean ' If set, you must follow the href for a link, not do a click.
    Private mForceDownloadLink As Boolean ' if set, you must download the file this link points at, not do a click
    Private ReadOnly mProcessNextDocumentComplete As Boolean ' exception to the general "process on _onload" rule from 3.11.
    '   If set, we have an internal navigation, so we should start processing on _DocumentComplete instead.
    Private mJustDidLink As Boolean ' indicates that we've just done a link, so can add a very
    'limited amount of text after a link before a newline is applied.
    Private mNewline As String ' whether we've just done a newline in the page
    Private mInArticle As Integer ' Whether we're in an article section. Not Boolean because they nest.

    Private mGoingBack As Boolean ' used to indicate we're going back
    Private mobjNavigationRecord As System.Collections.Generic.Dictionary(Of String, Integer) ' stores line numbers of where we've been before
    Private mblnErrorPage As Boolean ' whether there has been an error detected by modGlobals.gWebHost.webMain
    Private mHeadingLevel As String ' the tags for the heading level found
    'on the page: we check first that we can find a heading, and if
    'we can, set the heading level to that.
    Private mJustDidHeading As Boolean ' indicates that we've just found the heading for the page.
    'If the following text is a link, we have to handle it a bit differently.
    Private mSeekingInternalTarget As Boolean ' indicates that we are looking for
    'an internal target on this page parse called...
    Private mInternalTarget As String
    Private mControlKeyPressed As Boolean
    Private mShiftKeyPressed As Boolean
    Private mElementWithFocus As mshtml.IHTMLElement ' For Ajax handling. Tracks the element that has suddenly
    '   got the focus in IE and should therefore now get the focus in WebbIE.
    Private mSeekingFocusElement As Boolean ' indicates we are looking for an internal target node
    '   on this page: mElementWithFocus in fact.
    Private mInternalLinkNavigationStart As Integer ' The line number we started on when we did an internal
    '   link navigation.
    Private ReadOnly mBBCIPlayer As mshtml.IHTMLElement ' the embedded BBC iPlayer Flash object.
    Private mTerminateParsing As Boolean ' if set, stop parsing and present contents of the page. This isn't
    '   very neat, but we do it because of sites like http://www.epoznan.pl/, which fails to stop parsing.
    '   It keeps adding new nodes.
    Private mInForm As Boolean ' indicates that we are in a form.

    'Refresh levels for web browser
    Private Const REFRESH_NORMAL As Integer = 0 ' Perform a lightweight refresh that does not include sending the HTTP "pragma:nocache" header to the server.
    Private Const REFRESH_IFEXPIRED As Integer = 1 ' Perform a lightweight refresh if the page has expired.
    Private Const REFRESH_COMPLETELY As Integer = 3 ' Perform a full refresh that includes sending a "pragma:nocache" header to the server (HTTP URLs only).

    Private mFormClosing As Boolean = False

    Private Sub DoBack()
        On Error GoTo cantGoBack
        Dim lineNumber As Integer
        If btnBack.Enabled Then
            If mInternalLinkNavigationStart > -1 Then
                'aha, we're trying to do back on an internal link
                Call modAPIFunctions.SetCurrentLineIndex(txtText, mInternalLinkNavigationStart)
                mInternalLinkNavigationStart = -1
            Else
                'go back
                mGoingBack = True
                Call modAccesskeys.ClearAccessKeys()
                'Get the current line before we navigate - used
                'for back and forward. Used to be in BeforeNavigate, but
                'gets too confusing with all the frames and adverts triggering
                'BeforeNavigate
                'record the line number
                lineNumber = GetCurrentLineIndex(txtText)
                'we often get several  events for a page while it
                'loads (e.g. adverts) and we only want to record the line number when
                'there's actually something on the page, e.g. lineNumber > 1
                If lineNumber > 1 Then
                    If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.URL) Then
                        mobjNavigationRecord.Remove(modGlobals.gWebHost.URL)
                    End If
                    mobjNavigationRecord.Add(modGlobals.gWebHost.URL, lineNumber)
                End If
                'Start navigating timer.
                Me.tmrBusyAnimation.Enabled = True
                Call modGlobals.gWebHost.webMain.GoBack()
            End If
        Else
            Call PlayErrorSound()
        End If
        Exit Sub
cantGoBack:
        'can't go backwards, don't do anything
        Call StopBrowsers()
    End Sub

    Private Sub DoForward()
        On Error GoTo cantGoForward
        Dim lineNumber As Integer
        'Call PlayNavigationStartSound
        If mnuNavigateForward.Enabled Then
            mGoingBack = True
            'get the current line before we navigate - used
            'for back and forward. Used to be in BeforeNavigate, but
            'gets too confusing with all the frames and adverts triggering
            'BeforeNavigate
            'record the line number
            lineNumber = GetCurrentLineIndex(txtText)
            'we often get several  events for a page while it
            'loads (e.g. adverts) and we only want to record the line number when
            'there's actually something on the page, e.g. lineNumber > 1
            If lineNumber > 1 Then
                mobjNavigationRecord.Add(modGlobals.gWebHost.URL, lineNumber)
            End If

            Call modGlobals.gWebHost.webMain.GoForward()
        Else
            Call PlayErrorSound()
        End If
        Exit Sub
cantGoForward:
        'can't go forwards, don't do anything
        Call StopBrowsers()
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        On Error Resume Next
        Call DoBack()
    End Sub

    Private Sub btnHome_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHome.Click
        On Error Resume Next
        Call DoHome()
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        On Error Resume Next
        If NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) < 0 Then
            'User pressing Control, do IE refresh.
            Call frmWeb.webMain.Refresh(WebBrowserRefreshOption.Completely)
        Else
            'In the context of WebbIE, just refresh my rendering!
            Call RefreshCurrentPage()
        End If
    End Sub

    Private Sub btnStop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStop.Click
        On Error Resume Next
        Call StopBrowsers()
    End Sub

    Private Sub frmMain_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        On Error Resume Next
    End Sub

    Public Sub mnuLinksDownloadlink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksDownloadlinktarget.Click
        On Error Resume Next
        mForceDownloadLink = True
        Call UserPressedReturn()
        mForceDownloadLink = False
    End Sub


    Private Sub frmMain_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            'Detect CAPSLOCK is down, and if so don't do any key handling, since it's the standard screenreader control key.
            If modKeys.IsCapslockPressed Then
                Exit Sub
            End If

            Dim KeyCode As Integer = eventArgs.KeyCode
            'process user keypresses by calling the appropriate GUI component
            Dim ctrlDown As Boolean 'control is pressed
            Dim quickKeysOn As Boolean
            Dim handled As Boolean = True
            If Me.ActiveControl Is Nothing Then
                quickKeysOn = False
            ElseIf eventArgs.Alt Then
                ' If user is holding down Alt then don't do single-key action. Allows access keys to operate.
                quickKeysOn = False
            ElseIf Me.ActiveControl.Name.Contains("mnu") Then
                ' In a menu item, don't do single-key action.
                quickKeysOn = False
            Else
                quickKeysOn = (My.Settings.QuickKeys And (Me.ActiveControl.Name = Me.txtText.Name))
            End If
            ctrlDown = eventArgs.Control Or quickKeysOn
            Select Case KeyCode
                Case Is = System.Windows.Forms.Keys.Escape ' Stop loading
                    If btnStop.Enabled Then
                        Call StopBrowsers()
                    End If
                Case Is = System.Windows.Forms.Keys.F7
                    'goto headline
                    Call GotoHeadline(1)
                Case Is = System.Windows.Forms.Keys.D ' address bar
                    If ctrlDown Then
                        Call cboAddress.Focus()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.F ' Find
                    If ctrlDown Then
                        Call FindText()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.Back ' Go back
                    'check it isn't in the address field
                    If Me.ActiveControl Is Nothing Then
                        Call btnBack_Click(btnBack, New System.EventArgs())
                    ElseIf Me.ActiveControl.Name <> cboAddress.Name Then
                        Call btnBack_Click(btnBack, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.Right ' Go forward (with ALT)
                    If eventArgs.Alt Then
                        Call DoForward()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.L ' List links
                    If ctrlDown Then
                        Call ListLinks()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.Left ' Go back (with ALT)
                    If eventArgs.Alt Then
                        btnBack_Click(btnBack, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.F5 ' Refresh
                    btnRefresh_Click(btnRefresh, New System.EventArgs())
                    handled = False
                Case Is = System.Windows.Forms.Keys.Home ' Home (with Alt)
                    If eventArgs.Alt And ctrlDown And Not My.Settings.QuickKeys Then
                        Call SetHomepage()
                    ElseIf eventArgs.Alt Then
                        btnHome_Click(btnHome, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.H ' next headline
                    If ctrlDown Or quickKeysOn Then
                        KeyCode = 0
                        eventArgs.Handled = True
                        eventArgs.SuppressKeyPress = True
                        Call GotoHeadline(CInt(IIf(eventArgs.Shift, -1, 1)))
                    Else
                        handled = False
                    End If

                    'DEV: can't hard-code access keys - doesn't allow for I18N
                    '        Case Is = vbKeyD ' in case the access key doesn't work
                    '            If altDown Then
                    '                Call cboAddress.SetFocus
                    '                KeyCode = 0
                    '            End If
                    '        Case Is = vbKeyT 'in case the access key doesn't work
                    '            If altDown Then
                    '                Call txtText.SetFocus
                    '                KeyCode = 0
                    '            End If
                Case Is = System.Windows.Forms.Keys.W ' for Google search
                    If ctrlDown Then
                        KeyCode = 0
                        Call DoWebSearch()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.B ' For Favourites
                    If ctrlDown Then
                        KeyCode = 0
                        Call ShowFavoritesWindow()
                    Else
                        handled = False
                    End If
                    'We now have a bunch of shortcuts specifically for when are in QuickLinks
                    'mode, that is, when users can hit a single character to operate a function.
                Case Is = System.Windows.Forms.Keys.S 'for skip down
                    If quickKeysOn Then
                        Call DoSkip(SKIP_DOWN)
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.A ' select all
                    If quickKeysOn Then
                        Call mnuEditSelectall_Click(mnuEditSelectall, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.C ' copy
                    If quickKeysOn Then
                        Call mnuEditCopy_Click(mnuEditCopy, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.X ' cut
                    If quickKeysOn Then
                        Call EditCut()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.V ' paste
                    If quickKeysOn Then
                        Call Paste()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.U ' for skip up
                    If quickKeysOn Then
                        Call DoSkip(SKIP_UP)
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.K ' crop
                    If quickKeysOn Then
                        Call CropPage()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.R ' refresh
                    If quickKeysOn Or ctrlDown Then ' Need also ctrlDown to handle Ctrl+Shift+R
                        Call mnuNavigateRefresh_Click(mnuNavigateRefresh, New System.EventArgs())
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.E ' rss
                    If quickKeysOn Then
                        KeyCode = 0
                        Call RSS()
                    Else
                        handled = False
                    End If
                Case Is = System.Windows.Forms.Keys.F6
                    If eventArgs.Shift Then
                        Call SkipToFormElement(-1)
                    Else
                        handled = False
                    End If
                Case Else
                    handled = False
            End Select
            If handled Then
                eventArgs.Handled = True
                eventArgs.SuppressKeyPress = True
            End If
        Catch
        End Try
    End Sub

    Private Sub Paste()
        Try
            'process use clicking paste button in menu
            Dim currentControlName As String
            If Me.ActiveControl Is Nothing Then
                'Assume it's the web browser
                currentControlName = modGlobals.gWebHost.webMain.Name
            Else
                currentControlName = Me.ActiveControl.Name
            End If
            Select Case currentControlName
                Case modGlobals.gWebHost.webMain.Name
                    'Dim wb As SHDocVw.WebBrowser
                    'wb = CType(workBrowser.ActiveXInstance, SHDocVw.WebBrowser)
                    'Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_PASTE, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT)
                    Call modSendKeys.SendPaste()
                    'Call SendKeys.Send("^V")
                Case cboAddress.Name
                    cboAddress.SelectedText = My.Computer.Clipboard.GetText
                Case txtText.Name
                    Call Beep()
            End Select
        Catch
        End Try
    End Sub

    Public Sub mnuLinksFollowlinkaddress_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksFollowlinkaddress.Click
        On Error Resume Next
        mForceFollowAddress = True
        Call UserPressedReturn()
        mForceFollowAddress = False
    End Sub

    Public Sub ResizeToolbar()
        Try
            If My.Settings.ShowToolbar Then
                Me.MainToolStrip.Visible = True
                Dim ds As System.Windows.Forms.ToolStripItemDisplayStyle = CType(IIf(My.Settings.ToolbarCaptions, ToolStripItemDisplayStyle.ImageAndText, ToolStripItemDisplayStyle.Image), ToolStripItemDisplayStyle)
                For Each tsi As ToolStripItem In Me.MainToolStrip.Items
                    tsi.DisplayStyle = ds
                Next tsi

                If My.Settings.ToolbarCaptions Then
                    Me.MainToolStrip.Height = 90
                    picBusy.Image = My.Resources.timer_done_big
                Else
                    Me.MainToolStrip.Height = 74
                    picBusy.Image = My.Resources.timer_done
                End If
            Else
                ' No toolbar at all!
                Me.MainToolStrip.Visible = False
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuViewLinkinformation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error Resume Next
        Call txtText_KeyDown(txtText, New System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.F8))
    End Sub

    Public Sub mnuViewRSSNewsFeed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewRSSNewsFeed.Click
        On Error Resume Next
        Call RSS()
    End Sub

    Public Sub RSS()
        On Error Resume Next
        Call frmRSS.ShowDialog(Me)
    End Sub

    ''' <summary>
    ''' Page navigation is complete (enough) so call this to render the page. 
    ''' </summary>
    Public Sub ProcessAfterLoad()
        On Error Resume Next
        Dim lineNumber As Integer
        Call ProcessPage()
        'Right, we've just finished loading and displaying a page. Now, where
        'do we put the caret? Two factors:
        ' - there is an internal target to which we should move the caret
        ' - if we've gone forwards or backwards, put the caret where it was
        'You could go f/b to a page with an internal target, but you
        'should go to where the caret was, not where the target indicates.
        'So check f/b first.
        If mobjNavigationRecord Is Nothing Then
            mobjNavigationRecord = New System.Collections.Generic.Dictionary(Of String, Integer)
        End If
        If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.URL) Then
            'found an entry for this page.
            lineNumber = mobjNavigationRecord.Item(modGlobals.gWebHost.URL)
        End If
        If lineNumber = 0 Then
            If mSeekingInternalTarget Then
                Call GotoInternalTarget()
            End If
        Else
            'okay, we have somewhere to go.
            'Debug.Print "Decided to put caret at line:" & lineNumber
            Dim newSelectionStart As Integer = GetCharacterIndexOfLine(txtText, lineNumber)
            If newSelectionStart > -1 Then
                txtText.SelectionStart = newSelectionStart
            End If
            txtText.SelectionLength = 0
            Call txtText.ScrollToCaret()
            'This call sets the focus to the text area when the WebBrowser object steals it. You can 
            'observe this by navigating to google.com. 
            'I tried many approaches to this, but this works (testing Windows 10 Feb 2016). It didn't
            'work before but I was getting the API call wrong. Use the Analyze functions!
            Call NativeMethods.SetFocus(txtText.Handle)
        End If
    End Sub

    'Private Sub tmrProcessAfterLoad_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrProcessAfterLoad.Tick
    '    If modGlobals.gClosing Then
    '        Me.tmrProcessAfterLoad.Enabled = False
    '        Exit Sub
    '    End If

    '    Static busy As Boolean
    '    Try
    '        'Fires after the Load event fires on the main webpage.

    '        Dim lineNumber As Integer
    '        tmrProcessAfterLoad.Enabled = False

    '        'Call MessageBox.Show("tmrProcessAfterLoad")

    '        If busy Then
    '        Else
    '            busy = True
    '            Call ProcessPage()

    '            'Right, we've just finished loading and displaying a page. Now, where
    '            'do we put the caret? Two factors:
    '            ' - there is an internal target to which we should move the caret
    '            ' - if we've gone forwards or backwards, put the caret where it was
    '            'You could go f/b to a page with an internal target, but you
    '            'should go to where the caret was, not where the target indicates.
    '            'So check f/b first.
    '            If mobjNavigationRecord Is Nothing Then
    '                mobjNavigationRecord = New System.Collections.Generic.Dictionary(Of String, Integer)
    '            End If
    '            If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.webMain.Url.ToString()) Then
    '                'found an entry for this page.
    '                lineNumber = mobjNavigationRecord.Item(modGlobals.gWebHost.webMain.Url.ToString())
    '            End If
    '            If lineNumber = 0 Then
    '                If mSeekingInternalTarget Then
    '                    Call GotoInternalTarget()
    '                End If
    '            Else
    '                'okay, we have somewhere to go.
    '                'Debug.Print "Decided to put caret at line:" & lineNumber
    '                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
    '                txtText.SelectionLength = 0
    '                Call txtText.ScrollToCaret()
    '                'This call sets the focus to the text area when the WebBrowser object steals it. You can 
    '                'observe this by navigating to google.com. 
    '                'I tried many approaches to this, but this works (testing Windows 10 Feb 2016). It didn't
    '                'work before but I was getting the API call wrong. Use the Analyze functions!
    '                Call NativeMethods.SetFocus(txtText.Handle)
    '            End If
    '            busy = False
    '        End If
    '    Catch
    '        busy = False
    '    End Try
    'End Sub

    Private Sub GotoInternalTarget()
        Try
            Dim selStart As Integer = txtText.Text.IndexOf(TARGET_MARKER)
            If selStart > -1 Then
                'In theory I should be able to select the TARGET_MARKET text and then
                'delete it (txtText.SelectedText = "") but since this magically does
                'not work for no reason I can see, do with strings.
                Dim contents As String = txtText.Text.Substring(0, selStart) + txtText.Text.Substring(selStart + TARGET_MARKER.Length, txtText.Text.Length - selStart - TARGET_MARKER.Length)
                Call SetText(contents)
                txtText.SelectionStart = selStart
            End If
            mSeekingInternalTarget = False
        Catch
        End Try
    End Sub

    Private Sub tmrRefreshIfNotChange_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrRefreshIfNotChange.Tick
        If modGlobals.gClosing Then
            Me.tmrRefreshIfNotChange.Enabled = False
            Exit Sub
        End If
        Try
            tmrRefreshIfNotChange.Enabled = False
            If gExiting Then
            ElseIf gNoPageActionSoRefresh Then
                gNoPageActionSoRefresh = False
                Debug.Print("Triggered refresh!")
                Call RefreshCurrentPage()
                Call PlayDoneSound()
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrSetFocus_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrSetFocus.Tick
        On Error Resume Next
        tmrSetFocus.Enabled = False
        If modGlobals.gClosing Then
            Exit Sub
        End If
        'log "tmrSetFocus_Timer"
        If modGlobals.gFrmWebHasFocusAndItShouldNot Then
            modGlobals.gFrmWebHasFocusAndItShouldNot = False
            Call Me.Focus()
            tmrSetFocus.Enabled = True ' let's let it come back round...
        Else
            Call txtText.Focus()
        End If
    End Sub

    Private Sub mnuHelpWebbiehome_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpWebbIEOrg.Click
        On Error Resume Next
        Call GotoWebbIEDotOrgDotUK()
    End Sub

    Public Sub GotoWebbIEDotOrgDotUK()
        On Error Resume Next
        'go to home for webbie - website
        Call StartNavigating(modI18N.GetText("http://www.webbie.org.uk"))
    End Sub

    ''' <summary>
    ''' Move cursor to the next headline
    ''' </summary>
    ''' <param name="direction"></param>
    ''' <remarks></remarks>
    Private Sub GotoHeadline(ByVal direction As Integer)
        On Error Resume Next
        If GotoLineStarting(SECTION_MARKER_COMMON, direction, False) Then
            'Don't select - stops screenreaders reading line, they read the document instead.
            'Call SelectCurrentParagraph
            'okay, this will have selected the paragraph starting "Page Headline:". But
            'what if the headline is the next line as well? Happens with links.
            'Well, then, select the next line
            '        If Len(txtText.SelText) < Len(SECTION_MARKER_H1) + 10 Then
            '            nextLine = txtText.selStart + txtText.SelLength + Len(vbNewLine)
            '            nextLine = InStr(nextLine, txtText.Text, vbNewLine)
            '            If nextLine > 0 Then
            '                'found new end-of-line
            '                txtText.SelLength = nextLine - txtText.selStart
            '            End If
            '        End If
            Call txtText.ScrollToCaret()
        End If
    End Sub

    ''' <summary>
    ''' Moves cursor to line beginning with startsWith. Returns true if found.
    ''' </summary>
    ''' <param name="startsWith"></param>
    ''' <param name="direction"></param>
    ''' <param name="wrap"></param>
    ''' <returns>true if line starting with startsWith found</returns>
    ''' <remarks></remarks>
    Private Function GotoLineStarting(ByRef startsWith As String, Optional ByRef direction As Integer = 1, Optional ByRef wrap As Boolean = True) As Boolean
        Try
            Dim numberLines As Integer
            Dim i As Integer
            Dim lineContent As String
            Dim found As Boolean
            Dim currentLine As Integer
            Dim LinkLength As Integer = ID_LINK.Length

            numberLines = GetNumberOfLines(txtText)
            currentLine = GetCurrentLineIndex(txtText)
            Debug.Print("Start: " & currentLine)
            i = currentLine + direction
            'Debug.Print "Line: " & i
            While i >= 0 And i < numberLines And (Not found)
                lineContent = GetLine(txtText, i)
                If InStr(1, lineContent, startsWith, CompareMethod.Text) = 1 Or
                   InStr(1, lineContent, startsWith, CompareMethod.Text) = 1 + LinkLength + 2 Then
                    'found it
                    Debug.Print("Found")
                    found = True
                    txtText.SelectionStart = GetCharacterIndexOfLine(txtText, i)
                    txtText.SelectionLength = 0
                    'Call ScrollToCursor(txtText)
                    'Call PlayDoneSound ' Don't play the done sound, it's really intrusive. Aug 2010.
                End If
                i += direction
                'Debug.Print "Line: " & i
            End While
            If Not found And wrap Then
                If direction > 0 Then
                    i = 0
                Else
                    i = numberLines - 1
                End If
                While i >= 0 And i <= currentLine And Not found
                    lineContent = GetLine(txtText, i)
                    If InStr(1, lineContent, startsWith, CompareMethod.Text) = 1 Then
                        'found it
                        found = True
                        txtText.SelectionStart = GetCharacterIndexOfLine(txtText, i)
                        txtText.SelectionLength = 0
                        'Call ScrollToCursor(txtText)
                        'Call PlayDoneSound ' Don't play the done sound, it's really intrusive. Aug 2010.
                    End If
                    i += direction
                End While
            End If
            If Not found Then Call PlayErrorSound()
            GotoLineStarting = found
        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Sub mnuLinksSkipup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksSkipup.Click
        On Error Resume Next
        Call DoSkip(SKIP_UP)
    End Sub

    Public Sub mnuNavigateGotoform_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateGotoform.Click
        On Error Resume Next
        'Causes the cursor to jump to the next input element over any links
        Call SkipToFormElement(1)
    End Sub

    Private Sub SkipToFormElement(ByRef direction As Integer)
        Try
            Dim lineNumber As Integer
            Dim found As Boolean
            Dim line As String
            Dim numberLines As Integer

            lineNumber = GetCurrentLineIndex(txtText) + direction
            numberLines = GetNumberOfLines(txtText)
            found = False
            While lineNumber < numberLines And lineNumber >= 0 And Not found
                line = GetLine(txtText, lineNumber)
                If LineIsForm(line) Then
                    'found form element
                    found = True
                Else
                    'keep going
                    lineNumber += direction
                End If
            End While
            If found Then
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                txtText.SelectionLength = 0
                'Call ScrollToCursor(txtText)
                'Don't select - stops screenreaders reading line, they read the document instead.
                'Call SelectCurrentParagraph
            Else
                Call PlayErrorSound()
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuLinksPreviousLink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksPreviouslink.Click
        Try
            'go up to previous link

            Dim lineNumber As Integer
            Dim linkFound As Boolean

            lineNumber = GetCurrentLineIndex(txtText) - 1
            linkFound = False
            While lineNumber >= 0 And Not linkFound
                If txtText.Lines(lineNumber).Contains(ID_LINK) Then
                    'found a link
                    linkFound = True
                Else
                    lineNumber -= 1
                End If
            End While
            If linkFound Then
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                txtText.SelectionLength = 0
                'txtText.SelLength = Len(Trim(modAPIFunctions.GetCurrentLine(txtText)))
                'Call ScrollToCursor(txtText)
            Else
                Call PlayErrorSound()
            End If
        Catch
        End Try
    End Sub



    Public Sub mnuLinksSkiplinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksSkipDown.Click
        On Error Resume Next
        Call DoSkip(SKIP_DOWN)
    End Sub

    ''' <summary>
    ''' returns true if the line is a form element, false if not
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function LineIsForm(ByRef line As String) As Boolean
        On Error Resume Next

        LineIsForm = True
        If InStr(1, line, ID_SELECT, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_BUTTON, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_CHECKBOX, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_RADIO, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_TEXTBOX, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_PASSWORD, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_SUBMIT, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_FILE, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_RESET, CompareMethod.Binary) = 1 Then
        ElseIf InStr(1, line, ID_TEXTAREA, CompareMethod.Binary) = 1 Then
        Else
            'nope, this isn't a form element
            LineIsForm = False
        End If
    End Function

    ''' <summary>
    ''' returns true if the line is text, false if it is a link or form element or similar: used by mnuViewCroppage
    ''' </summary>
    ''' <param name="line"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function LineIsText(ByRef line As String) As Boolean
        On Error Resume Next
        LineIsText = False
        If line.Contains(ID_LINK) Then
            'found a link!
        ElseIf line.StartsWith(ID_SELECT, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_BUTTON, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_CHECKBOX, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_RADIO, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_TEXTBOX, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_PASSWORD, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_SUBMIT, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_FILE, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_RESET, StringComparison.InvariantCulture) Then
        ElseIf line.StartsWith(ID_TEXTAREA, StringComparison.InvariantCulture) Then
        Else
            'yep, we're going to use this line
            LineIsText = True
        End If
    End Function

    ''' <summary>
    ''' strip out link content from the page, hopefully leaving only main content: may trim out link content
    ''' occuring as part of main content, but should remove all the navigation info
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CropPage()
        Try
            'DEV: new version Jan 2007: doesn't just skip links, does some clever
            'text processing. See modParseAsText
            Static contents As String ' the old page before cropping
            Static preCropPosition As Integer ' Where the caret was before we cropped.
            Static where As String 'where we were on the page, roughly
            Dim i As Integer
            Dim localTextStart As Integer
            Dim localTextEnd As Integer
            Dim done As Boolean
            Dim targetLine As Integer
            Dim found As Integer
            Dim parsedAsText As String
            Dim hasCroppedWithArticle As Boolean
            Dim hasCroppedWithMain As Boolean
            Dim croppedContents As String = ""

            If mCropped Then
                'already cropped, uncropping. Get the text around the cursor, not
                'newlines, and try to find it.
                'Get leftwards
                localTextStart = Me.txtText.SelectionStart
                done = False
                While Not done
                    Try
                        localTextStart -= 1
                        If localTextStart <= 0 Then
                            localTextStart = 0
                            done = True
                        ElseIf AscW(Me.txtText.Text.Substring(localTextStart, 1)) < 32 Then
                            localTextStart += 1
                            done = True
                        ElseIf localTextStart < Me.txtText.SelectionStart - 5 Then
                            done = True
                        End If
                    Catch
                        done = True
                    End Try
                End While
                'get rightwards
                done = False
                localTextEnd = Me.txtText.SelectionStart
                While Not done
                    Try
                        localTextEnd += 1
                        If localTextEnd >= Me.txtText.Text.Length - 1 Then
                            localTextEnd = Me.txtText.Text.Length - 1
                            done = True
                        ElseIf AscW(Me.txtText.Text.Substring(localTextEnd, 1)) < 32 Then
                            localTextEnd -= 1
                            done = True
                        ElseIf localTextEnd > Me.txtText.SelectionStart + 25 Then
                            done = True
                        End If
                    Catch
                        done = True
                    End Try
                End While
                'get text to find. Put in some checks: got some reports of problems and hangs
                'in here.
                If localTextStart < 0 Then localTextStart = 0
                If localTextEnd > Me.txtText.TextLength - 1 Then localTextEnd = Me.txtText.TextLength - 1
                If localTextStart >= localTextEnd Then
                    'OK, the attempt to get a string failed. 
                    where = ""
                Else
                    where = Me.txtText.Text.Substring(localTextStart, localTextEnd - localTextStart)
                End If

                'replace text with old text
                Call SetText(contents)
                If where.Length > 0 Then
                    'With luck the "where" string, if it exists at all (if it
                    'is greater than zero characters in length) should be a substring of the contents
                    'string. However, because the parsing of cropped and uncropped is subtly different,
                    'this will sometimes be the case. So have to check. 
                    Dim selStart As Integer = Me.txtText.Text.IndexOf(where)
                    If selStart > -1 Then
                        'Good, we successfully created a real substring.
                        txtText.SelectionStart = selStart
                    Else
                        'Nope, we didn't create a substring, so we don't know where to place the cursor
                        'exactly. Put it back where it was before we cropped so the user doesn't get lost.
                        If preCropPosition > -1 Then
                            txtText.SelectionStart = preCropPosition
                        End If
                    End If

                End If
                'Don't select - stops screenreaders reading line, they read the document instead.
                'Call SelectCurrentParagraph
                mnuViewCroppage.Text = modI18N.GetText("&Crop page")
                mCropped = False
            Else
                'not cropped, cropping now.
                contents = Me.txtText.Text
                preCropPosition = Me.txtText.SelectionStart
                'Heuristic cropping.
                'first, find and remember the line we're going to restore if we uncrop
                targetLine = 0
                If LineIsText(GetCurrentLine(Me.txtText)) Then
                    'hurrah, a text line: this is nice and easy
                    targetLine = GetCurrentLineIndex(Me.txtText)
                Else
                    'a non-text line, which therefore will be lost in the processing: have to find the next line
                    'that is text
                    targetLine = -1
                    For i = GetCurrentLineIndex(Me.txtText) To GetNumberOfLines(Me.txtText)
                        If LineIsText(GetLine(Me.txtText, i)) Then
                            'good, found a text line: we'll use that
                            targetLine = i 'indicates we're using this system
                            Exit For
                        End If
                    Next i
                    'NB: if we didn't find somewhere, try working backwards: we'll eventually end up at line 1 if
                    'nowhere else (which contains the document title and therefore is always a text line)
                    If targetLine = -1 Then ' initial value
                        For i = GetCurrentLineIndex(Me.txtText) To 0 Step -1
                            If LineIsText(GetLine(Me.txtText, i)) Then
                                'good, found a text line: we'll use that
                                targetLine = i 'indicates we're using this system
                                Exit For
                            End If
                        Next i
                    End If
                End If
                If targetLine = -1 Then targetLine = 0 ' failed to match anything!
                'Assertion: by this point, targetLine is the line of text that we should end up on when we finish cropping, from 0 upwards.
                'Remember the text of this line so we can find it again.
                where = GetLine(Me.txtText, targetLine).Trim().Replace(vbNewLine, "").Replace(Chr(0), "")

                'Now do the crop. Three mechanisms.
                'HTML5 cropping: 
                If gPageHasMain Then
                    Dim articleStart As Integer = contents.IndexOf("►")
                    Dim articleEnd As Integer = contents.IndexOf("◄")
                    If articleEnd = -1 Then articleEnd = contents.Length - 1
                    If articleStart > 0 And articleEnd > 0 And articleEnd > articleStart Then
                        hasCroppedWithMain = True
                        croppedContents = contents.Substring(articleStart + 1, articleEnd - articleStart - 1)
                    End If
                End If
                If gPageHasAnArticle And Not hasCroppedWithMain Then
                    Dim articleStart As Integer = 0
                    Dim articleEnd As Integer = 0

                    articleStart = contents.IndexOf(modI18N.GetText("Article Start"))
                    articleEnd = contents.IndexOf(modI18N.GetText("Article End"))
                    Dim stopLoop As Boolean = False
                    While articleStart > 0 And articleEnd > 0 And articleEnd > articleStart And Not stopLoop
                        Try
                            articleStart += modI18N.GetText("Article Start").Length
                            croppedContents &= contents.Substring(articleStart, articleEnd - articleStart)
                            articleStart = contents.IndexOf(modI18N.GetText("Article Start"), articleStart + 1)
                            articleEnd = contents.IndexOf(modI18N.GetText("Article End"), articleEnd + 1)
                            'OK, did at least one chop.
                            hasCroppedWithArticle = True
                        Catch
                            hasCroppedWithArticle = False
                            stopLoop = True
                        End Try
                    End While
                End If
                If hasCroppedWithArticle Then
                    'OK, cropping work, that'll do!
                    Call SetText(croppedContents)
                ElseIf hasCroppedWithMain Then
                    'OK, main worked:
                    Call SetText(croppedContents)
                Else
                    parsedAsText = modParseAsText.ParseDocs 'in fact only works on
                    'webbrowser.Document, so won't handle any frames. I'm going to
                    'argue this is a good thing for the moment: less nonsense.
                    'We'll have to see experimentally if this is the case.
                    If parsedAsText.Length = 0 Then
                        'Ah, the new-style parsing - removing frames - returned absolutely nothing.
                        'Instead, parse the page directly.
                        Call CropOriginalPage()
                    Else
                        Call SetText(parsedAsText)
                    End If
                End If

                'OK, done the crop. Now place the cursor where we decided it should go, which was
                'shown by where. Unless where is ""!
                If where.Length = 0 Then
                    found = -1
                Else
                    found = txtText.Text.IndexOf(where)
                    While found = -1 And where.Length > 2
                        where = where.Substring(1, where.Length - 2) '  Mid(where, 2, Len(where) - 2)
                        found = txtText.Text.IndexOf(where, System.StringComparison.InvariantCulture)
                    End While
                End If
                If found > -1 Then txtText.SelectionStart = found
                'Don't select - stops screenreaders reading line, they read the document instead.
                'Call SelectCurrentParagraph
                mnuViewCroppage.Text = modI18N.GetText("Un&crop page")
                mCropped = True
                'DEV: 3.2.0 and earlier version, took out lines of links. This therefore worked
                'exactly like Skip Links, but broke the page. Better to do some clever text
                'processing, maybe? See above.
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuViewCroppage_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewCroppage.Click
        On Error Resume Next
        Call CropPage()
    End Sub

    ''' <summary>
    ''' 'crops the original text page according to the pre-3.2 mechanisms. This supports pages with
    ''' lots of IFRAMES and other contructs that prevent me reparsing the page
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CropOriginalPage()
        Try
            Dim i As Integer
            Dim content() As String
            Dim output As String = ""
            Dim targetLine As Integer

            'cropping links
            'first, find and remember the line we're going to restore if we uncrop
            targetLine = 0
            If LineIsText(Trim(Replace(GetCurrentLine(txtText), Chr(0), " "))) Then
                'hurrah, a text line: this is nice and easy
            Else
                'a non-text line, which therefore will be lost in the processing: have to find the next line
                'that is text
                For i = GetCurrentLineIndex(txtText) To GetNumberOfLines(txtText)
                    If LineIsText(Trim(Replace(GetLine(txtText, i), Chr(0), " "))) Then
                        'good, found a text line: we'll use that
                        targetLine = i 'indicates we're using this system
                        Exit For
                    End If
                Next i
                'NB: if we didn't find somewhere, try working backwards: we'll eventually end up at line 1 if
                'nowhere else (which contains the document title and therefore is always a text line)
                If targetLine = 0 Then ' initial value
                    For i = GetCurrentLineIndex(txtText) To 0 Step -1
                        If LineIsText(Trim(Replace(GetLine(txtText, i), Chr(0), " "))) Then
                            'good, found a text line: we'll use that
                            targetLine = i 'indicates we're using this system
                            Exit For
                        End If
                    Next i
                End If
                'assertion: targetLine holds the line we should move the cursor to. So do so!
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, targetLine)
            End If
            txtText.SelectionLength = 0
            txtText.SelectedText = TARGET_MARKER ' mark where the cursor was
            content = Split(txtText.Text, vbNewLine) ' split the content into lines
            'now chop out lines that aren't text
            'DEV: did try something with scoring, didn't work too well: need to work
            'at text block level if this is tried again (e.g. working out A in DIV
            'with lots of text versus A in LI with no text)
            For i = 0 To content.Length - 1 ' iterate through removing lines that aren't text
                'find out if this is/was the target line
                If LineIsText(content(i)) Then
                    'yep, we're going to use this line if it's long enough
                    If Len(content(i)) > 4 Then ' note "magic number": this removes
                        'lines that are just punctuation points, such as bullets
                        output = output & content(i) & vbNewLine
                    End If
                Else
                    'nope, we're not
                End If
            Next i
            SetText(output)
            Dim selStart As Integer = output.IndexOf(TARGET_MARKER, 0, StringComparison.InvariantCulture)
            If selStart > -1 Then
                txtText.SelectionStart = selStart
                txtText.SelectionLength = TARGET_MARKER.Length
                txtText.SelectedText = ""
            Else
                Call SetText(txtText.Text.Replace(TARGET_MARKER, ""))
            End If
        Catch
        End Try
    End Sub

    Private Sub SetText(text As String)
        Try
            Call StopBusyAnimation()
            Call txtText.Clear()
            txtText.SelectionIndent = 10
            txtText.SelectionRightIndent = 10
            txtText.Text = text
        Catch ex As ObjectDisposedException
            'We get this if you close WebbIE when the web browser is still navigating,
            'so the txtText object has been disposed, but the web browser events are
            'being cleaned up. So ignore, we're exiting.
        End Try
    End Sub

    Private Function GetUserFriendlyURL(ByRef url As String) As String
        Try
            Dim found As Integer
            Dim friendlyURL As String

            friendlyURL = url
            'take off ? queries
            found = InStr(1, friendlyURL, "?")
            If found > 0 Then
                friendlyURL = Microsoft.VisualBasic.Left(friendlyURL, Len(friendlyURL) - found - 1)
            End If
            'take off //
            found = InStr(1, friendlyURL, "//")
            If found > 0 Then
                friendlyURL = Microsoft.VisualBasic.Right(friendlyURL, Len(friendlyURL) - found - 1)
            End If
            'take off trailing /
            If Microsoft.VisualBasic.Right(friendlyURL, 1) = "/" Then
                friendlyURL = Microsoft.VisualBasic.Left(friendlyURL, Len(friendlyURL) - 1)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object GetUserFriendlyURL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetUserFriendlyURL = friendlyURL
        Catch
            Return url
        End Try
    End Function

    Public Sub mnuViewSource_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewSource.Click
        Try
            Dim doc3 As mshtml.IHTMLDocument3
            Dim path As String

            doc3 = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.IHTMLDocument3)
            frmTextView.Text = modI18N.GetText("Page HTML Source")
            'How do I work out the mimetype?
            'Using the mshtml interface. But why bother?
            path = System.IO.Path.GetTempFileName
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(path)
            Call sw.Write(doc3.documentElement.outerHTML)
            Call sw.Close()
            Call Shell("notepad.exe """ & path & """", AppWinStyle.NormalFocus)
        Catch
        End Try
    End Sub

    Private Sub cboaddress_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cboAddress.KeyPress
        Try
            If Microsoft.VisualBasic.AscW(eventArgs.KeyChar) = Keys.Return Then 'if return is hit
                eventArgs.Handled = True
                If cboAddress.SelectedIndex > -1 Then
                    cboAddress.Text = CStr(cboAddress.Items(cboAddress.SelectedIndex))
                End If
                Call UserEnteredURL(cboAddress.Text.Trim) ' Don't change case! Breaks Unix servers.
            End If
        Catch
        End Try
    End Sub

    Public Sub UserEnteredURL(ByVal targetURL As String)
        Try
            If targetURL = "" Then Exit Sub ' nowhere to go if the address bar is empty
            'okay, so there is something to get in the address box.
            'If it looks like a url or local file, go get it.
            'Otherwise, start a Google search.
            If targetURL = "home" Then
                'go to home
                Call btnHome_Click(btnHome, New System.EventArgs())
                'This "check if it's a string" is nice in theory, but there
                'are too many dumb things people can type or paste into the address
                'box for it to work - odd characters, spaces, local paths.
                'ElseIf System.Uri.IsWellFormedUriString(targetURL, UriKind.Absolute) Then
                '    'it's definitely a url or local file or config line
                '    Call StartNavigating(targetURL) 'navigate to specified address
                'ElseIf System.Uri.IsWellFormedUriString("http://" & targetURL, UriKind.Absolute) Then
                '    'Yep, again, looks like a URL.
                '    Call StartNavigating(targetURL)
                'Else
            ElseIf targetURL.ToLowerInvariant().Trim = "localhost" Then
                'Local web server!
                Call StartNavigating("localhost")
            ElseIf targetURL.Contains(".") Then
                'Looks ENOUGH like a web address!
                Call StartNavigating(targetURL)
            ElseIf targetURL.ToLowerInvariant().Trim = "about:blank" Then
                'Blank page.
                Call StartNavigating("about:blank")
            Else
                'No idea what this is, do a google search.
                Call frmGoogle.StartSearch(targetURL)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' UpgradeSettings. This function migrates your application's settings from the previous
    '''    /// version, if any, to this one. This is because Properties.Settings are saved to a 
    '''    /// different user folder with every version, so unless you explicitly call this function
    '''    /// then user settings will be lost with every upgrade.
    '''    /// You must create a String setting called "LastVersionRun"
    '''    /// Alasdair 11 June 2013
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpgradeSettings()
        Try
            If My.Settings.LastVersionRun <> Application.ProductVersion Then
                My.Settings.Upgrade()
                My.Settings.Reload()
                My.Settings.LastVersionRun = Application.ProductVersion
                My.Settings.Save()
            End If
        Catch
            'MessageBox.Show("Error in UpgradeSettings. Have you created a property called \"LastVersionRun\"?");
        End Try
    End Sub
    ' End of UpgradeSettings()

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            'load settings, displays, bookmarks when the program starts up.
            Dim commandLine As String
            Dim helpIndex As String
            Dim startedNavigating As Boolean
            Dim filePath As String = ""

            'Upgrade settings
            Call UpgradeSettings()

            'Check for updates. Only do this is WebbIEUpdater.dll is found. This is so that we can do a 
            'build without the updater for the Windows 10 Store. 
            Dim exeFolder As String = System.IO.Path.GetDirectoryName(Application.ExecutablePath)
            If System.IO.File.Exists(exeFolder & "\WebbIEUpdater.dll") Then
                Call WebbIE.Updater.CheckForUpdates("https://www.webbie.org.uk/webbrowser/updates.xml")
            End If

            'Do Windows XP style.
            Call Application.EnableVisualStyles()

            'Stop Javascript errors showing up in our web browsers.
            Call DisableScriptErrors()

            'Instantiate the WebBrowser object we'll use.
            modGlobals.gWebHost = New frmWeb

            While modGlobals.gWebHost Is Nothing
                Call Threading.Thread.Sleep(100)
                Application.DoEvents()
            End While
            Call modGlobals.gWebHost.Show()

            txtText.ReadOnly = True 'make txtText non-editable.
            txtText.DetectUrls = False ' Leave webbie to do this!
            txtText.ScrollBars = RichTextBoxScrollBars.Vertical

            'set the tabs correctly.
            Call SetupTabs()

            'language settings
            'get the language settings for the default character set for the user's system
            Call modCharacterSupport.InitSystemLocale()
            Call modCharacterSupport.InitCharsetMappings()

            'Load data structures
            Call LoadDataStructures()

            Call modGlobals.Initialise()

            staMain.Items.Item(0).Text = modI18N.GetText("Idle")

            txtText.Font = My.Settings.DisplayFont
            cboAddress.Font = My.Settings.DisplayFont

            'Do colours
            Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))

            modGlobals.gWebHost.webMain.ScriptErrorsSuppressed = Not My.Settings.AllowMessages

            Call ResizeToolbar()

            'query the registry for the current IE homepage
            gCurrentHomepage = RetrieveStartPage()
            'If we don't use the IE homepage, disable the ability to change it.
            mnuOptionsSetHomepage.Enabled = My.Settings.UseIEHomepage

            Call modI18N.DoForm(Me)

            ' Make full-screen or normal.
            Me.WindowState = My.Settings.WindowState

            Call cboAddress.Items.Clear()

            'check for commandline request to go somewhere.
            commandLine = Microsoft.VisualBasic.Command().ToLowerInvariant.Trim
            If commandLine = "" Then
                'OK, we're just starting up.
            ElseIf commandLine = "-hide" Then
                'Hide myself from the user. This is to comply with Microsoft's IE anti-trust suit. 
                If MessageBox.Show("To hide access to this program, you need to uninstall it by using Add/Remove Programs in Control Panel.\n\nWould you like to start Add/Remove Programs?", Application.ProductName, MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                    Dim proc As System.Diagnostics.Process = New System.Diagnostics.Process()
                    Dim psi As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo("appwiz.cpl")
                    psi.UseShellExecute = True
                    proc.StartInfo = psi
                    Call proc.Start()
                End If
                Call Application.Exit()
            ElseIf commandLine = "-reinstall" Then
                'Make myself the default web browser:
                Try
                    Dim regKey As Microsoft.Win32.RegistryKey
                    regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\.htm", True)
                    Call regKey.SetValue("", "WebbIE.HTM.4") ' Windows XP?
                    Call regKey.Close()
                    regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\.html", True)
                    Call regKey.SetValue("", "WebbIE.HTM.4") ' Windows XP?
                    Call regKey.Close()
                    regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Clients\StartMenuInternet", True)
                    Call regKey.SetValue("", "WEBBIE4.EXE") ' Default Browser.
                    Call regKey.Close()
                    'Below is something I think only applies to MIME types. See http://msdn.microsoft.com/en-us/library/windows/desktop/cc144154(v=vs.85).aspx
                    'regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Classes\MIME\Database\Content Type\text/html", True)
                    'Call regKey.SetValue("CLSID", "{D664177E-25AE-4B14-ADD0-FFD930C9C210}")
                    'Dim binValue As Byte() = {&H8, &H0, &H0, &H0}
                    'Call regKey.SetValue("Encoding", binValue, Microsoft.Win32.RegistryValueKind.Binary)
                    'Call regKey.SetValue("Extension", ".htm")
                    'Call regKey.Close()
                Catch ex As Exception
                    'Failed to set registry keys
                    MessageBox.Show(modI18N.GetText("WebbIE failed to set itself as the default web browser:") & vbNewLine & ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End Try

            ElseIf commandLine = "-show" Then
                'Show myself to the user, but don't make default. This is to comply with Microsoft's 
                'IE anti-trust suit. I can ignore it because Hide is done by removing the program,
                'so this never makes any sense to call. 
                Call Application.Exit()
            ElseIf commandLine <> "" Then
                'goody, we have an address already
                'strip off any "" first
                commandLine = commandLine.Replace("""", "").Replace("""", "")
                'is this an absolute path (e.g. "D:\page.htm") or a relative one (e.g. "page.htm")
                Try
                    If System.IO.File.Exists(commandLine) Then
                        filePath = commandLine
                    ElseIf System.IO.File.Exists(My.Application.Info.DirectoryPath & "\" & commandLine) Then
                        filePath = My.Application.Info.DirectoryPath & "\" & commandLine
                    ElseIf System.IO.File.Exists(My.Application.Info.DirectoryPath & commandLine) Then
                        filePath = My.Application.Info.DirectoryPath & commandLine
                    End If
                Catch ex As Exception
                    filePath = "" ' Probably an invalid path.
                End Try
                'Does file indicated by commandline exist?
                If Len(filePath) > 0 Then
                    'yes!
                    startedNavigating = True
                    If System.IO.Path.GetExtension(filePath).ToLowerInvariant = ".chm" And filePath.Contains("\") Then
                        'local help file
                        Call Me.Show()
                        frmParseHTMLHelp = New frmParseHTMLHelp
                        helpIndex = frmParseHTMLHelp.ConvertHTMLHelp(filePath)
                        If Len(helpIndex) > 0 Then
                            cboAddress.Text = helpIndex
                            startedNavigating = True
                        End If
                    Else
                        'just an HTML file
                        cboAddress.Text = filePath
                        startedNavigating = True
                    End If
                Else
                    'nope!
                    'web page
                    cboAddress.Text = commandLine 'set the address to the command line argument
                    startedNavigating = True
                End If

            End If

            gDisplayingYoutube = False


            If startedNavigating Then
                'okay, we've started going somewhere from the command prompt
                Call StartNavigating(cboAddress.Text)
            Else
                Call DoHome()
            End If

            'Load bookmarks after loading and setting off browsing, because it takes ages.
            tmrDelayLoadBookmarks.Enabled = True


#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
        'Subscribe to UI Automation events so we can spot when modGlobals.gWebHost.webMain steals focus and bring it back.
        focusHandler = New System.Windows.Automation.AutomationFocusChangedEventHandler(AddressOf OnFocusChanged)
        'This throws a System.ArgumentException in debug mode, but I THINK it's an internal thing you won't see when running.

        Call System.Windows.Automation.Automation.AddAutomationFocusChangedEventHandler(focusHandler)
#End If
        Catch ex As Exception
            Call Debug.Print("Exception in frmMain_Load: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Go to homepage
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DoHome()
        Try
            If My.Settings.UseIEHomepage Then
                cboAddress.Text = gCurrentHomepage
                Call StartNavigating(cboAddress.Text)
            Else
                Call GotoWebbIEHomePage()
            End If
        Catch
        End Try
    End Sub

    Private Sub LoadDataStructures()
        Try
            Dim i As Integer

            For i = 0 To MAX_NUMBER_BUTTON_INPUTS_SUPPORTED - 1
                Call selects(i).Initialize()
            Next i
        Catch
        End Try
    End Sub

    Private Sub frmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'indicate we should stop operations (parsing web pages)
            gExiting = True

#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
        'Unsubscribe from UI Automation events
        If (focusHandler IsNot Nothing) Then
            Call System.Windows.Automation.Automation.RemoveAutomationFocusChangedEventHandler(focusHandler)
        End If
#End If
            'save user settings 
            If Me.WindowState = FormWindowState.Normal Then
                My.Settings.WindowState = FormWindowState.Normal
            ElseIf Me.WindowState = FormWindowState.Maximized Then
                My.Settings.WindowState = FormWindowState.Maximized
            End If
            My.Settings.DisplayFont = Me.txtText.Font
            Call My.Settings.Save()
            'Restore script debugging settings
            Call RestoreScriptErrors()
        Catch
        End Try
    End Sub

    Public Sub SetHomepage()
        Try
            'Update the home page to the current url
            Call SetStartPage(cboAddress.Text)
            'update our global variable: gCurrentHomepage
            gCurrentHomepage = cboAddress.Text
            'update tooltip
            'Tell the user what we've done
            MsgBox(modI18N.GetText("Home page changed. WebbIE will come here when it starts or you select Home from the Navigate menu."), MsgBoxStyle.OkOnly)
        Catch
        End Try
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpAbout.Click
        On Error Resume Next
        Call ShowAbout()
    End Sub

    Public Sub ShowAbout()
        On Error Resume Next
        Call MsgBox(Application.ProductName & vbTab & Application.ProductVersion, MsgBoxStyle.Information, Application.ProductName)
    End Sub

    ''' <summary>
    ''' User selects back option from menu
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Public Sub mnuNavigateBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateBack.Click
        On Error Resume Next
        Call btnBack_Click(btnBack, New System.EventArgs())
    End Sub

    ''' <summary>
    ''' Convert invalid filenames into valid DOS names.
    ''' </summary>
    ''' <param name="inputFilename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetValidFilename(ByRef inputFilename As String) As String
        On Error Resume Next
        For Each c As Char In System.IO.Path.GetInvalidFileNameChars()
            inputFilename = inputFilename.Replace(c, "_")
        Next c
        Return inputFilename.Substring(0, 255)
    End Function


    Public Sub mnuOptionsChangeFont_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuOptionsFont.Click
        'user asks to change the font - show appropriate dialogue
        On Error GoTo errorHandler
        Me.fdMain.ShowApply = False
        Me.fdMain.AllowSimulations = False
        Me.fdMain.AllowVerticalFonts = False
        Me.fdMain.ShowEffects = False
        Me.fdMain.ShowHelp = False
        Me.fdMain.Font = CType(Me.txtText.Font.Clone(), System.Drawing.Font)
        If Me.fdMain.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Me.txtText.Font = fdMain.Font
        End If
        Exit Sub
errorHandler:
        Exit Sub
    End Sub

    Public Sub mnuEditCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditCopy.Click
        'process user menu command copy
        Try
            If Not Me.ActiveControl Is Nothing Then
                Select Case Me.ActiveControl.Name
                    Case "txtText"
                        If txtText.SelectedText <> "" Then ' only copy if we have something selected
                            Call My.Computer.Clipboard.Clear()
                            Call My.Computer.Clipboard.SetText(txtText.SelectedText, System.Windows.Forms.TextDataFormat.UnicodeText)
                        End If
                    Case "cboAddress"
                        If cboAddress.SelectedText <> "" Then ' only copy if we have something selected
                            Call My.Computer.Clipboard.Clear()
                            Call My.Computer.Clipboard.SetText(cboAddress.SelectedText)
                        End If
                    Case Else
                        Dim wb As SHDocVw.WebBrowser
                        wb = CType(modGlobals.gWebHost.webMain.ActiveXInstance, SHDocVw.WebBrowser)
                        Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_COPY, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT)
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub mnuEditPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditPaste.Click
        On Error Resume Next
        Call Paste()
    End Sub

    Public Sub mnuEditSelectall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditSelectall.Click
        Try
            'select everything on the current object.
            Dim currentObjectName As String

            If Me.ActiveControl Is Nothing Then
                currentObjectName = modGlobals.gWebHost.webMain.Name
            Else
                currentObjectName = Me.ActiveControl.Name
            End If
            If currentObjectName = modGlobals.gWebHost.webMain.Name Then
                Call modSendKeys.SendSelectAll()
            ElseIf currentObjectName = cboAddress.Name Then
                Call cboAddress.SelectAll()
            ElseIf currentObjectName = txtText.Name Then
                Call txtText.SelectAll()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Find the next instance of a word searched for by the user and select it. Display a message if not found. If no word to search for, call FindText()
    ''' </summary>
    Private Sub FindNext()
        Try
            '

            Dim where As Integer
            Dim originalPos As Integer
            Dim searchIn As String

            'If mIEVisible Then
            '    Dim tr As mshtml.IHTMLTxtRange
            '    Dim doc As mshtml.IHTMLDocument2
            '    doc = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.IHTMLDocument2)
            '    tr = CType(doc.selection.createRange, mshtml.IHTMLTxtRange)
            '    If tr Is Nothing Then
            '    ElseIf tr.text = "" Then
            '        'Start a new Find
            '        Dim wb As SHDocVw.WebBrowser
            '        wb = CType(modGlobals.gWebHost.webMain.ActiveXInstance, SHDocVw.WebBrowser)
            '        Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_FIND, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT, CObj(gfindText))
            '    Else
            '        'Find the next instance of the selected text.
            '        gfindText = tr.text
            '        Call tr.collapse(False)
            '        If tr.findText(gfindText) Then
            '            Call tr.select()
            '            Call tr.scrollIntoView()
            '        Else
            '            Call Beep()
            '        End If
            '    End If

            'Else
            If gfindText <> "" Then
                originalPos = txtText.SelectionStart 'go to start of the found word
                'make lower case and search onwards from the existing word
                searchIn = Mid(txtText.Text, originalPos + 2, Len(txtText.Text) - txtText.SelectionStart - 2)
                where = InStr(1, searchIn, gfindText, CompareMethod.Text)
                If (where > 0) Then 'if found
                    where += originalPos
                    'startLine = GetCurrentLineIndex(txtText)
                    txtText.SelectionStart = where
                    txtText.SelectionLength = Len(gfindText) 'highlight the word
                    'endLine = GetCurrentLineIndex(txtText)
                    'Call Scroll(txtText, endLine - startLine)
                    'Call ScrollToCursor(txtText)
                Else
                    'if unfound, display a warning
                    MsgBox(modI18N.GetText("No further occurrences found"), MsgBoxStyle.Information, Application.ProductName)
                End If
            Else 'if there is no text to search
                Call FindText()
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuEditFindnext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFindnext.Click
        On Error Resume Next
        Call FindNext()
    End Sub

    Private Sub FindText()
        Try
            'Find a word looked for in the main text box

            Dim where As Integer

            'in text view
            gfindText = InputBox(modI18N.GetText("Find what:"), modI18N.GetText("Find"), gfindText)
            'if OK is clicked
            If gfindText <> "" Then
                ' Find string in text
                where = InStr(txtText.SelectionStart + 2, txtText.Text, gfindText, CompareMethod.Text)
                If where > 0 Then ' If found..
                    Call txtText.Focus()
                    'startLine = GetCurrentLineIndex(txtText)

                    txtText.SelectionStart = where - 1 ' set selection start and
                    txtText.SelectionLength = Len(gfindText) ' set selection length
                    Call txtText.ScrollToCaret()
                Else
                    'try from start
                    where = InStr(1, txtText.Text, gfindText, CompareMethod.Text)
                    If where > 0 Then
                        'found it: is it where we started?
                        If where = txtText.SelectionStart + 1 Then
                            'whoops, already there
                            'Debug.Print "Already there"
                            MsgBox(modI18N.GetText("No more") & " " & gfindText, MsgBoxStyle.Information, modI18N.GetText("Word not found"))
                        Else
                            'nope, go there
                            Call txtText.Focus()
                            txtText.SelectionStart = where - 1
                            txtText.SelectionLength = Len(gfindText)
                        End If
                    Else
                        'not found at all
                        MsgBox(modI18N.GetText("Cannot find") & " " & gfindText, MsgBoxStyle.Information, modI18N.GetText("Word not found"))
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Public Sub mnuEditFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFindtext.Click
        On Error Resume Next
        Call FindText()
    End Sub

    Public Sub mnuNavigateForward_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateForward.Click
        On Error Resume Next
        Call DoForward()
    End Sub

    Public Sub mnuNavigateHome_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateHome.Click
        On Error Resume Next
        btnHome_Click(btnHome, New System.EventArgs())
    End Sub

    Public Sub mnuLinksNextLink_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksNextlink.Click
        Try
            Dim lineNumber As Integer
            Dim linkFound As Boolean

            If Me.txtText.Text <> "" Then
                lineNumber = GetCurrentLineIndex(txtText) + 1
                linkFound = False
                While lineNumber < GetNumberOfLines(txtText) And Not linkFound
                    If txtText.Lines(lineNumber).Contains(ID_LINK) Then
                        'found a link
                        linkFound = True
                    Else
                        lineNumber += 1
                    End If
                End While
                If linkFound Then
                    txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                    'Don't set length of link.
                    'txtText.SelLength = Len(Trim(GetCurrentLine(txtText)))
                    txtText.SelectionLength = 0
                    'Call ScrollToCursor(txtText)
                Else
                    Call PlayErrorSound()
                End If
            End If
        Catch
        End Try
    End Sub

#Region "Printing"
    'http://msdn.microsoft.com/en-us/library/system.drawing.printing.printdocument.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-2

    Private ReadOnly _printFont As Font
    Private ReadOnly _streamToPrint As System.IO.StreamReader

    Public Sub mnuFilePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFilePrint.Click
        'print the webpage
        Try
            'So this code prints directly from the WebBrowser control on frmWeb.
            Dim wb As SHDocVw.WebBrowser
            wb = CType(gWebHost.webMain.ActiveXInstance, SHDocVw.WebBrowser)
            Call wb.ExecWB(SHDocVw.OLECMDID.OLECMDID_PRINT, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT)
        Catch ex As Exception
            'Ah, this isn't written right: won't do anything.
            'Call System.Diagnostics.EventLog.WriteEntry("mnuFilePrint_Click", "Failed to print using OLECMDID_PRINT: " + ex.Message)
        End Try

        'Whereas THIS code prints the actual text in the text area. I've had a request to print the actual
        'web page, which of course is useful for tickets and bank statements and the like, so this makes sense.
        'If a user wants to print the whole page of actual text they can cut/paste it into Word or similar. They
        'can't get at the web page.
        'Dim pd As New System.Drawing.Printing.PrintDocument()
        'Dim printDialog As Windows.Forms.PrintDialog = New Windows.Forms.PrintDialog()
        'printDialog.Document = pd
        'If (printDialog.ShowDialog() = DialogResult.OK) Then

        '    Dim path As String = System.IO.Path.GetTempFileName
        '    Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(path, False, System.Text.Encoding.UTF8)
        '    Call sw.Write(txtText.Text)
        '    Call sw.Close()
        '    Call Application.DoEvents()
        '    Try
        '        _streamToPrint = New System.IO.StreamReader(path, System.Text.Encoding.UTF8)
        '        Try
        '            _printFont = CType(txtText.Font.Clone, Drawing.Font) 'New Font("Arial", 10)
        '            AddHandler pd.PrintPage, AddressOf Me.pd_PrintPage
        '            pd.Print()
        '        Finally
        '            _streamToPrint.Close()
        '        End Try
        '    Catch ex As Exception
        '        MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End If
    End Sub

    ' The PrintPage event is raised for each page to be printed. 
    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As System.Drawing.Printing.PrintPageEventArgs)
        Try
            Dim linesPerPage As Single = 0
            Dim yPos As Single = 0
            Dim count As Integer = 0
            Dim leftMargin As Single = ev.MarginBounds.Left
            Dim topMargin As Single = ev.MarginBounds.Top
            Dim line As String = Nothing

            ' Calculate the number of lines per page.
            linesPerPage = ev.MarginBounds.Height / _printFont.GetHeight(ev.Graphics)

            ' Print each line of the file. 
            While count < linesPerPage
                line = _streamToPrint.ReadLine()
                If line Is Nothing Then
                    Exit While
                End If
                yPos = topMargin + count * _printFont.GetHeight(ev.Graphics)
                ev.Graphics.DrawString(line, _printFont, Brushes.Black, leftMargin, yPos, New StringFormat())
                count += 1
            End While

            ' If more lines exist, print another page. 
            If (line IsNot Nothing) Then
                ev.HasMorePages = True
            Else
                ev.HasMorePages = False
            End If
        Catch
        End Try
    End Sub

#End Region

    Public Sub mnuNavigateRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateRefresh.Click
        On Error Resume Next
        If NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) < 0 And NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) < 0 Then
            'User pressing Shift and Control, do IE refresh.
            Call frmWeb.webMain.Refresh(WebBrowserRefreshOption.Completely)
        Else
            'In the context of WebbIE, just refresh my rendering!
            Call RefreshCurrentPage()
        End If
    End Sub

    Public Sub mnuNavigateStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNavigateStop.Click
        On Error Resume Next
        Call btnStop_Click(btnStop, New System.EventArgs())
    End Sub

    ''' <summary>
    ''' Prepare the links from the current page and show the lstLinks form to let
    ''' the user select one.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ListLinks()
        Try
            Dim i As Integer
            Dim tempLinks As linkStruct
            Dim swapping As Boolean 'whether bubble sorting is complete

            'Sort links into alphabetical order
            For i = 0 To gNumLinks - 1
                gSortedLink(i) = gLinks(i)
                System.Diagnostics.Debug.Print("LINK " & i & " " & gLinks(i).description)
            Next i

            'initialise swapping to ensure loop entry
            swapping = True

            'start bubble (up) sort
            While swapping
                'ensures the until loop breaks when swapping finishes
                swapping = False
                'count through the links
                For i = 0 To (gNumLinks - 2)
                    'if swapping is needed, perform the swap
                    If (StrComp(gSortedLink(i).description, gSortedLink(i + 1).description, CompareMethod.Text) > 0) Then
                        tempLinks = gSortedLink(i) 'save the later item
                        gSortedLink(i) = gSortedLink(i + 1) 'overwrite the later item with the earlier one
                        gSortedLink(i + 1) = tempLinks 'make the later item, the earlier one
                        'indicate that swapping still occurs
                        swapping = True
                    End If
                Next i
            End While
            Call frmLinks.PopulateList()
            frmLinks.cmdGo.Enabled = False
            Call frmLinks.ShowDialog(Me)
        Catch
        End Try
    End Sub

    Public Sub mnuLinksViewLinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLinksViewlinks.Click
        On Error Resume Next
        Call ListLinks()
    End Sub

    Private Sub mnuEditWebsearch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditWebsearch.Click
        On Error Resume Next
        Call DoWebSearch()
    End Sub

    ''' <summary>
    ''' Pop up a box to prompt user for a search phrase for google
    ''' </summary>
    Public Sub DoWebSearch()
        On Error Resume Next
        If Not frmGoogle.Visible Then
            Call frmGoogle.ShowDialog(Me)
        End If
    End Sub

    Private Sub StopBusyAnimation()
        Try
            tmrBusyAnimation.Enabled = False
            If My.Settings.ToolbarCaptions Then
                picBusy.Image = My.Resources.timer_done_big
            Else
                picBusy.Image = My.Resources.timer_done
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrBusyAnimation_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrBusyAnimation.Tick
        If modGlobals.gClosing Then
            Me.tmrBusyAnimation.Enabled = False
            Exit Sub
        End If
        'Not busy? Don't play. This is because we get a whole lot of clicking when we shouldn't. Fails to stop this timer
        'probably. 
        'But then we don't get lots of ticking when we should! 

        Try
            'every 0.3 second, move the animation. Play working sound 
            Static counter As Integer 'static retains value on each call

            'Debug.Print("tmrBusy! " & New System.Random().Next())
            Select Case counter
                Case 0
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer1_big
                    Else
                        picBusy.Image = My.Resources.timer1
                    End If
                    counter = 1
                    'Call Debug.Print("Working sound in image graphic")
                    Call PlayWorkingSound()
                Case 1
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer2_big
                    Else
                        picBusy.Image = My.Resources.timer2
                    End If
                    counter = 2
                Case 2
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer3_big
                    Else
                        picBusy.Image = My.Resources.timer3
                    End If
                    counter = 3
                Case 3
                    If My.Settings.ToolbarCaptions Then
                        picBusy.Image = My.Resources.timer4_big
                    Else
                        picBusy.Image = My.Resources.timer4
                    End If
                    counter = 0
            End Select
        Catch
        End Try
    End Sub

    Private Sub ClearPageData()
        On Error Resume Next
        'Clears the page arrays and other per-page data.
        Call modAccesskeys.ClearAccessKeys() ' clear any access keys
        numForms = 0
        numFileInputs = 0
        numTextInputs = 0
        numEmailInputs = 0
        numRangeInputs = 0
        numSearchInputs = 0
        gNumLinks = 0
        numTargets = 0
        numTables = 0
        numSelects = 0
        numImages = 0
        numPassInputs = 0
        numSubmitInputs = 0
        numResetInputs = 0
        numCheckboxInputs = 0
        numRadioInputs = 0
        numVideos = 0
        numButtonInputs = 0
        numTextAreaInputs = 0
        gPageHasAnArticle = False
        mInArticle = 0
        Call modFormLabelHandler.Clear()
        gImageHrefs = New Collection
        gRSSFeedURL = ""
        gPageHasMain = False
    End Sub

    Private Sub DisplayOutput()
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            'displays the contents of mOutput in the txtText text box but strips out any blank lines along
            'the way

            'Variables for Method 1 of removing blank lines
            Dim splitOutput() As String
            Dim i As Integer
            Dim numberSplitStrings As Integer
            Dim writtenBlank As Boolean
            Dim gotLine As String
            Dim nextLine As String
            Dim newPageContent As System.Text.StringBuilder = New System.Text.StringBuilder(mOutput.ToString.Length)

            If modGlobals.gWebHost.webMain.DocumentTitle = "" Then
                'No use telling the user there is no title...
            Else
                Call newPageContent.Append(modI18N.GetText("WEBPAGE:") & " " & modGlobals.gWebHost.webMain.DocumentTitle)
            End If
            Call newPageContent.Append(vbNewLine & vbNewLine)
            'Removing blank lines in the text output: three problems, (1) should we do it?
            '(2) how thoroughly do we do it (3) how do we make it efficient for big files?
            'One option is to check the size of the html file and if too big, simply
            'skip this process. Another might be to parse a big chunk and display that
            'while getting on with the rest - but appending would be messy.

            'Method 1
            'This is the previous method, which is slower for lots of lines but much
            'more thorough than the above method.

            'Split the output text to an array.
            splitOutput = Split(mOutput.ToString, vbNewLine)
            'Work out how many strings we have in the array.
            numberSplitStrings = -1
            numberSplitStrings = UBound(splitOutput)
            'Check we have the right amount of text to display. 
            If numberSplitStrings < 4000 And numberSplitStrings > -1 Then
                'First line is the next line. 
                nextLine = Trim(splitOutput(0))
                'Iterate through all lines in the text array.
                For i = 0 To numberSplitStrings - 1
                    'Get next two lines, gotLine and nextLine.
                    gotLine = nextLine
                    nextLine = splitOutput(i + 1).Trim
                    'Don't add an empty line.
                    If gotLine.Length = 0 Then
                        'Nope, don't add this.
                    Else
                        'OK, we have some text. Is it a line that should be blank? This should be BLANK_LINE_MARKER.
                        If gotLine = BLANK_LINE_MARKER Then
                            'Insert a blank line, unless we have already added a blank line. So we should only 
                            'have one blank line at a time. 
                            If Not writtenBlank Then
                                'Need to add a blank line. 
                                Call newPageContent.AppendLine()
                                writtenBlank = True
                            End If
                        Else
                            If i < numberSplitStrings Then
                                If gotLine.Length < 3 Then
                                    If nextLine.Contains(ID_LINK) Then
                                        'we have a line of text where we have a line immediately following but
                                        'the current line is only two characters: typically this is where we
                                        'have something like [edit] on Wikipedia. This will let us have
                                        'LINK 1: [edit]
                                        'rather than
                                        '[
                                        'LINK1: edit]
                                        nextLine = Microsoft.VisualBasic.Left(nextLine, InStr(1, nextLine, ":") + 1) & gotLine & " " & Microsoft.VisualBasic.Right(nextLine, Len(nextLine) - InStr(1, nextLine, ":") - 1)
                                        gotLine = ""
                                    End If
                                End If
                            End If
                            'Remove lines that just contain a section marker and nothing else - noise
                            'Reworked so I don't disappear section markers in links, so 
                            'Heading 2: 
                            'Link: Cats 
                            'becomes
                            'Link: Cats
                            'This approach should fix it: 
                            If gotLine.StartsWith(SECTION_MARKER_H1) And gotLine.Length <= SECTION_MARKER_H1.Length + 1 Then
                                '    gotLine = ""
                                If nextLine.StartsWith(ID_LINK) Then
                                    'Ah, rearrange.
                                    nextLine = ID_LINK & ": " & SECTION_MARKER_H1 & " " & nextLine.Replace(ID_LINK, "")
                                End If
                                gotLine = ""
                            ElseIf gotLine.StartsWith(SECTION_MARKER_H2) And gotLine.Length <= SECTION_MARKER_H2.Length + 3 Then
                                If nextLine.StartsWith(ID_LINK) Then
                                    'Ah, rearrange.
                                    nextLine = ID_LINK & ": " & SECTION_MARKER_H2 & " " & nextLine.Replace(ID_LINK, "")
                                End If
                                gotLine = ""
                            ElseIf gotLine.StartsWith(SECTION_MARKER_H3) And gotLine.Length <= SECTION_MARKER_H3.Length + 3 Then
                                If nextLine.StartsWith(ID_LINK) Then
                                    'Ah, rearrange.
                                    nextLine = ID_LINK & ": " & SECTION_MARKER_H3 & " " & nextLine.Replace(ID_LINK, "")
                                End If
                                gotLine = ""
                            ElseIf gotLine.StartsWith(SECTION_MARKER_H4) And gotLine.Length <= SECTION_MARKER_H4.Length + 3 Then
                                If nextLine.StartsWith(ID_LINK) Then
                                    'Ah, rearrange.
                                    nextLine = ID_LINK & ": " & SECTION_MARKER_H4 & " " & nextLine.Replace(ID_LINK, "")
                                End If

                                gotLine = ""
                            ElseIf gotLine.StartsWith(SECTION_MARKER_H5) And gotLine.Length <= SECTION_MARKER_H5.Length + 3 Then
                                If nextLine.StartsWith(ID_LINK) Then
                                    'Ah, rearrange.
                                    nextLine = ID_LINK & ": " & SECTION_MARKER_H5 & " " & nextLine.Replace(ID_LINK, "")
                                End If

                                gotLine = ""
                            ElseIf gotLine.StartsWith(SECTION_MARKER_H6) And gotLine.Length <= SECTION_MARKER_H6.Length + 3 Then
                                If nextLine.StartsWith(ID_LINK) Then
                                    'Ah, rearrange.
                                    nextLine = ID_LINK & ": " & SECTION_MARKER_H6 & " " & nextLine.Replace(ID_LINK, "")
                                End If
                                gotLine = ""
                            End If
                            'Tidy up punctuation
                            gotLine = gotLine.Replace("  ,", ",").Replace("[", " ").Replace("]", "").Replace(" ,", ",").Replace(" .", ".").Replace(" :", ":")
                        If Len(gotLine) > 0 Then
                                    Call newPageContent.AppendLine(gotLine)
                                    writtenBlank = False
                                End If
                            End If
                        End If
                    Call Application.DoEvents()
                Next i
                'Me.txtText.Clear()
                'Me.txtText.SelectionIndent = 20 Use this to apply a left margin!
                Call SetText(newPageContent.ToString())
            Else
                Call newPageContent.Append(Join(splitOutput, vbNewLine))
                'Lines containing only BLANK_LINE_MARKER should indeed be blank.
                Call SetText(newPageContent.ToString.Replace(BLANK_LINE_MARKER, ""))
            End If
        Catch
        End Try
    End Sub

    Private Function ExtractHREFLink(ByRef href As String) As String
        Try
            'Could probably do this with these two lines... TODO.
            'Dim newUrl As System.Uri = New System.Uri(href)
            'Return newUrl.AbsolutePath

            'strips out the functional parts of a URL (href) to the file or domain name - so it takes
            'off ports, get information, protocol information etc.

            Dim link As String ' the href of the link
            Dim postProtocolStart As Integer ' the position of the first character after the
            ' "whateverprototol://" bit
            Dim slashPosition As Integer ' the position of the last filename delimiter found
            Dim newSlashPosition As Integer ' the next one found
            Dim firstQuestionMark As Integer ' the position of a question mark,
            ' indicating the arguments passed to a page.  These will excised
            Dim linkLength As Integer ' the length of the bit we want to use

            link = Trim(href)
            firstQuestionMark = InStr(1, link, "?") ' the start of a query string
            If firstQuestionMark = 0 Then firstQuestionMark = Len(link) + 1
            'now excise any query string
            link = Mid(link, 1, firstQuestionMark - 1)

            postProtocolStart = InStr(1, link, "//") ' to start of ' if any

            If postProtocolStart > 0 Then
                'it's an absolute link - e.g. "http://site.com"
                'take off ending / if any
                If link.EndsWith("/") Then link = link.Substring(0, link.Length - 1)
                postProtocolStart += 2
                slashPosition = postProtocolStart
                'ASSERTION: slashPosition is the first character after the protocol
                'now find the possible filename
                newSlashPosition = InStr(slashPosition, link, "/")
                While newSlashPosition > 0
                    'ASSERTION: we've found a(nother) "/"
                    slashPosition = newSlashPosition + 1
                    newSlashPosition = InStr(slashPosition, link, "/")
                End While
                'ASSERTION: there is no (other) "/" character after slashPosition
                'if the position has moved on from the "//" then we've got a
                'local file, _unless_
                'it is also the end of the string.
                If slashPosition > postProtocolStart And slashPosition < Len(link) Then
                    'hurrah, we have a local file: it'll run from slashPosition to the end
                    linkLength = Len(link) - slashPosition + 1 ' plus one includes the character marked by
                    'slashPosition - I hate not counting from zero... ruddy Microsoft.VisualBasic...
                    link = Mid(link, slashPosition, linkLength)
                Else
                    'nope, we don't have a local file: the slash must be the end of the link,
                    'OR we haven't moved at all from the ' bit
                    'e.g. "www.yahoo.com/" or "http://www.yahoo.com", so just use the host: get this by
                    'going from the end of the "(ht)tp://" bit to the character
                    'before the end, which is a "/".
                    If Len(link) = 0 Then
                        link = modI18N.GetText("Unidentified link")
                    Else
                        link = Mid(link, InStr(1, link, "tp://") + 5, Len(link) - (InStr(1, link, "tp://") + 4))
                    End If
                End If
            Else
                'it's a relative link - e.g. "nextpage.htm" - and might have an argument
                'try just using the link itself, less any argument
                'get rid of any directories
                slashPosition = 1
                newSlashPosition = InStr(slashPosition, link, "/")
                While newSlashPosition > 0
                    'ASSERTION: we've found a(nother) "/"
                    slashPosition = newSlashPosition + 1
                    newSlashPosition = InStr(slashPosition, link, "/")
                End While
                linkLength = Len(link) - slashPosition + 1 ' plus one includes the character marked by
                'slashPosition - I hate not counting from zero... ruddy Microsoft.VisualBasic...
                link = Mid(link, slashPosition, linkLength)
            End If
            link = Microsoft.VisualBasic.Left(link, 64)
            'now strip off filename extensions for known file types
            If LCase(Microsoft.VisualBasic.Right(link, 4)) = ".htm" Then link = Microsoft.VisualBasic.Left(link, Len(link) - 4)
            If LCase(Microsoft.VisualBasic.Right(link, 5)) = ".html" Then link = Microsoft.VisualBasic.Left(link, Len(link) - 5)
            If LCase(Microsoft.VisualBasic.Right(link, 4)) = ".php" Then link = Microsoft.VisualBasic.Left(link, Len(link) - 4)
            If LCase(Microsoft.VisualBasic.Right(link, 4)) = ".asp" Then link = Microsoft.VisualBasic.Left(link, Len(link) - 4)
            If LCase(Microsoft.VisualBasic.Right(link, 5)) = ".shtm" Then link = Microsoft.VisualBasic.Left(link, Len(link) - 5)
            If LCase(Microsoft.VisualBasic.Right(link, 6)) = ".shtml" Then link = Microsoft.VisualBasic.Left(link, Len(link) - 6)
            ExtractHREFLink = link
        Catch
            Return href
        End Try
    End Function

    ''' <summary>
    ''' Process the loading HTML object model and display it as text. Call this after refresh, page navigation, or any time
    ''' you need to update this view to reflect a change to the DOM.
    ''' </summary>
    Private Sub ParseDocument()
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            'process through the website to create the text output
            If mFormClosing Then Return

            'clear data
            Call ClearPageData()

            Dim got As String = ""
            Dim doc As mshtml.HTMLDocument

            'tell user we've progressed from downloading to processing
            staMain.Items.Item(0).Text = modI18N.GetText("Examining")
            'display update in browser text box
            'Don't change the page until we've done loading and displaying.
            mOutput = New System.Text.StringBuilder(32000)
            'Reset the terminate flag
            mTerminateParsing = False
            'reset the cropped flag
            mCropped = False
            'check for an error page
            If mblnErrorPage Then
                'oh no, error page!
                mblnErrorPage = False
                Call mOutput.Append(modGlobals.gWebHost.webMain.Document.Body.InnerText)
                gRSSFeedURL = ""
            Else
                'parse the document into txtText
                mJustDidLink = False
                mNewline = ""
                'TODO handle non-HTML document types by checking document.mimeType.
                'process HTML if that is what we have
                'check to see if this document contains (or should contain)
                'the target id found, if any
                'TODO Handle frames for headings.
                If modGlobals.gWebHost.webMain.Document Is Nothing Then
                Else
                    doc = CType(modGlobals.gWebHost.webMain.Document.DomDocument, mshtml.HTMLDocument)
                    Dim bodyNode As mshtml.IHTMLDOMNode = CType(doc.body, mshtml.IHTMLDOMNode)
                    gRSSFeedURL = modRSS.CheckForRSSAlternate(CType(doc, mshtml.IHTMLDocument3))
                    If gRSSFeedURL <> "" Then
                        'Add link to news feed at top of page.
                        gNumLinks += 1
                        gLinks(gNumLinks - 1).element = Nothing
                        gLinks(gNumLinks - 1).description = "RSS"
                        gLinks(gNumLinks - 1).address = "FAKEA-RSS"
                        Call mOutput.Append(ID_LINK & IIf(My.Settings.NumberLinks, " " & gNumLinks, "").ToString & ": ")
                        Call mOutput.Append(modI18N.GetText("RSS News Feed") & vbNewLine)
                    End If
                    'TODO Handle frames for labels.
                    Call modFormLabelHandler.ProcessWebpageForLabels(doc)
                    Call IdentifyHeadings(CType(bodyNode, mshtml.IHTMLElement))
                    Call ParseNode(bodyNode)
                    Call mOutput.Append(Environment.NewLine)
                End If
            End If
            btnRSS.Enabled = (gRSSFeedURL <> "")
            'Replace form count with final tally
            If gShowFormsOnly Then
                Dim temp As String = mOutput.ToString.Replace("LKJLKJLKJLKJLKJLKJLKJ", CStr(numForms))
                mOutput = New System.Text.StringBuilder(temp)
            End If
            Call DisplayOutput()
            btnStop.Enabled = False
            btnRefresh.Enabled = True
            staMain.Items.Item(0).Text = modI18N.GetText("Done") 'once parsing is complete, change status to idle
            mnuLinksViewlinks.Enabled = True 'allow the links to be viewed in a seperate form
            cboAddress.SelectionLength = 0 ' don't select any of the text in the address bar
            cboAddress.SelectionStart = 1 ' put the cursor at the beginning
            Call cboAddress.Refresh() ' redraw it

            'If mIEVisible Then
            '    'don't do any focus changing if we're looking at the IE box
            'Else
            Call txtText.Focus()
            Call Application.DoEvents()
            txtText.SelectionLength = 0
            'End If
            Exit Sub
            'AccessDenied:
            'There appears no way (for obvious security reasons) to get anything out of o, so we would have to flag
            'this as a problem then use MSAA to navigate to the offending frame in the browser object. Ugh. Let's leave
            'that for the moment as unsupported.
            'Debug.Print Err.Number & " Error in ParseDocument()" ' if 70 then Access Denied.
            'Resume ResumeAccessDenied
        Catch
        Finally
            Call PlayDoneSound()
            Call StopBusyAnimation()
        End Try
    End Sub

    Private Sub ParseNode(ByRef node As mshtml.IHTMLDOMNode, Optional ByRef orderedListPrefix As String = "")
        Try
            'processes a document html node and all of its children and siblings (by calling itself.
            'Outputs results to output, which it returns
            ' mNewline is either Empty or vbNewline, and should be set accordingly

            Dim tagname As String ' the HTML node name
            Dim label As String
            'Dim htmlE As IHTMLElement ' the node as an IHTMLElement object, lets us get more stuff
            Dim linkContent As String = "" ' the content of an a node to display
            Dim nodeIterator As mshtml.IHTMLDOMNode ' used to iterate through any further node collections
            Dim nodeIterator2 As mshtml.IHTMLDOMNode ' used to iterate again
            Dim i As Integer ' counter
            Dim trimmedText As String ' used to strip spare spaces out of text
            Dim labelText As String = ""
            Dim notLabelText As String = ""
            Dim controlToBeLabelled As mshtml.IHTMLElement
            Dim href As String = ""
            Dim element As mshtml.IHTMLElement
            Dim element2 As mshtml.IHTMLElement2
            Dim element3 As mshtml.IHTMLElement3
            Dim ni As mshtml.IHTMLDOMChildrenCollection
            Dim isHidden As Boolean ' HTML5 hidden attribute.

            Dim validityCount As Integer

            Dim hasDisplayNone As Boolean
            Dim hasOnClick As Boolean
            Dim hasId As Boolean
            Dim hasName As Boolean
            Dim hasAccessKey As Boolean
            Dim isContentEditable As Boolean

            If mTerminateParsing Then
                'Stop parsing immediately. This is usually because we have some kind of error condition.
                Exit Sub
            ElseIf modGlobals.gClosing Then
                Exit Sub
            ElseIf node Is Nothing Then
                'No idea why we might get this, but I've observed it on the Economist website.
                Exit Sub
            ElseIf node.nodeType = TEXT_NODE Then
                'Text node: has no children by definition, so can just get content and carry on.
                'get rid of whitespace at end/start of line, INCLUDING non-breaking whitespace (Unicode
                '   value 160, NBSP)
                '3.10.5 Try RTrim instead of Trim.
                If node.nodeValue Is Nothing Then
                    trimmedText = "" ' Experienced on http://www.guardian.co.uk - text node with no nodeValue.
                Else
                    trimmedText = node.nodeValue.ToString.Replace(Chr(160), " ").Replace(vbTab, " ").Replace(vbCr, vbNewLine).Trim
                End If
                If trimmedText <> "" Then
                    If Len(trimmedText) > NUMBER_CHARS_AFTER_LINK_PERMITTED And mJustDidLink Then
                        Call mOutput.Append(vbNewLine & trimmedText & " ")
                    Else
                        Call mOutput.Append(trimmedText & " ")
                    End If
                    mJustDidHeading = False ' did the heading text fine.
                    mJustDidLink = False
                    mNewline = vbNewLine
                End If
            ElseIf node.nodeType <> ELEMENT_NODE Then
                'No need to process if not TEXT_NODE (above) or ELEMENT_NODE (below).
            Else
                Call Application.DoEvents()
                ' TODO add role="whatever" from HTML5.
                tagname = node.nodeName.ToUpperInvariant()
                element = CType(node, mshtml.IHTMLElement)
                If element.onclick Is Nothing Then
                    hasOnClick = False
                Else
                    hasOnClick = (element.onclick.ToString <> "")
                End If
                If element.id Is Nothing Then
                    hasId = False
                Else
                    hasId = (element.id <> "")
                End If
                Try
                    Dim o As Object = element.getAttribute("name")
                    If IsDBNull(o) Or IsNothing(o) Then
                        hasName = False
                    Else
                        hasName = o.ToString.Length > 0
                    End If
                Catch unAuthEx As UnauthorizedAccessException
                    hasName = False
                End Try
                element2 = CType(element, mshtml.IHTMLElement2)
                If element2.accessKey Is Nothing Then
                    hasAccessKey = False
                Else
                    hasAccessKey = (element2.accessKey.ToString <> "")
                End If
                isHidden = False
                If Not IsDBNull(element.getAttribute("hidden")) And tagname <> "BODY" Then
                    'Attribute is present, but ignore for BODY element. I have no idea why it would be set for
                    'the body element, but it is for Google and Facebook (Jan 2015)
                    isHidden = True
                End If
                If Not IsDBNull(element.getAttribute("aria-hidden")) And tagname <> "BODY" Then
                    'Attribute is present.
                    isHidden = CBool(element.getAttribute("aria-hidden"))
                End If
                If element2.currentStyle Is Nothing Then
                    hasDisplayNone = False
                ElseIf element2.currentStyle.display Is Nothing Then
                    hasDisplayNone = False
                Else
                    hasDisplayNone = (element2.currentStyle.display = "none")
                End If
                element3 = CType(element, mshtml.IHTMLElement3)
                isContentEditable = (element3.contentEditable = "true")
                'Check for onclick
                If hasOnClick Then
                    If tagname = "A" Or tagname = "INPUT" Or tagname = "APPLET" Or tagname = "BUTTON" Or tagname = "LABEL" Or tagname = "MAP" Or tagname = "OPTION" Or tagname = "OPTGROUP" Or tagname = "SELECT" Then
                        'already handled by normal node processing
                        'Debug.Print "A with onclick: " & node.getAttribute("onclick")
                        hasOnClick = False
                    Else
                        'Need to add link: hey, let's do this by faking an A node!
                        'But NOT for body nodes, because that would be bad!
                        'TODO However, this is a problem when an element with child nodes also
                        'has an onclick attribute. This can be seen on Facebook where UL
                        'elements with onclick contain A elements. So we can't just make
                        'something into an A tagName (what we did before, that is)
                        'So hasOnClick means you need to make it a click/link.
                    End If
                End If
                'attempt to speed things up
                If mSeekingInternalTarget Then
                    If hasId Then
                        If element.id = mInternalTarget Then
                            Call mOutput.Append(TARGET_MARKER)
                        End If
                    End If
                    If hasName Then
                        'If we don't have a matching id check for a matching name.
                        If element.getAttribute("name").ToString = mInternalTarget Then
                            Call mOutput.Append(TARGET_MARKER)
                        End If
                    End If
                End If
                If mSeekingFocusElement Then
                    If node Is mElementWithFocus Then
                        Call mOutput.Append(AJAX_TARGET_MARKER)
                        mSeekingFocusElement = False
                    End If
                End If
                If tagname = "H1" Or tagname = "H2" Or tagname = "H3" Or tagname = "H4" Or tagname = "H5" Or tagname = "H6" Then
                    'Make sure there is a blank line before Headings by adding a blank line marker.
                    Call mOutput.AppendLine(vbNewLine & BLANK_LINE_MARKER)
                    mNewline = ""
                End If
                If tagname = "FORM" And gShowFormsOnly Then
                    mInForm = True
                    numForms += 1
                    Call mOutput.AppendLine(mNewline & BLANK_LINE_MARKER & vbNewLine & GetText("Form") & ": " & numForms)
                    mNewline = ""
                    Call ParseChildren(node, , orderedListPrefix)
                    Call mOutput.AppendLine(mNewline & BLANK_LINE_MARKER & vbNewLine)
                    mNewline = ""
                    mInForm = False
                ElseIf gShowFormsOnly And Not mInForm Then
                    'Do no output, but process children
                    Call ParseChildren(node, , orderedListPrefix)
                ElseIf (hasDisplayNone Or isHidden) And tagname <> "AUDIO" Then  'check displayed
                    '3.7.4.
                    'No longer hide Display:None sections, because that's not compliant with the W3C standards.
                    'don't display this node at all! And don't do kids!
                    '3.8.0
                    'No, still hide Display:None sections, because people - like the BBC and Facebook - use it to hide error messages,
                    'and I don't want to show these to users. Damn the W3C.
                    'No, also, .display="none" does not mean "don't appear in the output" - it means "don't render" - I think.

                    'Also hide aria-hidden and hidden attribute elements.

                    'Aaargh, this stops the Google search box from displaying! June 2011. Ah, but not after the page has fully loaded.

                    'It also stops any AUDIO element without controls from displaying, so ignore this attribute if there is an 
                    'AUDIO element, since I plan to always provide a UI for them via WebbIE rather than relying on the web page.

                ElseIf tagname = "TEXTAREA" Or isContentEditable Then
                    'a multi-space text area, TEXTAREA
                    '            OR
                    'The "contentEditable" attribute is set, so this is a rich text edit area (from
                    'IE 5.5 onwards, and in HTML5) -> Treat as a TEXTAREA.
                    'But you can't cast it to an IHTMLTextAreaElement if it isn't one!
                    numTextAreaInputs += 1
                    If numTextAreaInputs > MAX_NUMBER_TEXT_AREA_INPUTS_SUPPORTED - 1 Then
                        mTerminateParsing = True
                    Else
                        textAreaInput(numTextAreaInputs - 1) = element
                        Call mOutput.Append(mNewline)
                        'check for a label
                        label = modFormLabelHandler.GetDescriptiveText(element)
                        If label <> "" Then
                            If Not label.EndsWith(":") Then label &= ":"
                            label &= " "
                        Else
                            'Don't use the Name attribute, it's specific to the code function
                            'of the page. Use Title, it's user-facing.
                            If element.getAttribute("title") Is Nothing Then
                            Else
                                label = element.getAttribute("title").ToString()
                            End If
                            If label <> "" Then
                                If Not label.EndsWith(":") Then label &= ":"
                                label &= " "
                            End If
                        End If
                        Call mOutput.Append(ID_TEXTAREA & IIf(My.Settings.NumberLinks, " " & numTextAreaInputs, "").ToString & ": ")
                        If Len(label) > 0 Then Call mOutput.Append(label)
                        If GetElementAttributeBoolean(element, "disabled") Then Call mOutput.Append(ID_DISABLED & " ")
                        If GetElementAttributeBoolean(element, "readOnly") Then Call mOutput.Append(ID_READONLY & " ")
                        'Rich text area, possibly, so need to parse contents as HTML, not just text.
                        Call ParseChildren(node, True, orderedListPrefix)
                    End If
                ElseIf tagname = "A" Then
                    'Debug.Print "Got A: " & node.tagname & vbTab & Left(element.innerText, 100)
                    'If mblnShowLinks Then 'check we're showing links normally, not non-processing them
                    'Dev: element.getattribute("href").tostring == node.href for As
                    'Also, we always get the complete canonical url UNLESS there is no href attribute at all.
                    If element.getAttribute("href") Is Nothing Then
                        If tagname <> "A" Then
                            href = "FAKEA!"
                        End If
                    Else
                        href = element.getAttribute("href").ToString
                    End If
                    If href <> "" Or hasOnClick Then
                        'okay, this is a standard href link node - that is, it is an anchor to another
                        'location, not a target for another link like <a name="scroll_here">
                        'OR it's a link that is powered by Javascript (3.8.0) such as the Search button on Facebook.
                        'conversely, if it _is_ just a target, the generic picking up of id and name will
                        'catch it
                        If gNumLinks > MAX_NUMBER_LINKS_SUPPORTED - 1 Then
                            'danger will robinson! Hangs WebbIE (3.2.3). Ameliorate,
                            'but don't fix, by exiting
                            If Len(element.innerText.Trim) > 0 Then
                                'got some text, though, so return that.
                                Call mOutput.Append(element.innerText.Trim & " ")
                                mNewline = vbNewLine
                            End If
                            mTerminateParsing = True
                        Else
                            '3.7.4
                            'okay, now parse the link content to get its text content.
                            linkContent = modParseA.ParseAnchor(node)
                            If linkContent = "" Then
                                'we've failed to find anything to put in the link - use the url
                                If href = "FAKEA!" Then
                                    linkContent = modI18N.GetText("Javascript")
                                    'This really sucks on pages like YouTube that have loads of these. Maybe I should
                                    'skip them?
                                    'Debug.Print "OH:" & element.outerhtml
                                Else
                                    'Try to get title of page from database of visited links, if we have it.
                                    linkContent = GetLocationTitle(href)
                                    'If failed, use filename.
                                    If linkContent = "" Then linkContent = ExtractHREFLink(href)
                                End If
                            End If
                            If linkContent = modI18N.GetText("Javascript") Then
                                'Bleh, only Javascript. Dump it.
                            Else
                                gNumLinks += 1
                                gLinks(gNumLinks - 1).element = element
                                gLinks(gNumLinks - 1).description = SanitiseLinkForDisplay(linkContent)
                                gLinks(gNumLinks - 1).address = href
                                'Debug.Print "Link :" & gLinks(gNumLinks).address
                                If mJustDidHeading Then
                                    'we've put "PAGE HEADING: " on the previous line all ready for the page headline,
                                    'it's turned out to be a link. Put the link text in so there is some text.
                                    Call mOutput.Append(linkContent.Trim & " ")
                                    mJustDidHeading = False
                                End If
                                Call mOutput.Append(mNewline & ID_LINK & IIf(My.Settings.NumberLinks, " " & gNumLinks, "").ToString & ": ")
                                Call mOutput.Append(linkContent.Trim & " ") '& " " I'm not sticking a " " at the end of the output because of the prevalence
                                '   of links like this: <a>London</a>, <a>New York</a>, <a>Manchester</a>. Don't want that to be London , New York , etc.
                                'Changed my mind: handle in post-processing. Otherwise I get lots of linkslikethis.
                                mNewline = vbNewLine
                                mJustDidLink = True
                            End If
                        End If
                    Else
                        'okay, this is an a node with no href AND no onclick, so it's simply acting as a target for
                        'another link
                        'This means we have to display the contents, if there are any,
                        'which means going through the child nodes.
                        Call ParseChildren(node, , orderedListPrefix)
                    End If
                    'check for an accesskey on this link
                    If hasAccessKey Then
                        Call modAccesskeys.AddAccessKey(CType(node, mshtml.IHTMLAnchorElement))
                    End If
                ElseIf tagname = "IMG" Then
                    Dim imgElement As mshtml.IHTMLImgElement
                    Dim imgElement2 As mshtml.IHTMLImgElement2
                    imgElement = CType(element, mshtml.IHTMLImgElement)
                    imgElement2 = CType(element, mshtml.IHTMLImgElement2)
                    If hasOnClick Then
                        'Treat it like an A element.
                        gNumLinks += 1
                        If gNumLinks > MAX_NUMBER_LINKS_SUPPORTED - 1 Then
                            'danger will robinson! Hangs WebbIE (3.2.3). Ameliorate,
                            'but don't fix, by exiting
                            If imgElement.alt <> "" Then
                                Call mOutput.Append(imgElement.alt & " ")
                                mNewline = vbNewLine
                            End If
                            mTerminateParsing = True
                        Else
                            'Create new link.
                            gLinks(gNumLinks - 1).element = element
                            If imgElement2.longDesc <> "" Then
                                gLinks(gNumLinks - 1).description = imgElement2.longDesc.ToString
                            End If
                            Call mOutput.Append(mNewline & ID_LINK & IIf(My.Settings.NumberLinks, " " & gNumLinks, "").ToString & ": ")
                            If imgElement.alt = "" Then
                                Call mOutput.Append(modI18N.GetText("Unnamed link or button."))
                            Else
                                Call mOutput.Append(imgElement.alt)
                            End If
                            Call mOutput.AppendLine()
                            mNewline = ""
                        End If
                    ElseIf imgElement2.longDesc <> "" Then
                        'okay, we have a longdesc set: since this is an accessibility feature, better support it
                        'we'll add a link to the URI, labelling it "IMAGE DESCRIPTION" and add the alt tag
                        gNumLinks += 1
                        If gNumLinks > MAX_NUMBER_LINKS_SUPPORTED - 1 Then
                            'danger will robinson! Hangs WebbIE (3.2.3). Ameliorate,
                            'but don't fix, by exiting
                            If imgElement.alt <> "" Then
                                'got some text, though, so return that.
                                Call mOutput.Append(imgElement.alt & " ")
                                mNewline = vbNewLine
                            End If
                            mTerminateParsing = True
                        Else
                            gLinks(gNumLinks - 1).element = element
                            gLinks(gNumLinks - 1).address = imgElement2.longDesc
                            Call mOutput.Append(mNewline & ID_LINK & IIf(My.Settings.NumberLinks, " " & gNumLinks, "").ToString & modI18N.GetText(": IMAGE DESCRIPTION"))
                            gLinks(gNumLinks).description = modI18N.GetText("IMAGE DESCRIPTION")
                            If imgElement.alt <> "" Then
                                Call mOutput.Append(" " & modI18N.GetText("FOR") & " " & imgElement.alt)
                                gLinks(gNumLinks - 1).description = modI18N.GetText("IMAGE DESCRIPTION FOR") & " " & imgElement.alt
                            End If
                            Call mOutput.AppendLine()
                            mNewline = ""
                        End If
                    Else
                        Call mOutput.Append(ParseIMGorAREA(element))
                    End If
                ElseIf tagname = "SELECT" Then
                    'A select node - has option and optgroup child nodes
                    numSelects = numSelects + 1
                    selects(numSelects - 1).node = CType(element, mshtml.IHTMLSelectElement)
                    If numSelects > MAX_NUMBER_SELECTS_SUPPORTED Then
                        'Whoops! Too many selects for WebbIE. 3.8.0
                        mTerminateParsing = True
                    Else
                        If hasName Then
                            selects(numSelects - 1).name = selects(numSelects - 1).node.name 'store name of menu
                        End If
                        'now go through the options in the select node
                        nodeIterator = node.firstChild
                        i = 0
                        While Not (nodeIterator Is Nothing)
                            'got a node - check it's an option
                            If nodeIterator.nodeName = "OPTION" Then
                                'got an option node
                                Dim o As mshtml.IHTMLOptionElement = CType(nodeIterator, mshtml.IHTMLOptionElement)
                                i += 1
                                selects(numSelects - 1).options(i) = o.text 'save menu items
                                If o.selected Then
                                    selects(numSelects - 1).selected = CShort(i)
                                End If
                            ElseIf nodeIterator.nodeName = "OPTGROUP" Then  ' check for option groups
                                'we have an option group - have to find the children of this one
                                nodeIterator2 = nodeIterator.firstChild
                                While Not (nodeIterator2 Is Nothing)
                                    'got a node or option group
                                    If nodeIterator2.nodeName = "OPTION" Then
                                        'got an option node
                                        i += 1
                                        'TODO I'm not happy with this loop. I fear nested elements may not be
                                        'caught by my not using recursion. 
                                        Dim o As mshtml.IHTMLOptionElement = CType(nodeIterator2, mshtml.IHTMLOptionElement)
                                        selects(numSelects - 1).options(i) = o.text
                                        If o.selected Then selects(numSelects - 1).selected = i
                                    End If
                                    nodeIterator2 = nodeIterator2.nextSibling
                                End While
                                'go back to the parent option groups
                            End If
                            nodeIterator = nodeIterator.nextSibling
                        End While
                        selects(numSelects - 1).size = i ' size of the menu
                        'now do the display
                        Call mOutput.Append(mNewline)
                        'check for a label
                        If hasId Then
                            label = modFormLabelHandler.GetDescriptiveText(element)
                        Else
                            label = ""
                        End If
                        If label <> "" Then
                            Call mOutput.AppendLine(label)
                        Else
                            'Check for a title attribute.
                            If element.title <> "" Then
                                label = element.title
                            End If
                            If label = "" Then
                                If hasName Then
                                    label = selects(numSelects - 1).node.name
                                End If
                            End If
                            If label <> "" Then
                                Call mOutput.AppendLine(label)
                            End If
                        End If
                        Call mOutput.Append(ID_SELECT & IIf(My.Settings.NumberLinks, " " & numSelects, "").ToString & ": (")
                        'display the currently-selected item
                        If selects(numSelects - 1).selected > 0 Then
                            Call mOutput.Append(selects(numSelects - 1).options(selects(numSelects - 1).selected))
                        End If
                        Call mOutput.Append(")")
                        If selects(numSelects - 1).node.disabled Then
                            Call mOutput.Append(ID_DISABLED)
                        End If
                        Call mOutput.AppendLine("")
                        mNewline = ""
                    End If
                ElseIf tagname = "INPUT" Then
                    'an input feature, the type of which is to be determined
                    Dim inputElement As mshtml.IHTMLInputElement
                    inputElement = CType(element, mshtml.IHTMLInputElement)
                    'I could put support for aria-hidden in here - next time!
                    Select Case inputElement.type
                        Case "button" 'an input button
                            numButtonInputs += 1
                            Call mOutput.Append(mNewline)
                            'check for a label
                            label = DeviseLabel(element)
                            If label <> "" Then
                                Call mOutput.AppendLine(label)
                            End If
                            If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                            If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                            Call mOutput.Append(ID_BUTTON & IIf(My.Settings.NumberLinks, " " & numButtonInputs, "").ToString & ": (")
                            If inputElement.value = "" Then
                                If Len(Trim(Replace(element.innerText, vbNewLine, ""))) = 0 Then
                                    Call mOutput.Append(GetText("Button"))
                                Else
                                    Call mOutput.Append(Trim(Replace(element.innerText, vbNewLine, "")))
                                End If
                            Else
                                Call mOutput.Append(inputElement.value)
                            End If
                            Call mOutput.Append(")" & vbNewLine)
                            mNewline = ""
                            buttonInput(numButtonInputs - 1) = element
                        Case "checkbox" ' a checkbox
                            numCheckboxInputs += 1
                            checkboxInput(numCheckboxInputs - 1) = inputElement
                            'display the checkbox
                            Call mOutput.Append(mNewline)
                            'check for a label
                            label = DeviseLabel(element)
                            If label <> "" Then
                                If Not label.EndsWith(":") Then label &= ":"
                            End If
                            If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                            If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                            Call mOutput.Append(ID_CHECKBOX & IIf(My.Settings.NumberLinks, " " & numCheckboxInputs, "").ToString & ": ")
                            If Len(label) > 0 Then Call mOutput.Append(label)
                            'display whether it's checked
                            If inputElement.checked Then
                                Call mOutput.Append(" " & ID_CHECKED)
                            Else
                                Call mOutput.Append(" " & ID_NOTCHECKED)
                            End If
                            mNewline = vbNewLine
                        Case "radio" ' a radio button
                            numRadioInputs = numRadioInputs + 1
                            radioInput(numRadioInputs - 1) = element
                            Call mOutput.Append(mNewline)
                            'check for a label
                            label = DeviseLabel(element)
                            If label <> "" Then
                                If Not label.EndsWith(":") Then label &= ":"
                            End If
                            Call mOutput.Append(ID_RADIO & IIf(My.Settings.NumberLinks, " " & numRadioInputs, "").ToString & ": ")
                            If Len(label) > 0 Then Call mOutput.Append(label & " ")
                            If inputElement.checked Then
                                Call mOutput.Append(ID_SELECTED)
                            Else
                                Call mOutput.Append(ID_NOTSELECTED)
                            End If
                            If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                            If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                            mNewline = vbNewLine
                        Case "password" ' password input box
                            numPassInputs += 1 'number of password inputs on page
                            If numPassInputs > MAX_NUMBER_PASSWORD_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                'store the password entry box
                                passwordInput(numPassInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = DeviseLabel(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                'If element.getattribute("disabled").tostring Then call moutput.append(ID_DISABLED
                                'If element.getattribute("readonly").tostring Then call moutput.append(ID_READONLY
                                Call mOutput.Append(ID_PASSWORD & IIf(My.Settings.NumberLinks, " " & numPassInputs, "").ToString & ": ")
                                Call mOutput.Append(New String(CChar("*"), Len(inputElement.value)) & vbNewLine)
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                mNewline = ""
                            End If
                        Case "submit" ' a submit button
                            numSubmitInputs += 1 'increase submit box counter
                            If numSubmitInputs > MAX_NUMBER_SUBMITS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = DeviseLabel(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                Call mOutput.Append(ID_SUBMIT & IIf(My.Settings.NumberLinks, " " & numSubmitInputs, "").ToString & ": (" & inputElement.value & ")")
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                Call mOutput.Append(vbNewLine)
                                mNewline = ""
                                submitInput(numSubmitInputs - 1) = element
                            End If
                        Case "hidden" ' hidden input information
                            'don't display anything!
                        Case "image" ' an image
                            numSubmitInputs += 1 'increase submit box counter
                            If numSubmitInputs > MAX_NUMBER_SUBMITS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                submitInput(numSubmitInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = DeviseLabel(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                Call mOutput.Append(ID_SUBMIT & IIf(My.Settings.NumberLinks, " " & numSubmitInputs, "").ToString & ": (")
                                'we _could_ check the alt tag to see if were blank, but if it were we wouldn't
                                'be able to put any more useful information in anyway (except possibly the file
                                'name of the image, which might be a forty-character absolute URL...)
                                Call mOutput.Append(inputElement.alt & ")")
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                Call mOutput.Append(vbNewLine)
                                mNewline = ""
                            End If
                        Case "file" ' browse-to and select file option
                            numFileInputs += 1
                            If numFileInputs > MAX_NUMBER_FILE_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                fileInput(numFileInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = DeviseLabel(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                Call mOutput.Append(ID_FILE & IIf(My.Settings.NumberLinks, " " & numFileInputs, "").ToString & ": (")
                                If inputElement.value = "" Then
                                    Call mOutput.Append(modI18N.GetText("No file selected)") & vbNewLine)
                                Else
                                    Call mOutput.Append(inputElement.value & ")")
                                End If
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                Call mOutput.Append(vbNewLine)
                                mNewline = ""
                            End If
                        Case "email" ' an email input box.
                            numEmailInputs = numEmailInputs + 1
                            If numEmailInputs > MAX_NUMBER_EMAIL_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                emailInput(numEmailInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = DeviseLabel(element)
                                Call mOutput.Append(ID_EMAILINPUT & IIf(My.Settings.NumberLinks, " " & numTextInputs, "").ToString & ": ")
                                If Len(label) > 0 Then Call mOutput.Append(label & " ")
                                'display contents if any
                                Call mOutput.Append(inputElement.value)
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                Call mOutput.AppendLine()
                                mNewline = ""
                            End If
                        Case "range" ' a slider
                            numRangeInputs = numRangeInputs + 1
                            If numResetInputs > MAX_NUMBER_RANGE_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                rangeInput(numRangeInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'Check for a label
                                label = DeviseLabel(element)
                                Call mOutput.Append(ID_RANGEINPUT & IIf(My.Settings.NumberLinks, " " & numRangeInputs, "").ToString & ": ")
                                If Len(label) > 0 Then Call mOutput.Append(label & " ")
                                'display contents if any
                                Call mOutput.Append(inputElement.value)
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                Call mOutput.AppendLine()
                                mNewline = ""
                            End If
                            'Case "text" ' a text input box
                            'Case "search" ' an HTML5 search box. TODO Finish!
                            '    numSearchInputs = numSearchInputs + 1
                            '    If numSearchInputs > MAX_NUMBER_SEARCH_INPUTS_SUPPORTED Then
                            '        mTerminateParsing = True
                            '    Else
                            '        searchInput(numSearchInputs - 1) = element
                            '        Call mOutput.Append(mNewline)
                            '        'check for a label
                            '        label = DeviseLabel(element)
                            '    End If
                            '    'TODO!!
                        Case Else ' "text" is the default if input type is not specified or recognised.
                            ' -- for example, HTML5 new input types.
                            'normal text input
                            numTextInputs = numTextInputs + 1
                            If numButtonInputs > MAX_NUMBER_TEXT_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                textInput(numTextInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = DeviseLabel(element)
                                Call mOutput.Append(ID_TEXTBOX & IIf(My.Settings.NumberLinks, " " & numTextInputs, "").ToString & ": ")
                                If Len(label) > 0 Then Call mOutput.Append(label & " ")
                                'display contents if any
                                Call mOutput.Append(inputElement.value)
                                If inputElement.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                If inputElement.readOnly Then Call mOutput.Append(ID_READONLY & " ")
                                Call mOutput.AppendLine()
                                mNewline = ""
                            End If
                    End Select
                ElseIf tagname = "VIDEO" Then
                    'HTML5 Video.
                    numVideos = numVideos + 1
                    If numVideos > MAX_NUMBER_VIDEOS_SUPPORTED Then
                        mTerminateParsing = True
                    Else
                        videos(numVideos - 1) = element
                        Call mOutput.Append(mNewline)
                        Call mOutput.Append(ID_VIDEO & vbNewLine)
                        mNewline = ""
                    End If
                ElseIf tagname = "AUDIO" Then
                    'HTML5 audio
                    numAudios = numAudios + 1
                    Debug.Print("Got AUDIO " & numAudios)
                    If numAudios > MAX_NUMBER_AUDIOS_SUPPORTED Then
                        mTerminateParsing = True
                    Else
                        audios(numAudios - 1) = element
                        Call mOutput.Append(mNewline)
                        Call mOutput.Append(ID_AUDIO & vbNewLine)
                        mNewline = ""
                    End If
                ElseIf tagname = "BUTTON" Then
                    'an input button. TODO These can contain HTML, should parse it!
                    Dim button As mshtml.IHTMLButtonElement = CType(element, mshtml.IHTMLButtonElement)
                    Select Case element.getAttribute("type").ToString
                        Case "button"
                            numButtonInputs = numButtonInputs + 1
                            If numButtonInputs > MAX_NUMBER_BUTTON_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                buttonInput(numButtonInputs - 1) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = modFormLabelHandler.GetDescriptiveText(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                Call mOutput.Append(ID_BUTTON & IIf(My.Settings.NumberLinks, " " & numButtonInputs, "").ToString & ": (")
                                labelText = Parse(node)
                                If Len(labelText) = 0 Then
                                    Call mOutput.Append(button.value)
                                Else
                                    Call mOutput.Append(labelText)
                                End If
                                Call mOutput.Append(")")
                                If button.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                Call mOutput.Append(vbNewLine)
                                mNewline = ""
                            End If
                        Case "reset"
                            numResetInputs = numResetInputs + 1
                            If numResetInputs > MAX_NUMBER_RESET_INPUTS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                resetInput(numResetInputs) = element
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = modFormLabelHandler.GetDescriptiveText(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                '3.8.0 - some buttons don't have value attributes, they just have content to use.
                                Call mOutput.Append(ID_RESET & IIf(My.Settings.NumberLinks, " " & numResetInputs, "").ToString & ": (")
                                labelText = Parse(node)
                                If Len(labelText) = 0 Then
                                    Call mOutput.Append(button.value)
                                Else
                                    Call mOutput.Append(labelText)
                                End If
                                Call mOutput.Append(")")
                                If button.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                Call mOutput.Append(vbNewLine)
                                mNewline = ""
                            End If
                        Case "submit"
                            numSubmitInputs += 1 'increase submit box counter
                            If numSubmitInputs > MAX_NUMBER_SUBMITS_SUPPORTED Then
                                mTerminateParsing = True
                            Else
                                Call mOutput.Append(mNewline)
                                'check for a label
                                label = modFormLabelHandler.GetDescriptiveText(element)
                                If label <> "" Then
                                    If Not label.EndsWith(":") Then label = label & ":"
                                    Call mOutput.Append(label & vbNewLine)
                                End If
                                '3.8.0 Buttons may not have a .value attribute, they may have innerText instead.
                                Call mOutput.Append(ID_SUBMIT & IIf(My.Settings.NumberLinks, " " & numSubmitInputs, "").ToString & ": (")
                                labelText = Parse(node)
                                If Len(labelText) = 0 Then
                                    Call mOutput.Append(button.value)
                                Else
                                    Call mOutput.Append(labelText)
                                End If
                                Call mOutput.Append(")")
                                If button.disabled Then Call mOutput.Append(ID_DISABLED & " ")
                                Call mOutput.Append(vbNewLine)
                                mNewline = ""
                                submitInput(numSubmitInputs - 1) = element
                            End If
                    End Select
                ElseIf tagname = "HR" Then
                    'a horizontal rule. TODO This sucks for screenreaders.
                    Call mOutput.AppendLine("_____________________________________________________")
                    mNewline = ""
                ElseIf tagname = "IFRAME" Or tagname = "FRAME" Then
                    'Uncomment this to remove iframe support. The advantage of this is that it removes
                    'lots of adverts. 
                    Dim cfie As CrossFrameIE = New CrossFrameIE
                    Dim frameElement As mshtml.IHTMLFrameBase2 = CType(element, mshtml.IHTMLFrameBase2)
                    Dim frameWindow As mshtml.IHTMLWindow2 = frameElement.contentWindow
                    Dim frameDoc As mshtml.IHTMLDocument2 = cfie.GetDocumentFromWindow(frameWindow)
                    If frameDoc Is Nothing Then
                        'Could not access this frame: don't include its contents.
                        System.Diagnostics.Debug.Print("Failed to get IFRAME")
                    Else
                        Dim bodyNode As mshtml.IHTMLDOMNode = CType(frameDoc.body, mshtml.IHTMLDOMNode)
                        Call ParseNode(bodyNode, orderedListPrefix)
                    End If
                ElseIf tagname = "AREA" Then
                        ' client-side image map - treat like a collection of links
                        ' first check it has an href - if not, discard it
                        Dim areaElement As mshtml.IHTMLAreaElement
                        areaElement = CType(element, mshtml.IHTMLAreaElement)
                        If areaElement.noHref Or areaElement.href = "" Then
                            'Treat as an image.
                            Call mOutput.Append(ParseIMGorAREA(element))
                        Else
                        'Treat as a link.
                        gNumLinks += 1
                        If areaElement.alt <> "" Then
                                'use the alt tag
                                linkContent = Trim(areaElement.alt)
                            Else
                                'no alt tag - use address
                                linkContent = ExtractHREFLink(areaElement.href)
                            End If
                            gLinks(gNumLinks).address = areaElement.href
                            Call mOutput.Append(mNewline)
                            Call mOutput.Append(ID_LINK & IIf(My.Settings.NumberLinks, " " & gNumLinks, "").ToString & ":   " & Trim(linkContent) & " ")
                            mNewline = vbNewLine
                        End If
                    ElseIf tagname = "BLOCKQUOTE" Then
                        'a quotation
                        Call mOutput.Append(mNewline & modI18N.GetText("[quotation marks]"))
                        mNewline = vbNewLine
                        Call ParseChildren(node, , orderedListPrefix)

                        If mOutput.ToString.Substring(mOutput.ToString.Length - 1, 1) = " " Then mOutput.Remove(mOutput.Length - 1, 1)
                        Call mOutput.Append(modI18N.GetText("[quotation marks]") & vbNewLine)
                        mNewline = ""
                    ElseIf tagname = "OBJECT" Or tagname = "APPLET" Then
                        'Nothing much I can do about that! I could render something, but then the user
                        'gets some kind of fallback message that says "Your browser does not work!"
                        'and gets all sad.
                    ElseIf tagname = "STYLE" Then
                        'content for browser - not intended for user to see
                    ElseIf tagname = "SCRIPT" Then
                        'code - not intended for user.
                    ElseIf tagname = "CAPTION" Then
                        'Table caption
                        Call mOutput.Append(mNewline & modI18N.GetText("Table Caption") & ": " & element.innerText & vbNewLine)
                        mNewline = ""
                    ElseIf tagname = "DEL" Then
                        'deleted content - don't do any of it or its child nodes! Visual users would
                        'probably see the text struck-through, hard for screenreader users.
                    ElseIf tagname = "LABEL" Then
                        'okay, this is a label node, so either (1) it contains information
                        'relating to a control node or (2) it contains a control node.
                        'If (1) don't display it: it'll be displayed when we do the control
                        'node. If (2) process normally. However, even if it has a "for" attribute
                        'explicitly assigning it to an input element, we must still check that it
                        'does not also itself contain an input element. If we does then (2) we 
                        'process it normally - otherwise the input element won't get rendered.
                        Dim forAttribute As String = ""
                        If element.getAttribute("for") Is Nothing Then
                        Else
                            forAttribute = element.getAttribute("for").ToString()
                        End If

                        If ContainsInputElement(element) Or forAttribute = "" Then
                            'This contains content including the input control, so we have to parse it normally. OR
                            'This does not have a "for" attribute, so it must refer to the contents of the LABEL
                            'node, so we do kids.
                            'DEV: This might be messy, but otherwise lots of LABEL action
                            'doesn't work.
                            'record non-form text and use it
                            'as the label for any form nodes
                            'Find the control to be labelled
                            controlToBeLabelled = FindInputElement(element)
                            If controlToBeLabelled Is Nothing Then
                                'Didn't find the required child element of element that is an input
                                'element - whoops, mistake on the part of the page coder. OK, 
                                'don't work out any labelling.
                            Else
                                'Got the element to be labelled.
                                labelText = ParseForText(node)
                                'Apply label text to the control we should already have got:
                                Call modFormLabelHandler.AddLabel(labelText, controlToBeLabelled)
                                'Do the child input element (even if there isn't one!)
                                Call ParseNode(CType(controlToBeLabelled, mshtml.IHTMLDOMNode), orderedListPrefix)
                            End If
                        End If
                    ElseIf tagname = "BDO" Then
                        'set-direction text - right-to-left or left-to-right
                        If element.getAttribute("dir") Is Nothing Then
                            'No dir specified: can't do anything with this. Parse children normally.
                            Call ParseChildren(node, , orderedListPrefix)
                        ElseIf element.getAttribute("dir").ToString = "ltr" Then
                            '"normal" left-to-right direction - process as "normal" text
                            Call ParseChildren(node, , orderedListPrefix)
                        Else
                            'reversed right-to-left direction
                            linkContent = ""
                            If node.hasChildNodes Then
                                ni = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                                For Each nodeIterator In ni
                                    linkContent = linkContent & ParseNodeText(nodeIterator)
                                    If mTerminateParsing Then Exit For
                                Next nodeIterator
                            End If
                            i = Len(linkContent)
                            While i > 0
                                If Mid(linkContent, i, 1) = vbLf Then
                                    'we've found a newline (CR + LF)
                                    Call mOutput.AppendLine()
                                    mNewline = ""
                                    i = i - 2
                                Else
                                    'normal character - add to output
                                    Call mOutput.Append(Mid(linkContent, i, 1))
                                    mNewline = vbNewLine
                                    i = i - 1
                                End If
                            End While
                        End If
                    ElseIf tagname = "PROGRESS" Then
                        'HTML5 progress bar! Now, if it is missing a "value" attribute, then it is an
                        'indeterminate progress bar. However, IE10 (Feb 2013) reports a value of "0",
                        'not that there is nothing there. So check for the "value" string in the HTML.
                        If element.outerHTML.ToLowerInvariant().Contains("value") Then
                            'Progress bar with value.
                            Dim maxValue As Double
                            If Double.TryParse(element.getAttribute("max").ToString(), maxValue) Then
                                'OK, got a max.
                            Else
                                maxValue = 1 ' default
                            End If
                            Dim currentValue As Double
                            If Double.TryParse(element.getAttribute("value").ToString(), currentValue) Then
                                'OK, got value
                                mOutput.Append(mNewline & modI18N.GetText("Progress") & " " & Math.Round(currentValue * 100.0 / maxValue, 0).ToString() & "%" & vbNewLine)
                                mNewline = ""
                            Else
                                'Failed to parse value! Broken HTML. I think I'll choose to write nothing, 
                                'so as to avoid causing confusion.
                            End If
                        Else
                            'This is an indeterminate progress bar: that is, it's a "you have to wait!"
                            'indicator. I could output "Please wait..."? Not sure if that is much use.
                            mOutput.Append(mNewline & modI18N.GetText("Please wait...") & vbNewLine)
                            mNewline = ""
                        End If

                        'don't do the sub-kids of an A node, an INPUT node, a FRAME node, an IMG node...
                        'TODO But what if it's an A node that contains an H2 node? Dumb, but possible!
                    Else
                    If tagname = "DD" Then
                        'We're not in an ordered list, that's handled above. So just stick on a newline
                        Call mOutput.Append(mNewline) '& "- " ' DEV: don't put the - in because it might be
                        'a list of links following, and this produces lots of erroneous lines with - on them
                        mNewline = ""
                    ElseIf tagname = "Q" Then
                        Call mOutput.Append(modI18N.GetText("[quotation marks]"))
                        mNewline = vbNewLine
                    ElseIf tagname = "ARTICLE" Then ' TODO Do work for Main
                        'HTML5 Article
                        mInArticle = mInArticle + 1
                        If mInArticle = 1 Then
                            Call mOutput.AppendLine(modI18N.GetText("Article Start"))
                            mNewline = ""
                        End If
                        gPageHasAnArticle = True
                    ElseIf tagname = "TABLE" Then
                        'OK, we've got a table node. Do we apply our exciting new table system? Most tables will simply be
                        'markup tables. So we don't want to confuse pages with all the crap.
                        'Let's assume, 3.8.0, Dec 2008, that having any two of caption, summary and TH nodes means it's
                        'a real table.
                        If element.getAttribute("summary") Is Nothing Then
                        Else
                            If element.getAttribute("summary").ToString <> "" Then validityCount = validityCount + 1
                        End If
                        If InStr(1, element.innerHTML, "<caption", CompareMethod.Text) > 0 Then validityCount = validityCount + 1
                        'Some tables are just layout tables but use th:
                        If InStr(1, element.innerHTML, "<th", CompareMethod.Text) > 0 And InStr(1, element.innerHTML, "<td", CompareMethod.Text) > 0 Then validityCount = validityCount + 1
                        If validityCount > 1 Then
                            numTables = numTables + 1
                            Call mOutput.Append(mNewline & ID_TABLE & IIf(My.Settings.NumberLinks, " " & numTables, "").ToString & ": ")
                            If element.getAttribute("summary") Is Nothing Then
                            Else
                                Call mOutput.Append(element.getAttribute("summary").ToString)
                            End If
                            Call mOutput.AppendLine()
                            tables(numTables) = element
                            mNewline = ""
                        End If
                    ElseIf tagname = "H1" Then
                        Call mOutput.Append(mNewline & SECTION_MARKER_H1 & ": ")
                        mNewline = vbNewLine
                    ElseIf tagname = "H2" Then
                        Call mOutput.Append(mNewline & SECTION_MARKER_H2 & ": ")
                        mNewline = vbNewLine
                    ElseIf tagname = "H3" Then
                        Call mOutput.Append(mNewline & SECTION_MARKER_H3 & ": ")
                        mNewline = vbNewLine
                    ElseIf tagname = "H4" Then
                        Call mOutput.Append(mNewline & SECTION_MARKER_H4 & ": ")
                        mNewline = vbNewLine
                    ElseIf tagname = "H5" Then
                        Call mOutput.Append(mNewline & SECTION_MARKER_H5 & ": ")
                        mNewline = vbNewLine
                    ElseIf tagname = "H6" Then
                        Call mOutput.Append(mNewline & SECTION_MARKER_H6 & ": ")
                        mNewline = vbNewLine
                    ElseIf tagname = "MAIN" Then
                        'HTML5 "main" element. Requires newline before.
                        Call mOutput.Append(vbNewLine & "► ")
                        mNewline = ""
                        gPageHasMain = True
                        'Debug.Print("Unprocessed tagname:" & tagname)
                    End If
                    'Parse children.
                    Call ParseChildren(node, , orderedListPrefix)
                    'do newline if the node is one requiring a newline
                    If (tagname = "BR" Or tagname = "H1" Or tagname = "H2" Or
                        tagname = "H3" Or tagname = "H4" Or tagname = "H5" Or
                        tagname = "H6" Or tagname = "ADDRESS" Or tagname = "CENTER" Or
                        tagname = "PRE" Or tagname = "LI" Or tagname = "DD" Or
                        tagname = "DT" Or tagname = "TR" Or tagname = "LEGEND" Or
                        tagname = "FIELDSET" Or tagname = "DIV") And mNewline = vbNewLine Then
                        Call mOutput.AppendLine()
                        mNewline = ""
                        mJustDidLink = False
                    ElseIf tagname = "MAIN" Then
                        Call mOutput.Append(" ◄" & vbNewLine)
                        mNewline = ""
                        mJustDidLink = False
                    ElseIf (tagname = "P" Or tagname = "DL" Or tagname = "OL" Or tagname = "UL" Or tagname = "TABLE") And mNewline = vbNewLine Then
                        Call mOutput.AppendLine(vbNewLine)
                        mNewline = ""
                        mJustDidLink = False
                    ElseIf tagname = "TD" Or tagname = "TH" Then
                        If Not mOutput.ToString.EndsWith(" ") And Not mOutput.ToString.EndsWith(vbNewLine) Then
                            Call mOutput.Append(" ")
                            mNewline = vbNewLine
                        End If
                    ElseIf tagname = "Q" Then
                        If mOutput.ToString.EndsWith(" ") Then mOutput.Remove(mOutput.Length - 1, 1)
                        Call mOutput.Append(modI18N.GetText("[quotation marks]"))
                        mNewline = vbNewLine
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Private Function GetElementAttribute(ByVal element As mshtml.IHTMLElement, ByVal attribute As String) As String
        Try
            If element.getAttribute(attribute) Is Nothing Then
                Return ""
            Else
                Return element.getAttribute(attribute).ToString
            End If
        Catch
            Return ""
        End Try
    End Function

    Private Function GetElementAttributeBoolean(ByVal element As mshtml.IHTMLElement, ByVal attribute As String) As Boolean
        Try
            If element.getAttribute(attribute) Is Nothing Then
                Return False
            ElseIf element.getAttribute(attribute).ToString() = "" Then
                Return False
            Else
                Try
                    Return CBool(element.getAttribute(attribute).ToString())
                Catch
                    'Because disabled="disabled" is true, default to true. 
                    Return True
                End Try
            End If
        Catch
            Return False
        End Try
    End Function

    Private Function DeviseLabel(ByVal element As mshtml.IHTMLElement) As String
        Try
            Dim label As String

            label = modFormLabelHandler.GetDescriptiveText(element)
            If label = "" Then
                'Parse for text. 
                label = ParseForText(CType(element, mshtml.IHTMLDOMNode))
            End If
            If label <> "" Then
                If Not label.EndsWith(":") Then label = label & ":"
            End If
            Return label.Trim
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Iterate through the children of node, sending them to ParseNode to get parsed.
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="addNewlines"></param>
    ''' <param name="orderedListPrefix"></param>
    Private Sub ParseChildren(ByVal node As mshtml.IHTMLDOMNode, Optional ByVal addNewlines As Boolean = False, Optional ByVal orderedListPrefix As String = "")
        Try
            Dim ni As mshtml.IHTMLDOMChildrenCollection
            Dim nodeIterator As mshtml.IHTMLDOMNode
            Dim listCount As Integer = 1
            Dim numberedList As Boolean
            Dim unNumberedList As Boolean

            'Is this a list node? If so, we'll need to add numbers or bullet points to children. 
            If node.nodeName = "OL" Then
                numberedList = True
            ElseIf node.nodeName = "UL" Then
                unNumberedList = True
            End If
            'OK, now do children.
            If node.hasChildNodes Then
                ni = CType(node.childNodes, mshtml.IHTMLDOMChildrenCollection)
                For Each nodeIterator In ni
                    'If we are in a list and this is a list item then number or bullet-point it.
                    If numberedList Then
                        If nodeIterator.nodeName = "LI" Then
                            Call mOutput.Append(mNewline)
                            'TODO Handle lists: breaks GetItemIndex at present. 
                            'Call mOutput.Append(orderedListPrefix & listCount & ". ")
                            listCount = listCount + 1
                            mNewline = ""
                        End If
                    ElseIf unNumberedList Then
                        If nodeIterator.nodeName = "LI" Then
                            Call mOutput.Append(mNewline)
                            'TODO Handle lists: breaks GetItemIndex at present. 
                            'Call mOutput.Append(ChrW(8226) & " ")
                            mNewline = ""
                        End If
                    End If
                    'Now parse the child using ParseNode. Need to pass through any list item start.
                    If numberedList Then
                        Call ParseNode(nodeIterator, orderedListPrefix & listCount - 1 & ".")
                    Else
                        Call ParseNode(nodeIterator, orderedListPrefix)
                    End If
                    'Check for "cancel" so we don't hang.
                    If mTerminateParsing Then Exit For
                    'Some set of children - TEXTAREA contents - need newlines added to each
                    'child so they look correct. TODO I should probably test this!
                    If addNewlines Then
                        mOutput.Append(mNewline)
                        mNewline = ""
                    End If
                Next nodeIterator
            End If
        Catch
        End Try
    End Sub


    ''' <summary>
    ''' stops all browser activity and ticking
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StopBrowsers()
        Call StopBusyAnimation()
        Call modGlobals.gWebHost.webMain.Stop()
    End Sub

    Private Function ParseIMGorAREA(ByVal element As mshtml.IHTMLElement) As String
        Try
            Dim imgElement As mshtml.IHTMLImgElement
            If My.Settings.ShowImages Then
                'okay, we're displaying images and we've found one: present the alt info UNLESS
                'the alt tag is "*", " " or ""
                imgElement = CType(element, mshtml.IHTMLImgElement)
                If imgElement.alt <> "" And imgElement.alt <> "*" Then
                    ParseIMGorAREA = mNewline & modI18N.GetText("IMAGE:") & " " & imgElement.alt & vbNewLine
                    If imgElement.src <> "" Then
                        Call gImageHrefs.Add(imgElement.src)
                    End If
                    mNewline = ""
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Sub RefreshCurrentPage()
        Try
            'Call this when you want to redisplay the current page without reloading it from the server -
            'for example when you've amended the DOM and want to show the differences

            Dim caretPosition As Integer ' position of caret

            Debug.Print("REFRESH CURRENT PAGE")
            'store the caret position and line number
            caretPosition = txtText.SelectionStart
            'clear text box
            txtText.Clear()
            'Parse page
            Call ParseDocument()
            If mSeekingInternalTarget Then
                Call GotoInternalTarget()
            ElseIf mElementWithFocus Is Nothing Then
                'simple refresh, restore the caret position
                txtText.SelectionStart = caretPosition
                txtText.SelectionLength = 0
            Else
                'we had an element that took the focus, need to move to it.
                mElementWithFocus = Nothing
                Dim i As Integer = InStr(1, txtText.Text, AJAX_TARGET_MARKER)
                If i > 0 Then
                    txtText.SelectionStart = i - 1
                    txtText.SelectionLength = Len(AJAX_TARGET_MARKER)
                    txtText.SelectedText = ""
                Else
                    'failed to find marker for ajax element, for some reason.
                    'Leave the cursor where it was.
                    txtText.SelectionStart = caretPosition
                    txtText.SelectionLength = 0
                End If
            End If
            Call txtText.ScrollToCaret()
            Call UpdateMenus()
            Call Me.txtText.Focus()
        Catch
        End Try
    End Sub

    Private Sub SkipToForm(ByRef direction As Integer)
        Try
            If GotoLineStarting(GetText("Form"), direction) Then
                'Don't select - stops screenreaders reading line, they read the document instead.
                'Call SelectCurrentParagraph
            End If
        Catch
        End Try
    End Sub

    Private Sub SkipLinks(ByRef direction As Integer, Optional ByRef reverse As Boolean = False)
        Try
            'Causes the cursor to jump to the start of the next paragraph over any links
            'If reverse is set then goes to the first line that IS a link or whatever.
            Dim lineNumber As Integer
            Dim found As Boolean
            Dim line As String
            Dim numberLines As Integer
            Dim startPreviousLine As Integer
            Dim endThisLine As Integer
            Dim foundNewline As Integer

            lineNumber = GetCurrentLineIndex(txtText) + direction
            numberLines = GetNumberOfLines(txtText)
            found = False
            While lineNumber < numberLines And lineNumber >= 0 And Not found
                'Debug.Print "Linenumber:" & lineNumber
                line = GetLine(txtText, lineNumber).Trim
                'Debug.Print "Line:[" & line & "]"
                If Not line.Contains(ID_LINK) And InStr(1, line, ID_SELECT, CompareMethod.Text) <> 1 And InStr(1, line, ID_BUTTON, CompareMethod.Text) <> 1 And InStr(1, line, ID_CHECKBOX, CompareMethod.Text) <> 1 And InStr(1, line, ID_RADIO, CompareMethod.Text) <> 1 And InStr(1, line, ID_TEXTBOX, CompareMethod.Text) <> 1 And InStr(1, line, ID_RANGEINPUT, CompareMethod.Binary) <> 1 And InStr(1, line, ID_EMAILINPUT, CompareMethod.Text) <> 1 And InStr(1, line, ID_PASSWORD, CompareMethod.Text) <> 1 And InStr(1, line, ID_SUBMIT, CompareMethod.Text) <> 1 And InStr(1, line, ID_FILE, CompareMethod.Text) <> 1 And InStr(1, line, ID_RESET, CompareMethod.Text) <> 1 And InStr(1, line, ID_TEXTAREA, CompareMethod.Text) <> 1 Then
                    'ASSERTION: The current line is none of the "focusable" items on the page, like a link.
                    'Debug.Print "Structural"
                    If reverse Then
                        'Not any of the types we're looking for, definitely not found.
                        found = False
                    Else
                        'hurrah, it's not a WebbIE structural item!  Now check it has some content
                        If UBound(Split(Trim(line), " ")) > 1 Then
                            'And some content: finally, the question is whether this line is
                            'the start of a paragraph or just some text.
                            'Debug.Print "content:" & line
                            If lineNumber > 0 Then
                                startPreviousLine = modAPIFunctions.GetCharacterIndexOfLine(txtText, lineNumber - 1)
                                If startPreviousLine = 0 Then startPreviousLine = 1
                                endThisLine = modAPIFunctions.GetCharacterIndexOfLine(txtText, lineNumber) + Len(line) - 2
                                foundNewline = InStr(startPreviousLine, txtText.Text, vbNewLine)
                                If foundNewline = 0 Then foundNewline = InStr(startPreviousLine, txtText.Text, vbCr)
                                If foundNewline = 0 Then foundNewline = InStr(startPreviousLine, txtText.Text, vbLf)
                                If (foundNewline > 0) And (foundNewline < endThisLine) Then
                                    'okay, this line is directly preceded by a newline
                                    found = True
                                Else
                                    'nope, no newline: this is part of a paragraph, or end of the page, so skip
                                    lineNumber = lineNumber + direction
                                End If
                            Else
                                'okay, we're at the start of the text, so call it found
                                found = True
                            End If
                        End If
                    End If
                Else
                    If reverse Then
                        'Found a link etc., good!
                        found = True
                    Else
                        'Nope, keep looking.
                        lineNumber = lineNumber + direction
                    End If
                End If
                If Not found Then
                    lineNumber = lineNumber + direction
                End If
            End While
            If found Then
                txtText.SelectionStart = 1
                txtText.SelectionLength = 0
                txtText.SelectionStart = GetCharacterIndexOfLine(txtText, lineNumber)
                txtText.SelectionLength = 0
                'We don't call SelectCurrentParagraph because the selection stops screenreaders (Thunder,
                'Narrator) from reading the line. Instead they read the whole document.
                'Call SelectCurrentParagraph
                Call txtText.ScrollToCaret()
            Else
                Call PlayErrorSound()
            End If
            Call txtText.Focus()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' update the current text area
    ''' </summary>
    ''' <param name="elementText"></param>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Public Sub UpdateTextarea(ByRef elementText As String, ByRef node As mshtml.IHTMLDOMNode)
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try

            Dim e As mshtml.IHTMLElement3
            Dim txt As mshtml.IHTMLTextAreaElement
            Dim i As Integer

            e = CType(node, mshtml.IHTMLElement3)
            txt = CType(node, mshtml.IHTMLTextAreaElement)
            txt.value = elementText
            Call e.setActive()

            modGlobals.gWebHost.Visible = True
            Call modGlobals.gWebHost.webMain.Focus()
            For i = 1 To 100 : System.Windows.Forms.Application.DoEvents() : Next i
            modGlobals.gWebHost.Visible = False
            Call txtText.Focus()
            txtText.SelectionLength = 0
            Call DoDelayedRefresh()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Updates the selected item with the selected option.
    ''' </summary>
    ''' <param name="selectNumber"></param>
    ''' <param name="optionNumber"></param>
    ''' <remarks></remarks>
    Public Sub UpdateSelection(ByRef selectNumber As Integer, ByRef optionNumber As Integer)
        Try

            'Debug.Print "sN:" & selectNumber & " oN:" & optionNumber
            Dim aNode As mshtml.IHTMLDOMNode ' used to iterate through the option nodes
            Dim bNode As mshtml.IHTMLDOMNode ' ditto
            Dim selectNode As mshtml.IHTMLDOMNode
            Dim htmlE As mshtml.IHTMLElement
            Dim nodeCount As Integer ' used to track which of the select's option nodes we're up to
            Dim htmlE3 As mshtml.IHTMLElement3 'used to fire DHTML event
            Dim childNodeCollection As mshtml.IHTMLDOMChildrenCollection

            staMain.Items.Item(0).Text = modI18N.GetText("Busy")
            'update the DOM with which item is selected: first clear all the items
            selectNode = CType(selects(selectNumber).node, mshtml.IHTMLDOMNode)
            childNodeCollection = CType(selectNode.childNodes, mshtml.IHTMLDOMChildrenCollection)
            'of the select at least)
            nodeCount = 0 ' start counting nodes
            For Each aNode In childNodeCollection
                If aNode.nodeName = "OPTION" Then
                    'we have an option node
                    htmlE = CType(aNode, mshtml.IHTMLElement)
                    Call htmlE.setAttribute("selected", False) ' make it not-selected
                    nodeCount = nodeCount + 1 ' increase the count
                    'If we have the newly selected node, set the selected attribute
                    If nodeCount = optionNumber Then ' really? it's 1-based array? TODO check this!
                        'this node is the selected one
                        Call htmlE.setAttribute("selected", True) ' make it selected
                    End If
                    'Debug.Print ("Just did " & nodeCount)
                ElseIf aNode.nodeName = "OPTGROUP" Then
                    'we have an option group
                    Dim grandChildNodeCollection As mshtml.IHTMLDOMChildrenCollection = CType(aNode.childNodes, mshtml.IHTMLDOMChildrenCollection)
                    For Each bNode In grandChildNodeCollection
                        If bNode.nodeName = "OPTION" Then
                            htmlE = CType(bNode, mshtml.IHTMLElement)
                            htmlE.setAttribute("selected", False) ' make it not-selected
                            nodeCount = nodeCount + 1 'increase the count
                            'If we have the newly selected node, set the selected attribute
                            If nodeCount = optionNumber Then ' Again, really? 1-based? TODO check!
                                'this node is the selected one
                                'Debug.Print bNode.attributes.getNamedItem("id").nodeValue
                                'Debug.Print bNode.attributes.getNamedItem("selected").nodeValue
                                Call htmlE.setAttribute("selected", True) ' make it selected
                                'Debug.Print bNode.attributes.getNamedItem("selected").nodeValue
                            End If
                        End If
                    Next bNode
                End If
            Next aNode

            'activate appropriate script events
            htmlE3 = CType(selects(selectNumber).node, mshtml.IHTMLElement3)
            Call htmlE3.FireEvent("onchange")
            'now update the display
            Call DoDelayedRefresh()
            staMain.Items.Item(0).Text = modI18N.GetText("Done")
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Called when WebbIE needs to go somewhere new
    ''' </summary>
    ''' <param name="url"></param>
    ''' <remarks></remarks>
    Public Sub StartNavigating(ByRef url As String)
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            If url.ToLowerInvariant().Trim = "about:blank" Then
                Call ClearPageData()
                Call txtText.Clear()
                Call txtText.Focus()
                Call StopBusyAnimation()
            Else
                tmrBusyAnimation.Enabled = True
                'Now reset the "we're only showing forms" mode
                gShowFormsOnly = False
                'Might be XML or HTML
                'finally send WebBrowser on its way, it's a normal file
                'Not going to just do this, going to count stuff.

                Call AddURLToRecent(url)
                cboAddress.Text = url ' update display
                Call PlayStartSound()
                modGlobals.gDesiredURL = url
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Call when page is loaded, including all frames
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub ProcessPage()
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try

            'Note that a page action HAS been triggered.
            gNoPageActionSoRefresh = False

            'update location bar with actual URL (inc. http://etc.)
            If modGlobals.gWebHost.URL = "" Then
            Else
                cboAddress.Text = modGlobals.gWebHost.URL
            End If

            'okay, now go through all the loaded docs and display
            Call ParseDocument()
            'Check for YouTube, and if found then autostart it.
            Call StartYoutubeIfFound()
            'Change display
            Call UpdateMenus()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Updates (enables or disables) menu items and buttons to reflect the state of the application.
    ''' Called (for example) by RefreshCurrentPage after it has run.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateMenus()
        Try
            mnuFilePrint.Enabled = True
            mnuFileSave.Enabled = True
            mnuViewCroppage.Text = modI18N.GetText("&Crop page")
            btnSkiplinks.Enabled = True
            btnHeading.Enabled = True
        Catch
        End Try
    End Sub

    Private Sub StartYoutubeIfFound()
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Try
            If modGlobals.gWebHost.URL.ToUpperInvariant.StartsWith("HTTP://WWW.YOUTUBE.COM") Or modGlobals.gWebHost.URL.ToUpperInvariant().StartsWith("HTTPS://WWW.YOUTUBE.COM") Then
                'Ah, got a YouTube link. Have to click on the video to play it.
                gDisplayingYoutube = True
                Call Application.DoEvents()
                modGlobals.gWebHost.Visible = True
                For i As Integer = 0 To 20
                    Call System.Threading.Thread.Sleep(5)
                    Call Application.DoEvents()
                Next i
                Dim pageUIA As System.Windows.Automation.AutomationElement = System.Windows.Automation.AutomationElement.FromHandle(modGlobals.gWebHost.webMain.Handle)
                Dim pc As System.Windows.Automation.PropertyCondition = New System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ClassNameProperty, "MacromediaFlashPlayerActiveX", System.Windows.Automation.PropertyConditionFlags.IgnoreCase)
                Dim playerUIA As System.Windows.Automation.AutomationElement = pageUIA.FindFirst(System.Windows.Automation.TreeScope.Descendants, pc)
                '        pageUIA.FindAll(TreeScope.Descendants,New PropertyCondition(AutomationElement.
                If playerUIA Is Nothing Then
                Else
                    Dim rect As System.Windows.Rect = playerUIA.Current.BoundingRectangle
                    Call NativeMethods.SetCursorPos(CInt(rect.Left) + 3, CInt(rect.Top) + 3)
                    Call NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, New IntPtr(0))
                    Call NativeMethods.mouse_event(NativeMethods.MOUSEEVENTF_LEFTUP, 0, 0, 0, New IntPtr(0))
                End If
                For i As Integer = 0 To 20
                    Call System.Threading.Thread.Sleep(5)
                    Call Application.DoEvents()
                Next i
                modGlobals.gWebHost.Visible = False
                gDisplayingYoutube = False
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Selects the current paragraph: call this after SkipLinks or similar
    ''' to give a visual representation of where we're up to.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SelectCurrentParagraph()
        Try
            Dim foundEnd As Integer
            Dim foundStart As Integer
            Dim startFrom As Integer

            'Get the end point
            startFrom = txtText.SelectionStart
            If startFrom = 0 Then startFrom = 1
            foundEnd = InStr(startFrom, txtText.Text, vbNewLine, CompareMethod.Binary)
            If foundEnd = 0 Then
                foundEnd = Len(txtText.Text)
            End If
            foundStart = InStrRev(txtText.Text, vbNewLine, startFrom, CompareMethod.Binary)
            If foundStart > 0 Then
                foundStart = foundStart + 1 ' Dev: otherwise we select the newline.
                'You'd think it would be +len(vbnewline), but that goes too far!
            End If
            txtText.SelectionStart = foundStart
            txtText.SelectionLength = foundEnd - foundStart
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' corrects the tabstops on the main form. Call during Form_Load
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetupTabs()
        Try
            'correct the tabs
            staMain.TabIndex = 0
            cboAddress.TabIndex = 1
            txtText.TabIndex = 2 '       alternate tabstop property
        Catch ex As Exception
        End Try
    End Sub

    'Private Sub GetPDFFromGoogle(ByRef url As String)

    '    cboAddress.Text = "http://www.google.com/search?q=cache:" & Replace(url, "http://", "")
    '    'UPGRADE_WARNING: Navigate2 was upgraded to Navigate and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    '    Call modGlobals.gWebHost.webMain.Navigate(New System.Uri(cboAddress.Text))
    'End Sub

    ''' <summary>
    ''' checks document for headings, if any, and populates headingLevel with the
    ''' tag to get for the highest-priority heading
    ''' </summary>
    ''' <param name="element"></param>
    ''' <remarks></remarks>
    Public Sub IdentifyHeadings(ByRef element As mshtml.IHTMLElement)
        Try
            Dim doc As mshtml.IHTMLDocument3

            doc = CType(element.document, mshtml.IHTMLDocument3)
            If doc.getElementsByTagName("H1").length = 0 Then
                'no H1: try for H2.
                If doc.getElementsByTagName("H2").length = 0 Then
                    'no H2: try for H3
                    If doc.getElementsByTagName("H3").length = 0 Then
                        'no H3: try for H4
                        If doc.getElementsByTagName("H4").length = 0 Then
                            'no H4: try for H5
                            If doc.getElementsByTagName("H5").length = 0 Then
                                'No H5: try for H7
                                If doc.getElementsByTagName("H6").length = 0 Then
                                    'No headings at all!
                                    mHeadingLevel = "NO HEADING ELEMENT! SDFJSAFLKJSDLFJS K"
                                Else
                                    mHeadingLevel = "H6"
                                End If
                            Else
                                mHeadingLevel = "H5"
                            End If
                        Else
                            mHeadingLevel = "H4"
                        End If
                    Else
                        'h3 found
                        mHeadingLevel = "H3"
                    End If
                Else
                    'h2 found
                    mHeadingLevel = "H2"
                End If
            Else
                'h1 found
                mHeadingLevel = "H1"
            End If
        Catch
        End Try
    End Sub

    Public Sub AddURLToRecent(Optional url As String = "")
        Try
            Static recent As String
            If url = "" Then url = cboAddress.Text
            If Len(url) > 0 Then
                url = Replace(url, "http://", "", , , CompareMethod.Text)
                url = Replace(url, "https://", "", , , CompareMethod.Text)
                If InStr(1, url, "?") > 0 Then
                    url = Microsoft.VisualBasic.Left(url, InStr(1, url, "?", CompareMethod.Binary) - 1)
                End If
                If InStr(1, recent, url, CompareMethod.Text) > 0 Then
                    'already in list
                Else
                    'add to list
                    Try
                        Call cboAddress.Items.Add(url)
                        recent = recent & "*" & url
                    Catch ex As System.ArgumentNullException
                        ' End up here if url is null: don't add to combobox.
                    End Try
                End If
            End If
        Catch
        End Try
    End Sub

    Private Function ContainsInputElement(ByRef e As mshtml.IHTMLElement) As Boolean
        Try
            If FindInputElement(e) Is Nothing Then
                Return False
            Else
                Return True
            End If
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Returns the element that is a child of e (or e itself) that is an input/form element, 
    ''' or Nothing if not found.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindInputElement(ByRef e As mshtml.IHTMLElement) As mshtml.IHTMLElement
        Try
            Dim childE As mshtml.IHTMLElement

            FindInputElement = Nothing
            Select Case UCase(e.tagName)
                Case "INPUT"
                    FindInputElement = e
                Case "DIR"
                    FindInputElement = e
                Case "BUTTON"
                    FindInputElement = e
                Case "ISINDEX"
                    FindInputElement = e
                Case "MENU"
                    FindInputElement = e
                Case "SELECT"
                    FindInputElement = e
                Case "TEXTAREA"
                    FindInputElement = e
                Case Else
                    Dim ci As mshtml.IHTMLElementCollection
                    ci = CType(e.children, mshtml.IHTMLElementCollection)
                    For Each childE In ci
                        FindInputElement = FindInputElement(childE)
                        If Not (FindInputElement Is Nothing) Then Exit For
                    Next childE
            End Select
        Catch
            Return Nothing
        End Try
    End Function

#If DEBUGGING Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUGGING did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Public Sub DummyFunction()
	Debug.Print DummyFunctionPutInToStopYouCompilingAndReleasingWithDebuggingTurnedOn
	End Sub
#End If

    Private Function GetItemIndex(ByVal txtBox As RichTextBox, ByVal tagSought As String) As Integer
        Try
            'Return 0-based index. s is the contents of the textbox to search. position is the caret position. tagSought is the line-start text.
            Dim found As Integer
            Dim lastFound As Integer
            Dim counter As Integer

            Dim position As Integer = txtBox.SelectionStart
            lastFound = 1
            found = 1
            counter = -1
            tagSought = vbLf & tagSought ' LF for RichTextBoxen.
            While found > 0 And found < position And counter < 100000
                counter = counter + 1
                lastFound = found
                found = txtBox.Text.IndexOf(tagSought, found + 1)
                'found = InStr(found + 1, s, tagSought, CompareMethod.Binary)
            End While
            'Return 0-based index.
            GetItemIndex = counter - 1
        Catch
            Return 0 ' Why not?
        End Try
    End Function

    ''' <summary>
    ''' Allow users to follow a particular link by number
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtText_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtText.KeyDown
        Try
            Dim KeyCode As Integer = e.KeyCode

            'Detect CAPSLOCK is down, and if so don't do any key handling, since it's the standard screenreader control key.
            If modKeys.IsCapslockPressed Then
                Exit Sub
            End If

            Dim handled As Boolean = False
            Dim currentLine As String
            Dim itemNumber As Integer
            Dim prompt As String
            Dim title As String
            Static busy As Boolean

            If busy Then
                Exit Sub
            End If
            busy = True
            mControlKeyPressed = e.Control
            mShiftKeyPressed = e.Shift
            'Debug.Print "MCKP " & mControlKeyPressed
            Dim linkNumber As Integer
            Dim inputString_Renamed As String
            If KeyCode = System.Windows.Forms.Keys.G And e.Control Then
                prompt = modI18N.GetText("Input link number:")
                title = modI18N.GetText("Follow link by number")
                inputString_Renamed = InputBox(prompt, title)
                If inputString_Renamed <> "" Then
                    linkNumber = CType(Val(inputString_Renamed), Integer) - 1
                    'Debug.Print linkNumber
                    If linkNumber >= 0 And linkNumber < gNumLinks Then
                        Call StartNavigating(gLinks(linkNumber).address)
                    Else
                        prompt = modI18N.GetText("No link with that number")
                        title = modI18N.GetText("Link not found")
                        MsgBox(prompt, MsgBoxStyle.Exclamation, title)
                    End If
                End If
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.F8 Then
                'pressed f8: if over link, show information

                currentLine = GetCurrentLine(txtText)
                If currentLine.Contains(ID_LINK) Then
                    'move past the word "LINK" and grab the next 4 characters (up to 9999 links are possible - array can't take it though)
                    'itemNumber = Val(Mid(currentLine, Len(ID_LINK) + 2, 4))
                    itemNumber = GetItemIndex(txtText, ID_LINK)
                    'Don't select - stops screenreaders reading line, they read the document instead.
                    'Call SelectCurrentParagraph
                    Dim details As String
                    details = GetText("Description:") & vbTab & gLinks(itemNumber).description & vbNewLine
                    details = details & GetText("Address:") & vbTab & gLinks(itemNumber).address & vbNewLine
                    details = details & GetText("Target:") & vbTab & gLinks(itemNumber).target & vbNewLine
                    details = details & GetText("FrameIndex:") & vbTab & gLinks(itemNumber).frameIndex & vbNewLine
                    details = details & GetText("Element.tagName:") & vbTab & gLinks(itemNumber).element.tagName & vbNewLine
                    details = details & GetText("Element.outerHTML:") & vbTab & gLinks(itemNumber).element.outerHTML & vbNewLine

                    Call MsgBox(details, MsgBoxStyle.Information, Application.ProductName)
                Else
                    Call PlayErrorSound()
                End If
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.Tab And e.Control Then
                'Don't be tempted to put any SendKeys in here to encourage screenreaders to read the new line.
                'This will be keypress-handled and since Ctrl is down you'll get SkipLinks happening.
                If e.Shift Then
                    Call SkipLinks(SKIP_UP, True)
                Else
                    Call SkipLinks(SKIP_DOWN, True)
                End If
                handled = True
            ElseIf e.Alt Then
                'pressed ALT and something: check for access key
                If Len(modAccesskeys.GetURL(KeyCode)) > 0 Then
                    Call StartNavigating(modAccesskeys.GetURL(KeyCode))
                    handled = True
                End If
                'ElseIf KeyCode = System.Windows.Forms.Keys.Space Then
                'OK, so we used to have "Space does PgDn like in IE". But in Thunder ScreenReader, CAPSLOCK+Space
                'is "read this window". And users have PgDn anyway. So leave it alone. November 2012.
                'Call(System.Windows.Forms.SendKeys.Send("{PGDN}"))
                'DEV: Tried to implement a zoom in and out function like in Mozilla,
                'but didn't work: won't update text box. Don't know why.
                '    ElseIf KeyCode = 187 Then ' Magic number indicating "=" character
                '        If (Shift And vbCtrlMask) > 0 Then
                '            'Zoom in.
                '            #If UNICONTROL_STATE = USE_UNICONTROLS Then
                '                txtText.Font.size = txtText.Font.size + 2
                '                Call Me.Refresh
                '            #Else
                '                txtText.Font.size = txtText.Font.size + 2
                '                Call txtText.Refresh
                '            #End If
                '        End If
                '    Else
                '        Debug.Print "KeyCode:" & KeyCode
            ElseIf KeyCode = System.Windows.Forms.Keys.D1 And e.Control Then
                'Jump to next/previous Header 1
                Call GotoLineStarting(SECTION_MARKER_H1, CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D2 And e.Control Then
                'Jump to next/previous Header 2
                Call GotoLineStarting(SECTION_MARKER_H2, CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D3 And e.Control Then
                'Jump to next/previous Header 3
                Call GotoLineStarting(SECTION_MARKER_H3, CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D4 And e.Control Then
                'Jump to next/previous Header 4
                Call GotoLineStarting(SECTION_MARKER_H4, CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D5 And e.Control Then
                'Jump to next/previous Header 5
                Call GotoLineStarting(SECTION_MARKER_H5, CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.D6 And e.Control Then
                'Jump to next/previous Header 6
                Call GotoLineStarting(SECTION_MARKER_H6, CType(IIf(e.Shift, -1, 1), Integer), False)
                handled = True
            ElseIf KeyCode = System.Windows.Forms.Keys.Return Then
                Call UserPressedReturn()
                handled = True
            End If
            If handled Then
                e.Handled = True
                e.SuppressKeyPress = True
            End If
            busy = False
        Catch
        End Try
    End Sub

    'Private Function GetItemIndex(txtBox As System.Windows.Forms.RichTextBox, id As String) As Integer
    '    'Get the index of the id on the line containing the caret. 
    '    'Assumes that the first line (0th position) in the textbox says "WebPage" - that is
    '    'it can't have the id there.
    '    Dim endOfLineWithCaret As Integer = txtBox.Text.IndexOf(vbCr, txtBox.SelectionStart)
    '    If endOfLineWithCaret = 0 Then endOfLineWithCaret = txtBox.Text.Length 'last line on page!

    '    Dim textToCountIn As String = txtBox.Text.Substring(0, endOfLineWithCaret)
    '    'OK, how many times does the id occur in this?
    '    Dim count As Integer = 0
    '    Dim foundAt As Integer = -1
    '    foundAt = textToCountIn.IndexOf(vbLf & id, 0)
    '    While foundAt > 0
    '        count = count + 1
    '        foundAt = textToCountIn.IndexOf(vbLf & id, foundAt + 1)
    '    End While
    '    Return count - 1 ' ids go from 0

    '    'This code produces mis-counts - probably the .Lines confusion.
    '    ''Dim idCount As Integer = 0
    '    ''Dim charactersGot As Integer = 0
    '    ''For lineIndex As Integer = 0 To txtBox.Lines.Length - 1
    '    ''    'Is this a line that contains id?
    '    ''    If txtBox.Lines(lineIndex).Contains(id) Then
    '    ''        'It is. Increment the id count.
    '    ''        idCount = idCount + 1
    '    ''    End If
    '    ''    'How many characters on this line, now?
    '    ''    charactersGot = charactersGot + txtBox.Lines(lineIndex).Length + vbNewLine.ToString.Length
    '    ''    'Right, have we now whizzed past the caret position?
    '    ''    If charactersGot > txtBox.SelectionStart Then
    '    ''        'Ah! yes. Stop now.
    '    ''        Exit For
    '    ''    End If
    '    ''Next lineIndex
    '    ' ''item index is 0-based, and we will have counted the one we found in our current line, so:
    '    ''idCount = idCount - 1
    '    ''Return idCount

    '    ''We are positioned at SelectionStart, which is a position 0 or more.
    '    ''We know that there is at least one id, because there is one on this line.
    '    ''It may be at the same position as the caret, or to the left of it.
    '    ''So if we start at 0, and stop:
    '    ''   When we get to an id with an index greater than SelectionStart
    '    ''   Or we run out of ids, meaning the id we found was the last one in the text.

    '    'Dim idCount As Integer = 0
    '    'Dim upTo As Integer = 0
    '    'Dim selectionStart As Integer = txtBox.SelectionStart
    '    'Dim s As String = txtBox.Text

    '    'While upTo <= selectionStart
    '    '    idCount = idCount + 1 ' idCount is therefore always at least 1, so we assume a 
    '    '    '   1-based array in effect. 

    '    '    'Find the next instance of the id:
    '    '    upTo = InStr(upTo + 1, s, id, CompareMethod.Binary)
    '    '    'upTo will either be the next id position, in which case we'll go round the loop
    '    '    'again if it is still less than selectionStart, or 0, indicating this was the last one. 
    '    '    If upTo = 0 Then
    '    '        Exit While ' OK, we're done.
    '    '    End If
    '    'End While
    '    'Return idCount
    'End Function

    ''' <summary>
    ''' This handles user interaction with the txtText area - in other words, txtText_KeyPress.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UserPressedReturn()
        Try
            Dim currentLine As String = GetCurrentLine(txtText)
            'find out what this line actually contains:
            If currentLine.Contains(ID_LINK) Then
                Call PressedReturnOnLink()
            Else
                Call PressedReturnOnNotLink(currentLine)
            End If
        Catch
        End Try
    End Sub

    Private Sub PressedReturnOnLink()
        Try
            Dim itemNumber As Integer = GetItemIndex(txtText, ID_LINK)
            'Debug.Print("Got link to " & gLinks(itemNumber).address)
            'itemNumber = GetItemIndex(txtText, ID_LINK)
            'check it isn't a javascript link
            If gLinks(itemNumber).address.ToLowerInvariant.StartsWith("javascript:") Or gLinks(itemNumber).address.EndsWith("#") Then
                'javascript, like href="javascript:GotoPage(4)".  Do a click event and refresh
                gForceNavigation = True ' 3.7.4
                Call gLinks(itemNumber).element.click()
            ElseIf gLinks(itemNumber).address.StartsWith("FAKEA-RSS") Then
                'News.
                Call RSS()
            ElseIf gLinks(itemNumber).address.ToLowerInvariant.StartsWith("mailto:") Then
                'it's a mailto link - open default mail client by clicking item
                gForceNavigation = True ' 3.7.4
                Call gLinks(itemNumber).element.click()
                Call DoDelayedRefresh()
            ElseIf gLinks(itemNumber).address.Contains("#") And Not gLinks(itemNumber).address.EndsWith("#") Then
                'Internal link! 
                Dim url As System.Uri = New System.Uri(gLinks(itemNumber).address)
                Debug.Print(url.ToString)
                'Slightly oddly, I can't get the #target part of the url from System.Uri (unless I'm being stupid)
                Dim urlWithoutTarget As String = url.Scheme & Uri.SchemeDelimiter.ToString & url.Host.ToString & url.PathAndQuery.ToString
                Dim targetURL As String = url.ToString.Replace(urlWithoutTarget, "")
                mInternalTarget = targetURL.Replace("#", "")
                mSeekingInternalTarget = True
                Call DoDelayedRefresh()
            ElseIf mForceDownloadLink Or mShiftKeyPressed Then
                'Download link target to desktop instead of opening it.
                'TODO Test this works!
                Dim targetURL As String = gLinks(itemNumber).address
                If targetURL Is Nothing Then
                    Call Beep()
                ElseIf targetURL = "" Then
                    Call Beep()
                Else
                    Try
                        Dim target As System.Uri = New System.Uri(targetURL)
                        Dim parts() As String = target.LocalPath.Split(CChar("/"))
                        Dim filename = parts(parts.Length - 1)
                        For Each c As Char In System.IO.Path.GetInvalidFileNameChars
                            filename = filename.Replace(c, "_")
                        Next c
                        filename = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\" & filename
                        Call modHTTPGetter.GetFile(gLinks(itemNumber).address, filename)
                    Catch ex As Exception
                        Call Beep()
                    End Try
                End If
            Else
                'First, get the current line before we navigate - used
                'for back and forward. Used to be in BeforeNavigate, but
                'gets too confusing with all the frames and adverts triggering
                'BeforeNavigate
                'record the line number
                Dim lineNumber As Integer = GetCurrentLineIndex(txtText)
                'we often get several  events for a page while it
                'loads (e.g. adverts) and we only want to record the line number when
                'there's actually something on the page, e.g. lineNumber > 1
                If lineNumber > 1 Then
                    If mobjNavigationRecord.ContainsKey(modGlobals.gWebHost.URL) Then
                        mobjNavigationRecord.Remove(modGlobals.gWebHost.URL)
                    End If
                    mobjNavigationRecord.Add(modGlobals.gWebHost.URL, lineNumber)
                End If
                'Now to handle the click
                'DEV: two options here, for a link of some sort:
                'Option 1: click it, assume IE starts navigating
                'Option 2: extract the url and start going there manually
                'In theory, do (1), since it supports more Javascript
                'In practice, this can fail. For example, IBM.com in October 2009 has a Javascript-based
                'set of drop-down navigation bars. They have valid hrefs but clicking does nothing.
                'Added complication: on working in development, I notice that sometimes links don't work. Return does nothing.
                'Checking it out (December 2008) I find this is because I'm getting an Error 70 Permission Denied on the
                '.element object. So let's add some code (it'll be 3.8) that handles the permission problems.
                'I've tried checking .address and just starting navigation
                'if this is > 1 (to account for href="#") but .address
                'always has the whole URL, even if the code just says "#".
                'Options:
                '0 Just use href. Clearly doesn't work on Javascript.
                '1 I could parse the html to look for href="#" or href=""
                '  or a complete lack of href, and do a .click if so.
                '2 I could check for some onclick code, and if so do an
                '  onclick: if not, do a manual click. If this fails do the href navigation.
                '3 I could do a .click on everything
                '[If I don't have the element I should always do the manual
                'method, obviously]

                'Turns out (3.11) that more stuff breaks if you don't do click. So do .click and
                'provide a new menu option to force following the address.
                'First, do I even have something to click, or is the user forcing following the address?
                'Debug.Print "Item number: " & itemNumber
                Dim doManualNavigation As Boolean = False
                If gLinks(itemNumber).element Is Nothing Or mForceFollowAddress Then
                    'have to do it manually
                    doManualNavigation = True
                ElseIf NativeMethods.GetKeyState(NativeMethods.VK_CONTROL) < 0 Then
                    'User pressing control and return, so do follow. 3.11
                    doManualNavigation = True
                Else
                    Dim link As mshtml.IHTMLAnchorElement = TryCast(gLinks(itemNumber).element, mshtml.IHTMLAnchorElement)
                    If link Is Nothing Then
                        doManualNavigation = True
                    ElseIf link.target Is Nothing Then
                        'OK, no target attribute. If there is you can't click...
                        '(Although shouldn't the New Window code catch this?
                        'Ah, but maybe that's different from _blank - which 
                        'creates a new tab. Well, works now.)
                    Else
                        'Don't support frames, so navigate by hand 
                        doManualNavigation = True
                    End If

                    If doManualNavigation Then
                        'Wait until we see if we click okay.
                    Else
                        'can click. But do we want to? IBM links (www.ibm.com October 2009) don't produce any
                        'action when you click them - they have JavaScript functions attached, applied
                        'after the page loads. So reverse the previous presumption (3.10.2-) that
                        'if script, do click. Now it's if href, do navigate, otherwise click.
                        'Dev: This >1 length check is supposed to catch href="#" and suchlike, but I think the
                        '.address value gets populated from a property that is resolved to the canonical url, so
                        'it will always be "www.test.com/index.htm#" rather than just "#" - in other words, .address>1
                        'will always be true. Unless there is no href at all:
                        'OK, we have a url, which means we're going to follow it.
                        'got an address to use: but one further complication,
                        'if this link is in a frame then you have to click
                        'it or you break the frame handling.
                        'So 3.11: always do click. See commented-out code below for alternatives, but
                        'this will break lots of sites. October 2009.
                        Try
                            Call gLinks(itemNumber).element.click()
                        Catch
                            doManualNavigation = True
                        End Try
                    End If
                    If doManualNavigation Then
                        Call StartNavigating(gLinks(itemNumber).address)
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub PressedReturnOnNotLink(ByVal currentLine As String)
        Try
            If currentLine.StartsWith(ID_SELECT) Then
                'Form input button:
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_SELECT)
                Dim selectInput As mshtml.IHTMLSelectElement = CType(selects(itemNumber).node, mshtml.IHTMLSelectElement)
                If mControlKeyPressed Then
                    'control and return = submit form.
                    gForceNavigation = True
                    Dim n As mshtml.IHTMLSelectElement = selects(itemNumber).node
                    Call n.form.submit()
                ElseIf selectInput.disabled Then
                    'Can't change, not enabled.
                    Call PlayErrorSound()
                Else
                    Call frmSelect.Populate(itemNumber)
                    Call frmSelect.ShowDialog(Me)
                    Call txtText.Focus()
                    txtText.SelectionLength = 0
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_BUTTON) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_BUTTON)
                'In 4.4 I tried this cast: but it might be an IHTMLInputElement, so the cast fails.
                'Dim button As mshtml.IHTMLButtonElement = CType(buttonInput(itemNumber), mshtml.IHTMLButtonElement) 'simulate a click of this button
                'Use HTMLElement instead.
                Dim button As mshtml.IHTMLElement = CType(buttonInput(itemNumber), mshtml.IHTMLElement) 'simulate a click of this button
                Dim doClick As Boolean = True
                Try
                    If mControlKeyPressed Then
                        'User has held down control and hit return, so click the button.
                        doClick = True
                    ElseIf CType(button.getAttribute("disabled").ToString, Boolean) Then
                        'Can't click button, it's disabled.
                        doClick = False
                    Else
                        doClick = True
                    End If
                Catch ex As Exception
                    'Can get Access Denied from button.getAttribute, so if we do click the button anyway. 
                    'But then we're doomed in any case because (experience suggests) we'll fail to click
                    'the button. 
                    'Generally we're getting Access Denied because there has been a page navigation but
                    'we have not noticed and we're trying to access a DOM that isn't there any more.
                    'Try clicking Refresh when you run into this problem and see if the button you
                    'are clicking is still actually there!
                End Try
                If doClick Then
                    gForceNavigation = True ' 3.7.4
                    Call button.click()
                    'Don't select - stops screenreaders reading line, they read the document instead.
                    'Call SelectCurrentParagraph
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_CHECKBOX) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_CHECKBOX)

                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form
                    gForceNavigation = True ' 3.7.4
                    Call checkboxInput(itemNumber).form.submit()
                Else
                    Dim checkbx As mshtml.IHTMLInputElement
                    checkbx = checkboxInput(itemNumber)
                    'We could do a .checked = not .checked...
                    'checkbx.checked = Not checkbx.checked
                    '...but better to do a click in case there's a function attached.
                    If checkbx.disabled Then
                        'Disabled, so can't do anything!
                        Call PlayErrorSound()
                    Else
                        Dim htmlE As mshtml.IHTMLElement = CType(checkbx, mshtml.IHTMLElement)
                        gForceNavigation = True ' 3.7.4
                        Call htmlE.click()
                        'refresh the page
                        Call DoDelayedRefresh()
                    End If
                End If
            ElseIf currentLine.StartsWith(ID_RADIO) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_RADIO)
                Dim radioElement As mshtml.IHTMLInputElement
                radioElement = CType(radioInput(itemNumber), mshtml.IHTMLInputElement)
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form
                    gForceNavigation = True ' 3.7.4
                    Call radioElement.form.submit()
                ElseIf radioElement.disabled Then
                    Call PlayErrorSound()
                Else
                    'okay, now we have to uncheck the other radio buttons in the form
                    For i = 0 To numRadioInputs - 1
                        If radioInput(i).getAttribute("name").ToString = radioElement.name Then
                            Call radioInput(i).setAttribute("checked", False)
                        End If
                    Next i
                    'do a click in case there's a function
                    gForceNavigation = True ' 3.7.4
                    Dim htmlE As mshtml.IHTMLElement = CType(radioElement, mshtml.IHTMLElement)
                    Call htmlE.click()
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_PASSWORD) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_PASSWORD)
                Dim htmlE As mshtml.IHTMLElement = passwordInput(itemNumber)
                Dim passwordE As mshtml.IHTMLInputElement
                passwordE = CType(htmlE, mshtml.IHTMLInputElement)
                'Debug.Print "itemNumber: " & itemNumber
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form
                    gForceNavigation = True ' 3.7.4
                    Call passwordE.form.submit()
                ElseIf passwordE.disabled Then
                    Call PlayErrorSound()
                Else
                    'check for a label
                    Dim passwordPrompt As String = modFormLabelHandler.GetDescriptiveText(htmlE)
                    If passwordPrompt.Length = 0 Then passwordPrompt = modI18N.GetText("Enter your password")
                    'Dim newPassword As String = InputBox(passwordPrompt, Application.ProductName, "") ' Assume you don't get a default
                    Dim pw As New frmPassword
                    Call pw.SetLabel(passwordPrompt)
                    If pw.ShowDialog(Me) = DialogResult.OK Then
                        'password value!
                        System.Windows.Forms.Application.DoEvents()
                        Call txtText.Focus()
                        txtText.SelectionLength = 0
                        System.Windows.Forms.Application.DoEvents()
                        Call htmlE.setAttribute("value", pw.GetPassword) 'give password to IE
                    End If
                    Call pw.Close()
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_SUBMIT) Then
                'itemNumber = Val(Mid(currentLine, Len(ID_SUBMIT) + 2, 4))
                gForceNavigation = True
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_SUBMIT)
                Dim htmlE As mshtml.IHTMLElement = submitInput(itemNumber)
                If mControlKeyPressed Then
                    'Control and Enter: force a submit of the form.
                    'A Submit might be a number of different items: an IHTMLButtonElement with the "submit" type, or
                    'an IHTMLInputElement. Can't cast between them, need to check what they are first.
                    If htmlE.tagName = "INPUT" Then
                        Dim inputE As mshtml.IHTMLInputElement = CType(htmlE, mshtml.IHTMLInputElement)
                        Call inputE.form.submit()
                    ElseIf htmlE.tagName = "BUTTON" Then
                        Dim inputE As mshtml.IHTMLButtonElement = CType(htmlE, mshtml.IHTMLButtonElement)
                        Call inputE.form.submit()
                    End If
                Else
                    'Just do a click on this element, assume that'll work.
                    Call htmlE.click()
                    'But what if the submit button has not created a page change? This happens, for
                    'example, when you're on Facebook and update your Status or something. So we
                    'have to either (1) refresh all the time or (2) check to see if there isn't a
                    'page change, and if not, do a refresh.
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_FILE) Then
                'itemNumber = Val(Mid(currentLine, Len(ID_FILE) + 2, 4))
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_FILE)
                Dim inputE As mshtml.IHTMLInputElement = CType(fileInput(itemNumber), mshtml.IHTMLInputElement)
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form
                    gForceNavigation = True ' 3.7.4
                    Call inputE.form.submit()
                ElseIf inputE.disabled Then
                    'disabled input, don't let the user use it.
                    Call PlayErrorSound()
                Else
                    gForceNavigation = True ' 3.7.4
                    Call CType(inputE, mshtml.IHTMLElement).click()
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_RESET) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_RESET)
                gForceNavigation = True ' 3.7.4
                Call CType(resetInput(itemNumber), mshtml.IHTMLElement).click()
                Call DoDelayedRefresh()
            ElseIf currentLine.StartsWith(ID_TEXTAREA) Then
                'itemNumber = Val(Mid(currentLine, Len(ID_TEXTAREA) + 2, 4))
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_TEXTAREA)
                Dim inputE As mshtml.IHTMLTextAreaElement = CType(textAreaInput(itemNumber), mshtml.IHTMLTextAreaElement)
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form
                    gForceNavigation = True ' 3.7.4
                    Call inputE.form.submit()
                ElseIf inputE.disabled Then
                    'disabled textarea, don't let the user change it.
                Else
                    'check for a label
                    frmTextareaInput.gAreaLabel = modFormLabelHandler.GetDescriptiveText(CType(inputE, mshtml.IHTMLElement))
                    Call frmTextareaInput.Populate(CType(inputE, mshtml.IHTMLDOMNode))
                    Call frmTextareaInput.ShowDialog(Me)
                    'Call CType(inputE, mshtml.IHTMLElement2).blur() Tried this to make Facebook comment fields work, but they
                    'just won't, boo. 
                    Call DoDelayedRefresh()
                End If
            ElseIf currentLine.StartsWith(ID_TEXTBOX) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_TEXTBOX)
                Dim inputE As mshtml.IHTMLInputElement = CType(textInput(itemNumber), mshtml.IHTMLInputElement)
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form.
                    gForceNavigation = True ' 3.7.4
                    Call inputE.form.submit()
                ElseIf inputE.disabled Then
                    'disabled checkbox, don't let the user change it.
                    Call PlayErrorSound()
                Else
                    Dim promptText As String = modFormLabelHandler.GetDescriptiveText(CType(inputE, mshtml.IHTMLElement))
                    'TODO Add placeholder and autofill data: see https://developers.google.com/web/updates/2015/06/checkout-faster-with-autofill
                    If promptText = "" Then promptText = GetText("Enter text input")
                    Dim newText As String = InputBox(promptText, GetText("Text input"), inputE.value)
                    'TODO differentiate between cancelled text input and blank deliberate text input.
                    If newText <> "" Then
                        inputE.value = newText
                        Call DoDelayedRefresh()
                    End If
                End If
            ElseIf currentLine.StartsWith(ID_RANGEINPUT) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_RANGEINPUT)
                Dim inputE As mshtml.IHTMLInputElement = CType(rangeInput(itemNumber), mshtml.IHTMLInputElement)
                Dim e As mshtml.IHTMLElement = CType(inputE, mshtml.IHTMLElement)
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form.
                    gForceNavigation = True ' 3.7.4
                    Call inputE.form.submit()
                ElseIf inputE.disabled Then
                    Call PlayErrorSound()
                Else
                    Dim promptText As String = modFormLabelHandler.GetDescriptiveText(CType(inputE, mshtml.IHTMLElement))
                    If promptText = "" Then promptText = ID_RANGEINPUT
                    Dim sldr As frmSlider = New frmSlider
                    Dim minSlider As Integer
                    Dim maxSlider As Integer
                    Dim valueSlider As Integer
                    Dim valString As String = ""
                    Try
                        valString = e.getAttribute("min").ToString()
                        If valString.Length = 0 Then
                            minSlider = 0 ' Default
                        Else
                            minSlider = CType(valString, Integer)
                        End If
                    Catch ex As Exception
                        minSlider = 0
                    End Try
                    valString = ""
                    Try
                        valString = e.getAttribute("max").ToString()
                        If valString.Length = 0 Then
                            maxSlider = 100
                        Else
                            maxSlider = CType(valString, Integer)
                        End If
                    Catch ex As Exception
                        maxSlider = 100
                    End Try
                    valString = ""
                    Dim defaultSliderValue As Integer = CType(Math.Round((maxSlider - minSlider) / 2, 0), Integer) + minSlider
                    Try
                        valString = e.getAttribute("value").ToString()
                        If valString = "" Then
                            valString = defaultSliderValue.ToString()
                        Else
                            valueSlider = CType(valString, Integer)
                        End If
                    Catch ex As Exception
                        valueSlider = defaultSliderValue
                    End Try
                    'Now assign
                    Try
                        If sldr.trbMain.Maximum < minSlider Then
                            sldr.trbMain.Maximum = minSlider + 1
                        End If
                        sldr.trbMain.Minimum = minSlider
                        sldr.trbMain.Maximum = maxSlider
                        If valueSlider < minSlider Then
                            sldr.trbMain.Value = minSlider
                        ElseIf valueSlider > maxSlider Then
                            sldr.trbMain.Value = maxSlider
                        Else
                            sldr.trbMain.Value = valueSlider
                        End If
                        sldr.trbMain.LargeChange = CType((sldr.trbMain.Maximum - sldr.trbMain.Minimum) / 5, Integer)
                        sldr.trbMain.SmallChange = CType(sldr.trbMain.LargeChange / 4, Integer)
                    Catch ex As Exception
                        'Oh, who knows? Use the HTML5 defaults.
                        sldr.trbMain.Minimum = 0
                        sldr.trbMain.Maximum = 100
                        sldr.trbMain.Value = 50
                        sldr.trbMain.LargeChange = 10
                        sldr.trbMain.SmallChange = 5
                    End Try
                    sldr.Title = promptText
                    If sldr.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                        Dim newValueSlider As Integer
                        Try
                            newValueSlider = CType(sldr.trbMain.Value.ToString(), Integer)
                            inputE.value = sldr.trbMain.Value.ToString()
                        Catch ex As Exception
                            'Failed to get the value out of the slider. Just prompt the user.
                            Dim newValue As String = InputBox(promptText, Application.ProductName)
                            If newValue <> "" Then
                                inputE.value = newValue
                            End If
                        End Try
                        Call DoDelayedRefresh()
                    End If
                End If
            ElseIf currentLine.StartsWith(ID_EMAILINPUT) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_EMAILINPUT)
                Dim inputE As mshtml.IHTMLInputElement = CType(emailInput(itemNumber), mshtml.IHTMLInputElement)
                If mControlKeyPressed Then
                    'User has held down control and hit return, so submit the form.
                    gForceNavigation = True ' 3.7.4
                    Call inputE.form.submit()
                ElseIf inputE.disabled Then
                    Call PlayErrorSound()
                Else
                    Dim promptText As String = modFormLabelHandler.GetDescriptiveText(CType(inputE, mshtml.IHTMLElement))
                    If promptText = "" Then promptText = GetText("Enter email address")
                    Dim newText As String = InputBox(promptText, GetText("Email input"), inputE.value).Trim()
                    'Is this a valid email address?
                    Dim failureMessage As String = ""
                    While Not IsValidEmail(newText, failureMessage) And newText <> ""
                        If failureMessage.Length > 0 Then
                            MessageBox.Show(failureMessage, GetText("Email input"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        newText = InputBox(promptText, GetText("Email input"), newText).Trim()
                        System.Windows.Forms.Application.DoEvents()
                    End While
                    'TODO differentiate between cancelled text input and blank deliberate text input.
                    If newText <> "" Then
                        inputE.value = newText
                        Call DoDelayedRefresh()
                    End If
                End If
            ElseIf currentLine.StartsWith(ID_TABLE) Then
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_TABLE)
                'TODO Add table functions back in.
                'Call frmTable.DisplayTable(tables(itemNumber))
                'Call frmTable.Show(vbModeless, Me)
            ElseIf currentLine.StartsWith(ID_VIDEO) Then
                'Video! Um, what shall I do?
                Dim mediaControlForm As frmMediaControl = New frmMediaControl
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_VIDEO)
                mediaControlForm.SetVideoElement(itemNumber)
                Call mediaControlForm.ShowDialog(Me)
            ElseIf currentLine.StartsWith(ID_AUDIO) Then
                Dim mediaControlForm As frmMediaControl = New frmMediaControl
                Dim itemNumber As Integer = GetItemIndex(txtText, ID_AUDIO)
                mediaControlForm.SetAudioElement(itemNumber)
                Call mediaControlForm.ShowDialog(Me)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Call this to turn on a delayed refresh of the display: that is, it'll use a time to wait a bit then
    ''' do a screen refresh UNLESS a page navigation has commenced in the meantime. 
    ''' </summary>
    Public Sub DoDelayedRefresh()
        tmrRefreshIfNotChange.Enabled = True
        gNoPageActionSoRefresh = True
    End Sub


    ''' <summary>
    ''' Takes a proposed email address and does some simple checks that it is valid, returning true/false
    ''' and an informative error message. Does NOT comply with the appropriate RFC. 
    ''' </summary>
    ''' <param name="emailAddress"></param>
    ''' <param name="failureMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValidEmail(emailAddress As String, ByRef failureMessage As String) As Boolean
        Try
            If emailAddress.Length = 0 Then
                failureMessage = ""
                Return False
            ElseIf emailAddress.IndexOf("@") = -1 Then
                failureMessage = modI18N.GetText("Invalid email address: you must have an @ somewhere in the address!")
                Return False
            Else
                Dim parts() As String
                Dim atSymbol As Char = "@'".ToCharArray()(0)
                parts = emailAddress.Split(atSymbol)
                If parts.Length > 2 Then
                    failureMessage = modI18N.GetText("Invalid email address: you must have only one @ in the address!")
                    Return False
                ElseIf parts(1).Length = 0 Then
                    failureMessage = modI18N.GetText("Invalid email address: there must be something after the @ symbol!")
                    Return False
                ElseIf Not parts(1).Contains(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the part after the @ symbol must have a full stop in it!")
                    Return False
                ElseIf parts(1).EndsWith(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the part after the @ symbol must not end with a full stop!")
                    Return False
                ElseIf parts(1).StartsWith(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the part after the @ symbol must not start with a full stop!")
                    Return False
                ElseIf parts(0).Length = 0 Then
                    failureMessage = modI18N.GetText("Invalid email address: there must be something before the @ symbol!")
                    Return False
                ElseIf parts(0).StartsWith(".") Then
                    failureMessage = modI18N.GetText("Invalid email address: the email address must not start with a full stop!")
                    Return False
                Else
                    ' I guess that's okay!
                    failureMessage = ""
                    Return True
                End If
            End If
        Catch
            Return True
        End Try
    End Function

    Private Sub txtText_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtText.KeyUp
    End Sub

    Private Function SanitiseLinkForDisplay(ByVal linkText As String) As String
        Try
            Dim outputText As String
            'TODO Get more stuff to use this!
            If linkText.Length > 0 Then
                outputText = Replace(linkText, "_", " ")
                outputText = Replace(outputText, "-", " ")

                'I think I'm doing this to create something useful when you get links
                'which can only be labelled with the url, like:
                'amazon.com/ref.php?kjkljklj=sdfsdf&sdljkjk=fkljoji...
                'But I can't see any benefit of this. So comment it out for now!
                'If outputText.Contains(".") Then
                'If outputText.LastIndexOf(".") > outputText.Length - 6 Then
                'outputText = outputText.Substring(0, outputText.LastIndexOf("."))
                'End If
                'End If
                Return outputText
            Else
                Return ""
            End If
        Catch
            Return linkText
        End Try
    End Function

#If MONITOR_FOCUS Then
    Private Sub tmrFocusMonitor_Tick(sender As System.Object, e As System.EventArgs) Handles tmrFocusMonitor.Tick
        Static busy As Boolean

        Debug.WriteLine("tmrFocusMonitor running")
        If busy Then
        Else
            busy = True
            If frmGoogle.Visible Or Not Me.Visible Then
                Debug.WriteLine("tmrFocusMonitor inactive")
            Else
                'If mIEVisible Then
                '    'Fine, web browser probably has the focus.
                'Else
                Debug.WriteLine("tmrFocusMonitor valid to act")
                'So, if the mControlName is "", then MSAA has broken, probably because the modGlobals.gWebHost.webMain
                'control has seized focus. We can fix this now:
#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
                If mControlName = "" And mControlType = "ControlType.Document" Then
                    Debug.WriteLine("tmrFocusMonitor ACTING")
                    Call Me.Focus()
                    Call System.Windows.Forms.Application.DoEvents()
                    Call txtText.Focus()
                End If
#End If
                'If Me.ActiveControl Is Nothing Then
                '    'Whoops! Nothing has the focus. Head to txtText.
                '    Call Me.Focus()
                '    Call System.Windows.Forms.Application.DoEvents()
                '    Call txtText.Focus()
                'ElseIf Me.ActiveControl.Name = "cboAddress" Then
                '    'OK, we're fine with being in cboAddress.
                '    'Call Me.Focus()
                '    'Call System.Windows.Forms.Application.DoEvents()
                '    'Call cboAddress.Focus()
                'Else
                '    'Otherwise we should be in txtText.
                '    Call Me.Focus()
                '    Call System.Windows.Forms.Application.DoEvents()
                '    Call txtText.Focus()
                'End If
                'End If
            End If
            busy = False
        End If
    End Sub
#End If

#If USE_UIA_FOR_FOCUS_TRACKING = "true" Then
    Private Sub OnFocusChanged(src As Object, e As System.Windows.Automation.AutomationFocusChangedEventArgs)
        On Error Resume Next
        'Cast the object to an AutomationElement.
        Dim elementFocused As System.Windows.Automation.AutomationElement
        elementFocused = CType(src, System.Windows.Automation.AutomationElement)
        mControlName = ""
        If elementFocused Is Nothing Then
            mControlName = "Unknown"
        ElseIf elementFocused.Current Is Nothing Then
            mControlName = "Unknown"
        Else
            If elementFocused.Current Is Nothing Then
                mControlName = "Unknown"
            Else
                mControlName = elementFocused.Current.ClassName
                mControlType = elementFocused.Current.ControlType.ProgrammaticName
                Debug.WriteLine("mControlType:" & mControlType)
            End If
        End If
    End Sub
#End If
    Private Sub mnuFavoritesGotofavorites_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFavoritesGotofavorites.Click
        On Error Resume Next
        Call GotoWebbIEHomePage()
    End Sub

    Private Sub GotoWebbIEHomePage()
        Try
            Dim path As String = System.IO.Path.GetTempFileName
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(path)
            Call sw.Write(modFavorites.GenerateHomepageHTML)
            Call sw.Close()
            Call StartNavigating(path)
        Catch
        End Try
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
        Try
            'open an html file on disk or internet if a url

            Dim Path As String

            cdlgOpen.Filter = modI18N.GetText("HTML file") & " (*.htm, *.html, *.mht)|*.htm;*.xhtm;*.html;*.xhtml;*.mht|" & modI18N.GetText("HTML Help files") & " (*.chm)|*.chm|" & modI18N.GetText("PDF Files") & " (*.pdf)|*.pdf|" & modI18N.GetText("All files") & " (*.*)|*.*;"
            cdlgOpen.ShowReadOnly = False
            cdlgOpen.CheckFileExists = False
            cdlgOpen.CheckPathExists = False
            Call cdlgOpen.ShowDialog()
            'if we have a valid filename, try to load it...
            Dim cmdline As String
            Dim outputFile As String
            If cdlgOpen.FileName <> "" Then
                'okay, is this a valid file?
                If System.IO.File.Exists(cdlgOpen.FileName) Then
                    'Compiled HTML help file?
                    If StrComp(Microsoft.VisualBasic.Right(cdlgOpen.FileName, 3), "chm", CompareMethod.Text) = 0 Then
                        'compiled html help
                        frmParseHTMLHelp = New frmParseHTMLHelp
                        Path = frmParseHTMLHelp.ConvertHTMLHelp((cdlgOpen.FileName))
                        If Len(Path) > 0 Then
                            Call StartNavigating(Path)
                        End If
                    ElseIf LCase(Microsoft.VisualBasic.Right(cdlgOpen.FileName, 4)) = ".pdf" Then
                        'PDF file!
                        outputFile = Replace(cdlgOpen.FileName, ".pdf", "", , , CompareMethod.Text)
                        cmdline = """" & My.Application.Info.DirectoryPath & "\pdftohtml.exe"""
                        Dim argumentLine As String = """" & cdlgOpen.FileName & """ """ & outputFile & """ -c -p -noframes"
                        Dim si As System.Diagnostics.ProcessStartInfo = New System.Diagnostics.ProcessStartInfo(cmdline, argumentLine)
                        si.CreateNoWindow = True
                        Dim makeProcess As System.Diagnostics.Process = System.Diagnostics.Process.Start(si)
                        Call makeProcess.WaitForExit()
                        outputFile = outputFile & ".html"
                        Call StartNavigating(outputFile)
                    Else
                        'html
                        Call StartNavigating((cdlgOpen.FileName))
                    End If
                Else
                    'nope: is it a url?
                    Dim newURL As System.Uri = New System.Uri(cdlgOpen.FileName.Replace(".html", "").Replace(".htm", ""))
                    If newURL.IsWellFormedOriginalString And newURL.IsAbsoluteUri Then
                        Call StartNavigating(newURL.ToString)
                    Else
                        Call PlayErrorSound()
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        On Error Resume Next
        'save the current webpage  
        Call SaveAs()
    End Sub

    Private Sub EditCut()
        Try
            'process user menu command Cut

            Dim selStart As Integer ' position of the cursor in the text to be cut from
            'only works in address bar
            If Me.ActiveControl.ToString() = cboAddress.Text And cboAddress.SelectedText <> "" Then
                Call My.Computer.Clipboard.Clear()
                Call My.Computer.Clipboard.SetText(cboAddress.SelectedText)
                selStart = cboAddress.SelectionStart
                cboAddress.Text = Microsoft.VisualBasic.Left(cboAddress.Text, cboAddress.SelectionStart) & Microsoft.VisualBasic.Right(cboAddress.Text, Len(cboAddress.Text) - cboAddress.SelectionStart - cboAddress.SelectionLength)
                cboAddress.SelectionStart = selStart
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuEditCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        On Error Resume Next
        Call EditCut()
    End Sub

    Private Sub mnuLinksSkiplinksDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLinksSkipDown.Click
        On Error Resume Next
        Call DoSkip(SKIP_DOWN)
    End Sub

    Private Sub DoSkip(SkipDirection As Integer)
        Try
            If gShowFormsOnly Then
                Call SkipToForm(SkipDirection)
            Else
                Call SkipLinks(SkipDirection)
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuNavigateGotoheadline_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNavigateGotoheadline.Click
        On Error Resume Next
        Call GotoHeadline(SKIP_DOWN)
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptionsOptions.Click
        Try
            Dim frmOptions As frmOptionsForm = New frmOptionsForm
            If frmOptions.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                'TODO is this right? Script errors = messages?
                modGlobals.gWebHost.webMain.ScriptErrorsSuppressed = Not My.Settings.AllowMessages
                'If we don't use the IE homepage, disable the ability to change it.
                mnuOptionsSetHomepage.Enabled = My.Settings.UseIEHomepage
                Call ResizeToolbar()
                Call RefreshCurrentPage()
            End If
        Catch
        End Try
    End Sub

    Private Sub mnuFavoritesAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFavoritesAdd.Click
        Try
            'adds a bookmark to the IE favorites list
            Dim name As String = modGlobals.gWebHost.webMain.DocumentTitle
            Dim url As String = modGlobals.gWebHost.URL
            Dim c As Char

            For Each c In System.IO.Path.GetInvalidFileNameChars
                name = name.Replace(c, "")
            Next c

            name = name.Trim

            If name = "" Then name = GetText("Untitled Webpage")
            Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Favorites) & "\" & name & ".url")
            Call sw.WriteLine("[DEFAULT]")
            Call sw.WriteLine("BASEURL=" & url)
            Call sw.WriteLine("[InternetShortcut]")
            Call sw.WriteLine("URL=" & url)
            Call sw.Close()
            'Add to bookmarks menu.
            Dim tsi As ToolStripItem = mnuFavorites.DropDownItems.Add(name, Nothing, AddressOf FavoriteClickHandler)
            tsi.Tag = url
        Catch
        End Try
    End Sub

    Private Sub mnuFavoritesOrganise_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFavoritesOrganise.Click
        Try
            Dim folderPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Favorites)
            Call Shell("explorer """ & folderPath & """", AppWinStyle.NormalFocus)
            'Ideally I could do something like track the folder and see if it changes and then reload bookmarks. 
        Catch
        End Try
    End Sub

    Private Sub FavoriteClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim tsi As ToolStripItem = CType(sender, System.Windows.Forms.ToolStripItem)
            Call StartNavigating(CStr(tsi.Tag))
        Catch
        End Try
    End Sub

    Public Sub LoadBookmarks(ByVal currentMenu As ToolStripMenuItem, ByVal dir As System.IO.DirectoryInfo)
        Try
            'Iterate through dir's files, adding to currentMenu.
            Dim f As System.IO.FileInfo
            Dim childDir As System.IO.DirectoryInfo
            Dim url As String
            Dim name As String
            Dim tsi As System.Windows.Forms.ToolStripItem
            Dim tsmi As System.Windows.Forms.ToolStripMenuItem
            'Do subfolders. 
            For Each childDir In dir.EnumerateDirectories
                name = ""
                name = childDir.Name
                tsmi = New System.Windows.Forms.ToolStripMenuItem(name)
                Call currentMenu.DropDownItems.Add(tsmi)
                Call LoadBookmarks(tsmi, childDir)
            Next childDir
            'Do favorites. It seems more standard to put folders first, I think? Though
            'going to your favorites doesn't think that way: it does links first. TODO.
            For Each f In dir.EnumerateFiles()
                name = ""
                name = f.Name
                url = modIniFile.GetString("InternetShortcut", "URL", "", f.FullName)
                If name <> "" And url <> "" Then
                    name = Replace(name, ".url", "", , , CompareMethod.Text)
                    tsi = currentMenu.DropDownItems.Add(name, Nothing, AddressOf FavoriteClickHandler)
                    tsi.Tag = url
                End If
                System.Windows.Forms.Application.DoEvents()
            Next f
        Catch
        End Try
    End Sub

    Private Sub SaveAs()
        On Error Resume Next
        'saves the current webpage as an html file
        Call modGlobals.gWebHost.webMain.ShowSaveAsDialog()
    End Sub


    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        Try
            'respond to user exiting WebbIE

            '    Dim EndCall As VbMsgBoxResult
            '    'if the user is dialled up on program exit
            '    If (ActiveConnection() = True And gblnModem = True) Then
            '        'prompt for disconnection
            '        EndCall = MsgBox("Would you like to disconnect your Internet connection as well as leave WebbIE?", vbYesNo, "Active Internet connection")
            '    End If
            '    'follow request as above
            '    If EndCall = vbYes Then
            '        Call HangUp
            '    End If
            'unload form
            Me.Close()
        Catch
        End Try
    End Sub

    Private Sub btnHeading_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHeading.Click
        Try
            'TODO check for pressing Shift, which should do skip up.
            Call GotoHeadline(SKIP_DOWN)
            Call txtText.Focus()
        Catch
        End Try
    End Sub

    Private Sub btnSkiplinks_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSkiplinks.Click
        Try
            Call DoSkip(SKIP_DOWN)
            'If mIEVisible Then
            'Else
            Call txtText.Focus()
            'End If
        Catch
        End Try
    End Sub

    Private Sub mnuOptionsSetHomepage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptionsSetHomepage.Click
        On Error Resume Next
        Call SetHomepage()
    End Sub

    Private Sub btnRSS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRSS.Click
        On Error Resume Next
        Call RSS()
    End Sub

    Private Sub mnuHelpManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpManual.Click
        On Error Resume Next
        Call modI18N.ShowHelp()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        On Error Resume Next
        Call DoWebSearch()
    End Sub

    Public Sub DoGotoAddressbar()
        Try
            cboAddress.Text = String.Empty
            Call cboAddress.Focus()
        Catch
        End Try
    End Sub

    Private Sub btnCrop_Click(sender As System.Object, e As System.EventArgs) Handles btnCrop.Click
        On Error Resume Next
        Call CropPage()
    End Sub

    Private Sub WebpageToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuViewIE.Click
        Try
            modGlobals.gShowingWebpage = True
            Call Me.Hide()
            Call gWebHost.Show()
        Catch
        End Try
    End Sub

    Private Sub ShowFavoritesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuFavoritesShowfavorites.Click
        On Error Resume Next
        Call ShowFavoritesWindow()
    End Sub

    Private Sub ShowFavoritesWindow()
        On Error Resume Next
        Call frmFavorites.ShowDialog(Me)
    End Sub

    Private Sub mnuFavorites_Click(sender As System.Object, e As System.EventArgs) Handles mnuFavorites.Click

    End Sub

    Private Sub frmMain_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        On Error Resume Next
        modGlobals.gClosing = True
        mFormClosing = True
    End Sub

    ''' <summary>
    ''' Starts off the "load the bookmarks and populate the bookmarks menu" process, which is really slow
    ''' because of all the menu work. So start the timer after loading so the program starts faster and
    ''' THEN loads in all the bookmark menu.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub tmrDelayLoadBookmarks_Tick(sender As System.Object, e As System.EventArgs) Handles tmrDelayLoadBookmarks.Tick
        If modGlobals.gClosing Then
            Me.tmrDelayLoadBookmarks.Enabled = False
            Exit Sub
        End If
        Try
            tmrDelayLoadBookmarks.Enabled = False
            Call LoadBookmarks(mnuFavorites, New System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Favorites)))
        Catch
        End Try
    End Sub

    Private Sub mnuOptionsColour_Click(sender As Object, e As EventArgs) Handles mnuOptionsColour.Click
        Dim fcs As frmColourSelect = New frmColourSelect
        Call fcs.SetCurrentSelection(My.Settings.ColourScheme)
        If fcs.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            ' Update global colours
            My.Settings.ColourScheme = fcs.SelectedColourScheme
            Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
        End If
    End Sub

    ''' <summary>
    ''' Used in mnuHelpTeamviewer_Click to store where TeamViewerQS.exe will be saved to (so when
    ''' it has downloaded it can be run.)
    ''' </summary>
    ''' <remarks></remarks>
    Private mTeamViewerPath As String = ""
    Private WithEvents mWebClient As System.Net.WebClient

    Private Sub mnuHelpTeamviewer_Click(sender As Object, e As EventArgs) Handles mnuHelpTeamviewer.Click
        mWebClient = New System.Net.WebClient()
        mTeamViewerPath = System.IO.Path.GetTempFileName() + ".exe"
        Call mWebClient.DownloadFileAsync(New System.Uri("http://www.webbie.org.uk/download/TeamViewerQS.exe"), mTeamViewerPath)
        '    wc.DownloadProgressChanged += wc_DownloadProgressChangedTV;
        
    End Sub

    Private Sub mWebClient_DownloadFileCompleted(sender As Object, e As System.ComponentModel.AsyncCompletedEventArgs) Handles mWebClient.DownloadFileCompleted
        Try
            staMain.Items.Item(0).Text = modI18N.GetText("Idle")
            System.Diagnostics.Process.Start(mTeamViewerPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub mWebClient_DownloadProgressChanged(sender As Object, e As Net.DownloadProgressChangedEventArgs) Handles mWebClient.DownloadProgressChanged
        Try
            staMain.Items.Item(0).Text = "TeamViewer " & e.ProgressPercentage & "%"
        Catch ex As Exception

        End Try
    End Sub
End Class