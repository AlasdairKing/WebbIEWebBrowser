''' <summary>
''' Hosts the WebBrowser object and handles detection of page navigation. 
''' </summary>
Public Class frmWeb

    ''' <summary>
    '''1 For suppressing JavaScript errors. http://support.microsoft.com/?kbid=279535.
    '''2 For detecting the end of navigation.
    ''' </summary>
    Private WithEvents mObjWind As mshtml.HTMLWindow2

    ''' <summary>
    ''' Used to trap and hide JavaScript errors. <see cref="mObjWind_onerror"/> 
    ''' </summary>
    Private mObjEvent As mshtml.CEventObj

    'Problems with detecting page finishing, again:
    'I have discovered, as of Feb 2017, that neither the DocumentComplete event fires NOT the JS event fires
    'for slashdot.org. Which sucks! Don't know what to do. 
    'But as of March 2017 this is fixed without intervention from me. I wrote a mechanism - see mIfIHaventHadAnyFinishedEventByNowJustProcessThePage
    '- but since I can't replicate the problem I can't fix it, so all commented out - but below.

    '''' <summary>
    '''' All the URLs to process. Used in working out when navigation has finished. <see cref="tmrCheckForNavigationComplete_Tick(Object, EventArgs)"/> 
    '''' </summary>
    Private ReadOnly mListOfUrlsToProcess As System.Collections.Generic.List(Of String) = New Generic.List(Of String)

    '''' <summary>
    '''' If ths time is reached and the webbrowser is no longer busy and <see cref="mListOfUrlsToProcess"/> has any
    '''' entries, go ahead and do it. 
    '''' </summary>
    Private mIfIHaventHadAnyFinishedEventByNowJustProcessThePage As System.DateTime

    ''' <summary>
    ''' If set, then when webMain fires DocumentComplete, process
    ''' the page. It's better to rely on mObjEvent, but that fails sometimes - I don't know why, and it might
    ''' be an artifact of running in the debugger.
    ''' </summary>
    Private mFallbackToDocumentComplete As Boolean = False

    Private Sub webMain_StatusTextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles webMain.StatusTextChanged
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            frmMain.staMain.Items.Item(1).Text = CType(eventSender, WebBrowser).StatusText
        Catch
            frmMain.staMain.Items.Item(1).Text = ""
        End Try
    End Sub

    Private Sub webMain_DocumentTitleChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles webMain.DocumentTitleChanged
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Dim title As String
        Try
            title = CType(eventSender, WebBrowser).DocumentTitle
            title = "WebbIE - " & title.Replace(vbNewLine, " ")
        Catch
            title = "WebbIE"
        End Try

        Try
            frmMain.Text = title
            Me.Text = title
        Catch
        End Try
    End Sub

    Private Sub webMain_CanGoBackChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles webMain.CanGoBackChanged
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            frmMain.btnBack.Enabled = webMain.CanGoBack
            frmMain.mnuNavigateBack.Enabled = webMain.CanGoBack
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' "Occurs when the WebBrowser control has navigated to a new document and has begun loading it."
    ''' This function tries to identify when a navigation starts and capture the necessary references
    ''' so that I tell get when the Javascript onLoad event fires, not just when the WebBrowser
    ''' object finishes navigating. This allows for better Javascript support.
    ''' If something goes wrong then set mFallbackToDocumentComplete to true and get out.
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks></remarks>
    Private Sub webMain_Navigated(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles webMain.Navigated
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Dim mObjDoc As mshtml.HTMLDocument

        Debug.Print("webMain_Navigated: " & eventArgs.Url.ToString())

        If modGlobals.gUseJavascriptForPageLoad = PageLoadStrategies.UseJavaScriptOnLoad Then
            'We are going to use Javascript to detect page loading, which is set up on this function.
        Else
            'We are going to use DocumentComplete.
            Return
        End If

        Try
            mObjDoc = CType(webMain.Document.DomDocument, mshtml.HTMLDocument)
        Catch
            mFallbackToDocumentComplete = True
            Exit Sub
        End Try

        Debug.Print("webMain_Navigated: " + eventArgs.Url.ToString())

        'Debug.Print("READYSTATE: " & webMain.ReadyState) Checking readystate doesn't seem to help with the
        'error below. It doesn't matter if you try webMain.ReadyState or mObjDoc.readystate, they only return
        '4 or "loading" in any case. So not much use. Although I haven't tried this out when there is a 
        'valid online webpage that DOESN'T throw the error, so that might be worth checking out.
        Try
            mObjWind = CType(mObjDoc.parentWindow, mshtml.HTMLWindow2)
            'OK, that worked, we'll use it to detect end of navigation.
            Debug.Print("Success in webMain_Navigated. 101. " & mObjDoc.GetType.ToString)
            'No need to fall back, check we aren't globally doing so. 
            'No, don't fall back, got object okay. 
            mFallbackToDocumentComplete = False
#If DEBUG Then
            Dim ts As System.IO.StreamWriter = New System.IO.StreamWriter(Environment.CurrentDirectory & "\error.txt", True, System.Text.Encoding.UTF8)
            Call ts.WriteLine("100 OK to COM navigate: " & webMain.Document.Url.ToString() & " " & mObjDoc.GetType.ToString)
            Call ts.Close()
#End If
        Catch comEx As System.Runtime.InteropServices.COMException
            'Failed to navigate via COM events. We could use DocumentComplete and timer. 
            'Call MessageBox.Show("Error!!! " & webMain.Url.ToString)
            Debug.Print("Exception: 10000001 modGlobals.gWebHost.webMain_Navigated COMException: " & comEx.Message & " " & mObjDoc.GetType.ToString)
            'However, it appears that if you wait until you get a susequent event, all is well. So just try later.
            'Really? Having some problems with that.  Try falling back to DocumentComplete.
            mFallbackToDocumentComplete = True
#If DEBUG Then
            Dim ts As System.IO.StreamWriter = New System.IO.StreamWriter(Environment.CurrentDirectory & "\error.txt", True, System.Text.Encoding.UTF8)
            Call ts.WriteLine("101 NOT OK to COM navigate: " & webMain.Document.Url.ToString())
            Call ts.Close()
#End If
        Catch
            Debug.Print("Unhandled Exception trying to initialise JS load detection mechanism.")
            mFallbackToDocumentComplete = False
        End Try
    End Sub

    Private Sub mObjWind_onerror(description As String, url As String, line As Integer) Handles mObjWind.onerror
        If modGlobals.gClosing Then
            Exit Sub
        End If
        'Trap and prevent script errors. See http://support.microsoft.com/?kbid=279535.
        Try
            mObjEvent = CType(mObjWind.event, mshtml.CEventObj)
            mObjEvent.returnValue = True
        Catch ex As Exception
            'Fail silently. This isn't the only mechanism, and I don't have anything useful
            'to tell the user if it fails.
            Debug.Print("mObjWind_onerror failed.")
        End Try
    End Sub

    ''' <summary>
    ''' The first of two mechanisms for detecting the end of navigation, meaning we should call <see>frmMain.ProcessAfterLoad()</see>. Based on 
    ''' the JavaScript onload event. 
    ''' </summary>
    Private Sub mObjWind_onload() Handles mObjWind.onload
        'Call MessageBox.Show("mObjWind_onload")
        If modGlobals.gClosing Then
            Exit Sub
        End If
        If gUseJavascriptForPageLoad = PageLoadStrategies.UseDocumentCompleteOnly Then
            ' Don't care if the event fired, we're always going to 
            ' use DocumentComplete.
            Call System.Diagnostics.Debug.Print("Aborted mObjWnd_onload")
            Return
        End If


        If webMain.Url.ToString.ToUpperInvariant() <> "HTTP:///" Then
            Call System.Diagnostics.Debug.Print("D! Did mObjWnd_onload")
            '' ASSUME that the JavaScript method correctly identified the end of page navigation and
            ''we are done with all the pages:
            'Call mListOfUrlsToProcess.Clear()
            ' Now process the page.
            Call frmMain.ProcessAfterLoad()
        Else
            Call System.Diagnostics.Debug.Print("S! Skipped mObjWnd_onload")
        End If
    End Sub

    ''' <summary>
    ''' This fires when navigation commences, and fires for each frame or page.
    ''' https://msdn.microsoft.com/en-us/library/system.windows.forms.webbrowser.navigating(v=vs.110).aspx
    ''' </summary>
    ''' <param name="eventSender"></param>
    ''' <param name="eventArgs"></param>
    ''' <remarks>It does (did) lots of stuff in WebbIE 3, like handling internal links. But this is all buggy 
    ''' and hard, so I'm trying to fix it by simplifying. </remarks>
    Private Sub webMain_Navigating(ByVal eventSender As System.Object,
                                   ByVal eventArgs As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles webMain.Navigating
        If modGlobals.gClosing Then
            Exit Sub
        End If

        '' In eight seconds, if we haven't heard anything back, process this page if the browser has finished. 
        'mIfIHaventHadAnyFinishedEventByNowJustProcessThePage = Now() + New TimeSpan(0, 0, 8)

        Dim webMainUrl As String
        If Me.webMain.Url Is Nothing Then
            ' webMainUrl is its default value, ""
        Else
            Try
                webMainUrl = webMain.Url.ToString()
            Catch
                ' webMainUrl is its default value, ""
            End Try
        End If
        Dim eventArgsUrl As String
        If eventArgs.Url Is Nothing Then
            eventArgsUrl = ""
        Else
            Try
                eventArgsUrl = eventArgs.Url.ToString()
            Catch
                eventArgsUrl = ""
            End Try
        End If

        Try
            'New navigation, don't fall back. (Not sure of the logic here: maybe I should check that the navigating item is, in fact,
            'the webBrowser and not a frame?)
            'If gUseJavascriptForPageLoad Then
            'mFallbackToDocumentComplete = False
            'Else
            '   mFallbackToDocumentComplete = True ' Don't care if the event fired, we're always going to 
            ' use DocumentComplete.
            'End If

            If eventArgsUrl.Contains("websearch_89789798") Then
                'Open websearch.
                eventArgs.Cancel = True
                Call frmMain.DoWebSearch()
            ElseIf eventArgsUrl.Contains("goto_89789798") Then
                'Go to address bar.
                eventArgs.Cancel = True
                Call frmMain.DoGotoAddressbar()
                'ElseIf eventArgsUrl.ToUpperInvariant.StartsWith("HTTPS") And Not webMainUrl.ToUpperInvariant.StartsWith("HTTPS") Then
                'Don't handle HTTPS in HTTP pages! 
                'Why Not? It is a problem when you have HTTP in HTTPS pages, but why the reverse?
                'Don't know, so remove this check because simpler, and I have - as always - problems with 
                'pages not loading.
                'Call System.Diagnostics.Debug.Print("Don't handle https in http pages")
            ElseIf eventArgsUrl.ToUpperInvariant().StartsWith("JAVASCRIPT:") Then
                'Don't handle! You get these, for example, in www.youtube.com (23 Feb 2014)
                Call System.Diagnostics.Debug.Print("Don't handle Javascript URLs")
            ElseIf eventArgsUrl.ToUpperInvariant().StartsWith("ABOUT:BLANK") Then
                'Don't handle. You get these, for example, in www.google.co.uk when you hit the "screenreader users,
                'turn off instant" link. 23 Feb 2014.
                Call System.Diagnostics.Debug.Print("Don't handle about:blank")
            Else
                'Now reset the "we're only showing forms" mode
                gShowFormsOnly = False

                'Dev: you'll see at this point that I tried to be clever and use pDisp as the navigating object
                'to catch navigation. This led, eventually, to an enormous memory leak and a pretty non-functional
                'WebbIE. So fall back to relying on ReadyState, and worry about it breaking another time.

                frmMain.btnStop.Enabled = True 'enable toolbar buttons

                'Don't start the animation here: starts randomly every so often.
                'frmMain.tmrBusyAnimation.Enabled = True 'start animation

                'Same here. Because random stuff happens. Probably AJAX.
                'frmMain.staMain.Items.Item(0).Text = modI18N.GetText("Downloading")
                'TODO Move this stuff until AFTER a successful navigation, because you get loads of
                'crap navigation event starts, like iframes for adverts loading.
                frmMain.mnuFavoritesAdd.Enabled = True 'enable adding a bookmark option
                Call frmMain.AddURLToRecent()
                frmMain.mnuLinksViewlinks.Enabled = False 'allow the links to be viewed in a separate form
                'Reset one-time "Must navigate" flag. 3.7.4
                gForceNavigation = False
                Call System.Diagnostics.Debug.Print("Started Navigating:" & eventArgsUrl)

                'Remember that we've decided that this URL should be processed.
                Call mListOfUrlsToProcess.Add(eventArgsUrl)
            End If
        Catch

        End Try
    End Sub

    Private Sub webMain_NewWindow(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles webMain.NewWindow
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            'This is intended to ask the user if they want to allow a pop-up Window.  However, pop-up
            'windows are almost certainly advertisements.  So, there is a user setting "mnuAllowpopups",
            'which defaults to unchecked.  If this is the case, a pop-up window will be suppressed.
            'If it is true, then a popup is allowed but the user is asked to okay it.
            'See http://support.microsoft.com/kb/q184876/ for more information.
            If My.Settings.AllowPopupWindows Then
                ''        'Let's do a new frmMain and pass this to it. Again, see http://support.microsoft.com/kb/184876.
                ''        'Set frmNew = New frmMain
                ''        'frmNew.Visible = True
                ''        'frmNew.modGlobals.gWebHost.webMain.RegisterAsBrowser = True
                ''        'Set ppDisp = frmNew.modGlobals.gWebHost.webMain.object
                'This is a lovely idea, but WebbIE is completely not designed for multiple windows. Each window
                'uses the same array of links etc. So different windows are actually representing the same page.
                'Lots of work to make pop-ups work properly.
                'Better idea: just shell WebbIE with url. May 2007
                'Call Shell(App.Path & "\" & App.EXEName & " " & ppDisp.LocationURL, vbNormalFocus)
                'Ah, this is also a great idea. But you can't get the location url from the ppDisp object because
                'it is Nothing when this event is called.
                'So just let the popup happen. 3.12.2.
            Else
                'popups are not permitted by the user settings
                Call PlayErrorSound()
                'Call MessageBox.Show(modGlobals.gWebHost.webMain.Url.ToString)
                frmMain.staMain.Items.Item(1).Text = modI18N.GetText("Pop-up window blocked")
                eventArgs.Cancel = True
            End If
        Catch
        End Try
    End Sub

    Private Sub webMain_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles webMain.DocumentCompleted
        If modGlobals.gClosing Then
            Exit Sub
        End If

        Dim url As String
        Try
            url = e.Url.ToString()
        Catch
            url = ""
        End Try
        'replacement for using the change in the status bar to indicate download status...!

        'We have got something back, wait another 8 seconds before last-ditch fallback. See mListOfUrlsToProcess.
        mIfIHaventHadAnyFinishedEventByNowJustProcessThePage = Now() + New TimeSpan(0, 0, 8)

        Try
            'I'm so bored of handling the odd behaviour of WebBrowser. This allows me to use the IDE more effectively.
            If String.Compare(url, "http:///", True) = 0 Then
                Exit Sub
            End If

            If webMain.ReadyState = WebBrowserReadyState.Complete Then
                Call Debug.Print("DocumentCompleted - ReadyState= Complete for " & url)
            Else
                Call Debug.Print("DocumentCompleted - ReadyState= NOT Complete for " & url)
            End If
            ''OK, we've handled this particular URL. Or at least - careful - we've had a DocumentComplete event for this
            ''particular URL, and our logic will run below, so we have the change to process it here or somewhere else. 
            ''The important thing is that we have had a chance to process it: if this event - or our other methods - 
            ''do not fire, then we do NOT have a chance, and we have to notice that we never handled this url somewhere
            ''else and decide what to do with it. This is all to do with detecting the end of navigation and slashdot.org
            ''not firing DocumentComplete in Feb 2017. 
            If mListOfUrlsToProcess.Contains(url) Then
                mListOfUrlsToProcess.Remove(url)
            End If
            If mFallbackToDocumentComplete And webMain.ReadyState = WebBrowserReadyState.Complete Then
                System.Diagnostics.Debug.Print("Doing DocumentCompleted fallback " + webMain.Url.ToString() + " " + e.Url.ToString())
                If webMain.Url.ToString() = url Then
                    Call frmMain.ProcessAfterLoad()
                Else
                    Call Debug.Print("DocumentCompleted = did not navigate: " + webMain.Url.ToString() + " " + url)
                End If
            Else
                Call Debug.Print("DocumentCompleted = relying on JS OR not readystate complete yet: " + webMain.Url.ToString())
                Call Debug.Print("- URL of new page is: " + url)
                If webMain.Url.ToString() <> url Then
                    If webMain.ReadyState = WebBrowserReadyState.Complete Then
                        Debug.Print("Subsequent IFRAME!")
                        Call frmMain.ProcessAfterLoad()
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.Print("Exception in DocumentCompleted: " & ex.Message)
        End Try

    End Sub

    Private Sub webMain_CanGoForwardChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles webMain.CanGoForwardChanged
        On Error Resume Next
        If modGlobals.gClosing Then
        Else
            frmMain.mnuNavigateForward.Enabled = webMain.CanGoForward
        End If
    End Sub

    Private Sub frmWeb_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        If modGlobals.gClosing Then
        ElseIf modGlobals.gShowingWebpage Then
            'OK, user requested this.
        Else
            'Happens on www.YouTube.com!
            modGlobals.gFrmWebHasFocusAndItShouldNot = True
            frmMain.tmrSetFocus.Enabled = True
        End If
    End Sub

    Private Sub frmWeb_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        webMain.ScriptErrorsSuppressed = True ' in case we create a new instance
        '   and forget to set it in the IDE!
    End Sub

    Public Sub New()
        Try
            Call WebBrowserUtility.SetWebBrowserEmulation()
        Catch
        End Try
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Try
            Call SwitchBackToMain()
        Catch
        End Try
    End Sub

    Private Sub frmWeb_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If modGlobals.gClosing Then
            Exit Sub
        End If
        Try
            If e.CloseReason = CloseReason.UserClosing Then
                e.Cancel = True
                Call SwitchBackToMain()
            End If
        Catch
        End Try
    End Sub

    Private Sub SwitchBackToMain()
        Try
            Call Me.Hide()
            Call frmMain.DoDelayedRefresh()
            frmMain.Visible = True
            modGlobals.gShowingWebpage = False
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' This is a fall-back because some web pages - e.g. YouTube, so probably pages with Flash - will
    ''' grab the key input and won't let the form get the keypress. So use the Windows API to check for
    ''' the key and close the window if Escape is pressed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub tmrCheckForEscape_Tick(sender As System.Object, e As System.EventArgs) Handles tmrCheckForEscape.Tick
        Try
            If modGlobals.gClosing Then
                tmrCheckForEscape.Enabled = False
            ElseIf Me.Visible Then
                If NativeMethods.GetKeyState(NativeMethods.VK_Escape) < 0 Then
                    Call SwitchBackToMain()
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub tmrCheckForClosing_Tick(sender As Object, e As EventArgs) Handles tmrCheckForClosing.Tick
        On Error Resume Next
        If modGlobals.gClosing Then
            tmrCheckForClosing.Enabled = False
            Call Me.Close()
        End If
    End Sub

    Private Sub tmrCheckForNavigating_Tick(sender As Object, e As EventArgs) Handles tmrCheckForNavigating.Tick
        If modGlobals.gDesiredURL <> "" Then
            Call Me.webMain.Navigate(modGlobals.gDesiredURL)
            modGlobals.gDesiredURL = ""
        End If
    End Sub

    ''' <summary>
    ''' If the browser has finished navigating, but we have urls we haven't processed, then fire off processing.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub tmrCheckForNavigationComplete_Tick(sender As Object, e As EventArgs) Handles tmrCheckForNavigationComplete.Tick
        If mListOfUrlsToProcess.Count > 0 Then
            If Now > mIfIHaventHadAnyFinishedEventByNowJustProcessThePage Then
                If Not webMain.IsBusy Then
                    Call mListOfUrlsToProcess.Clear()
                    Call frmMain.ProcessAfterLoad()
                End If
            End If
        End If
    End Sub

    Public Function URL() As String
        If webMain.Document Is Nothing Then
            Return ""
        ElseIf webMain.Document.Url Is Nothing Then
            Return ""
        Else
            Return webMain.Document.Url.ToString()
        End If
    End Function
End Class