Public Class clsInternetExplorer
    ''' <summary>
    ''' Used to work out if a key is pressed or not.
    ''' </summary>
    ''' <param name="vKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto, CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall)> _
    Private Shared Function GetKeyState(ByVal vKey As Integer) As Short
    End Function
    Private Const VK_Escape As Long = &H1B ' 27

    Private WithEvents webMain As SHDocVw.InternetExplorer

    Private Delegate Sub DelegateCommandStateChangedToolStrip(Command As Integer, Enable As Boolean)
    Private Delegate Sub DelegateCommandStateChangedMenuStrip(Command As Integer, Enable As Boolean)
    Private Delegate Sub DelegateStatusTextChanged(Text As String)
    Private Delegate Sub DelegateDocumentTitleChanged(Text As String)
    Private Delegate Sub DelegateGeneric()

    Public Owner As frmMain

    Public Sub CreateIE(OwnerForm As frmMain)
        If IsNothing(Me.webMain) Then
            Me.webMain = New SHDocVw.InternetExplorer
            AddHandler Me.webMain.StatusTextChange, AddressOf webMain_StatusTextChanged
            AddHandler Me.webMain.TitleChange, AddressOf webMain_DocumentTitleChanged
            AddHandler Me.webMain.CommandStateChange, AddressOf webMain_CommandStateChanged
            AddHandler Me.webMain.BeforeNavigate2, AddressOf webMain_BeforeNavigate2
            AddHandler Me.webMain.NewWindow2, AddressOf webMain_NewWindow2
            AddHandler Me.webMain.DocumentComplete, AddressOf webMain_DocumentComplete
            AddHandler Me.webMain.NavigateError, AddressOf webMain_NavigateError
        End If
        Me.Owner = OwnerForm
    End Sub

    Private Sub webMain_NavigateError(pDisp As Object, ByRef URL As Object, ByRef Frame As Object, ByRef StatusCode As Object, ByRef Cancel As Boolean) Handles webMain.NavigateError
        'This is the topmost window and it has errored: stop navigating. 
        'todo check this is the topmost window!
        Call System.Diagnostics.Debug.Print("NavigateError: " & URL.ToString)
        '(It's okay if it's a frame or something erroring)
        If Me.Owner.InvokeRequired Then
            Call Me.Owner.Invoke(New DelegateGeneric(AddressOf ProcessAfterLoad))
        Else
            Call ProcessAfterLoad()
        End If

    End Sub

    Public Sub GoBack()
        If IsNothing(Me.webMain) Then
        Else
            Call Me.webMain.GoBack()
        End If
    End Sub

    Public ReadOnly Property HTMLSource As String
        Get
            If IsNothing(Me.webMain) Then
                Return ""
            Else
                Try
                    Dim doc3 As mshtml.IHTMLDocument3
                    doc3 = CType(Me.webMain.Document, mshtml.IHTMLDocument3)
                    Return doc3.documentElement.outerHTML
                Catch
                    Return ""
                End Try
            End If
        End Get
    End Property

    Public Sub GoForward()
        If IsNothing(Me.webMain) Then
        Else
            Call Me.webMain.GoForward()
        End If
    End Sub

    Public ReadOnly Property LocationURL As String
        Get
            If IsNothing(Me.webMain) Then
                Return ""
            Else
                Try
                    Return Me.webMain.LocationURL
                Catch ex As Exception
                    Return ""
                End Try
            End If
        End Get
    End Property

    Public Sub Navigate(ByVal url As String)
        If Not IsNothing(Me.webMain) Then
            Try
                Me.webMain.Navigate2(CType(url, Object))
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub

    Public Sub StopBrowser()
        If Not IsNothing(Me.webMain) Then
            Try
                If Me.webMain.Busy Then Call Me.webMain.Stop()
            Catch
            End Try
        End If
    End Sub

    Public Sub ShowIE()
        If IsNothing(Me.webMain) Then
        Else
            Try
                Me.webMain.Visible = True
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub

    Public ReadOnly Property Document As Object
        Get
            Try
                If IsNothing(Me.webMain) Then
                    Return Nothing
                ElseIf IsNothing(Me.webMain.Document) Then
                    Return Nothing
                Else
                    Try
                        Return webMain.Document
                    Catch
                        Throw
                    End Try
                End If
            Catch ex As System.Runtime.InteropServices.COMException
                System.Diagnostics.Debug.Print("Handled COM Exception in trying to get Document from InternetExplorer object " & ex.ErrorCode & " " & ex.Message)
                Return Nothing
            End Try
        End Get
    End Property


    Public ReadOnly Property LocationName As String
        Get
            If IsNothing(Me.webMain) Then
                Return ""
            Else
                Try
                    Return Me.webMain.LocationName
                Catch ex As Exception
                    Return ""
                End Try
            End If
        End Get
    End Property

    ''' <summary>
    ''' Update the main form status bar.
    ''' </summary>
    ''' <param name="Text"></param>
    ''' <remarks></remarks>
    Private Sub StatusTextChanged(Text As String)
        Me.Owner.staMain.Items.Item(1).Text = Text
    End Sub

    Private Sub webMain_StatusTextChanged(text As String) Handles webMain.StatusTextChange
        If Me.Owner.staMain.InvokeRequired Then
            Call Me.Owner.staMain.Invoke(New DelegateStatusTextChanged(AddressOf StatusTextChanged), {text})
        Else
            Call StatusTextChanged(text)
        End If
    End Sub

    Private Sub DocumentTitleChanged(Text As String)
        Me.Owner.Text = "WebbIE - " & Text
    End Sub

    Private Sub webMain_DocumentTitleChanged(Text As String) Handles webMain.TitleChange
        If Me.Owner.InvokeRequired Then
            Call Me.Owner.Invoke(New DelegateDocumentTitleChanged(AddressOf DocumentTitleChanged), {Text})
        Else
            Call DocumentTitleChanged(Text)
        End If
    End Sub

    Private Sub CommandStateChangedMenu(Command As Integer, Enable As Boolean)
        If Command = SHDocVw.CommandStateChangeConstants.CSC_NAVIGATEBACK Then
            Me.Owner.mnuNavigateBack.Enabled = Enable
        ElseIf Command = SHDocVw.CommandStateChangeConstants.CSC_NAVIGATEFORWARD Then
            Me.Owner.mnuNavigateForward.Enabled = Enable
        End If
    End Sub

    Private Sub CommandStateChangedToolStrip(Command As Integer, Enable As Boolean)
        If Command = SHDocVw.CommandStateChangeConstants.CSC_NAVIGATEBACK Then
            Me.Owner.btnBack.Enabled = Enable
        End If
    End Sub

    Private Sub webMain_CommandStateChanged(Command As Integer, Enable As Boolean) Handles webMain.CommandStateChange
        If Me.Owner.MainToolStrip.InvokeRequired Then
            Call Me.Owner.MainToolStrip.Invoke(New DelegateCommandStateChangedToolStrip(AddressOf CommandStateChangedToolStrip), {Command, Enable})
        Else
            Call CommandStateChangedToolStrip(Command, Enable)
        End If
        If Me.Owner.MainMenuStrip.InvokeRequired Then
            Call Me.Owner.MainMenuStrip.Invoke(New DelegateCommandStateChangedMenuStrip(AddressOf CommandStateChangedMenu), {Command, Enable})
        Else
            Call CommandStateChangedMenu(Command, Enable)
        End If
    End Sub

    Private Sub DoWebSearch()
        Call Me.Owner.DoWebSearch()
    End Sub

    Private Sub DoGotoAddressbar()
        Call Me.Owner.DoGotoAddressbar()
    End Sub

    Private Sub webMain_BeforeNavigate2(pDisp As Object, ByRef URL As Object, ByRef Flags As Object, ByRef TargetFrameName As Object, ByRef PostData As Object, ByRef Headers As Object, ByRef Cancel As Boolean) Handles webMain.BeforeNavigate2
        'This fires when navigation commences, and fires for each frame or page.
        'It does (did) lots of stuff in WebbIE 3, like handling internal links. But this is all buggy and hard, so I'm trying to fix
        'it by simplifying. 
        Dim webMainUrl As String
        If Me.webMain.LocationURL Is Nothing Then
            webMainUrl = ""
        Else
            Try
                webMainUrl = webMain.LocationURL.ToString
            Catch
                webMainUrl = ""
            End Try
        End If
        Dim newUrl As String
        If URL Is Nothing Then
            newUrl = ""
        Else
            Try
                newUrl = URL.ToString
            Catch
                newUrl = ""
            End Try
        End If

        Try
            'New navigation, don't fall back. (Not sure of the logic here: maybe I should check that the navigating item is, in fact,
            'the webBrowser and not a frame?)
            If newUrl.Contains("websearch_89789798") Then
                'Open websearch.
                Cancel = True
                If Me.Owner.InvokeRequired Then
                    Call Me.Owner.Invoke(New DelegateGeneric(AddressOf DoWebSearch))
                Else
                    Call DoWebSearch()
                End If
            ElseIf newUrl.Contains("goto_89789798") Then
                'Go to address bar.
                Cancel = True
                If Me.Owner.InvokeRequired Then
                    Call Me.Owner.Invoke(New DelegateGeneric(AddressOf DoGotoAddressbar))
                Else
                    Call DoGotoAddressbar()
                End If
            ElseIf newUrl.StartsWith("https") And Not webMainUrl.StartsWith("https") Then
                'Don't handle!
                Call System.Diagnostics.Debug.Print("Don't handle https in http pages")
            Else
                'Note that a page action HAS been triggered.
                If Me.Owner.InvokeRequired Then
                    Call Me.Owner.Invoke(New DelegateGeneric(AddressOf NavigationStarted))
                Else
                    Call NavigationStarted()
                End If
                Call System.Diagnostics.Debug.Print("Started Navigating:" & newUrl)
            End If
        Catch
        End Try
    End Sub

    Private Sub NavigationStarted()
        Call Me.Owner.NavigationStarted()
    End Sub

    Private Sub webMain_NewWindow2(ByRef ppDisp As Object, ByRef Cancel As Boolean) Handles webMain.NewWindow2
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
                ''        'frmNew.modGlobals.gWebHost.RegisterAsBrowser = True
                ''        'Set ppDisp = frmNew.modGlobals.gWebHost.object
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
                Cancel = True
                If Me.Owner.staMain.InvokeRequired Then
                    Call Me.Owner.staMain.Invoke(New DelegateGeneric(AddressOf PopupBlocked))
                Else
                    Call PopupBlocked()
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub PopupBlocked()
        Call PlayErrorSound()
        Me.Owner.staMain.Items.Item(1).Text = modI18N.GetText("Pop-up window blocked")
    End Sub

    Private Sub ProcessAfterLoad()
        Me.Owner.tmrProcessAfterLoad.Enabled = True
    End Sub

    Private Sub webMain_DocumentComplete(pDisp As Object, ByRef URL As Object) Handles webMain.DocumentComplete
        Dim url2 As String = CType(URL, String)
        Call System.Diagnostics.Debug.Print("DocumentComplete:" & url2)
        If String.Compare(url2, "http:///") = 0 Then
            'I'm so bored of handling the odd behaviour of WebBrowser. This allows me to use the IDE more effectively.
        ElseIf webMain.LocationURL = url2 Then
            If Me.Owner.InvokeRequired Then
                Call System.Diagnostics.Debug.Print("Invoking timer")
                Call Me.Owner.Invoke(New DelegateGeneric(AddressOf ProcessAfterLoad))
            Else
                Call System.Diagnostics.Debug.Print("Starting timer directly")
                Call ProcessAfterLoad()
            End If
        End If

    End Sub

End Class
