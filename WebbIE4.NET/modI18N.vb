Option Strict On
Option Explicit On
Module modI18N

    'modI18N 
    'Code for I18N the WebbIE programs.
    '19 Nov 2012        First version.
    '27 Mar 2013        Added DoEvents to form I18N code to improve feel.

    ''' <summary>
    ''' Internationalisation for VB.Net projects using XML files
    ''' </summary>
    ''' <remarks>
    ''' HOW TO USE
    ''' 1 Create ASSEMBLYNAME.Language.Xml (e.g. WebbIE4.Language.Xml) file in same folder as program.
    ''' 2 Copy in Common.Language.Xml to same folder.
    ''' 3 Amend as required.
    ''' 4 When debugging call SetDebug() to output a list of incomplete translations to the desktop.
    ''' </remarks>



    Private _applicationXML As System.Xml.XmlDocument
    Private _languageCode As String
    Private _initialised As Boolean
    Private _applicationXMLFound As Boolean = False
    Private _commonXML As System.Xml.XmlDocument
    Private _commonXMLFound As Boolean = False
    Private _debug As Boolean = False

    'TODO cleanup.
    'So for general string use we need a simple "GetText()" function that returns either (1) the translation or
    '(2) the original, which should be English so the code is readable.
    '   Is it in Application translations?
    '   If not, is it in Common translations?
    '   If not found, return original string.
    'For translating forms we need to do each control. 
    'FIRST for the .Text property:
    '   1 Is it one of the controls where we don't translate the .Text property? If yes, stop.
    '   2 Does it have a .Text property? Yes:
    '       3 If it has a tag, try tag.Text in the Application translations
    '       4 If this is not found, or there is no tag, try formName.controlName.Text in the Application translations
    '       5 If this is not found, then try controlName.Text in Common translations
    '       6 If this is not found, then leave the control unchanged (should stay in English)
    'SECOND for the .AccessibleName property:
    '   1 Does it have an .AccessibleName property? If no, stop.
    '   2 Is the .AccessibleName property ""? If yes, stop.
    '   3 Try formName.controlName.AccessibleName in the Application translations
    '   4 If this is not found, then leave the control unchanged. 

    'Use cases
    '   A simple string.
    '   A string containing quotation marks (get converted to single quotation marks)
    '   A string containing an ampersand.
    '   A control that has a .Text property but is not converted (ComboBox)
    '   A control that is set to not convert (WebBrowser) 
    '   A control that has an AccessibleName in addition to a .Text property and not being converted.
    '   A control that has a tag.
    '   A control that is not in the application language file but the common language file.


    Public Sub SetDebug()
        On Error Resume Next
        _initialised = False
        Call Initialise()
        _languageCode = "xx"
        _debug = True
    End Sub

    Private Sub WriteDebug(s As String)
        On Error Resume Next
        Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\I18N Log.log", True, System.Text.Encoding.UTF8)
        Call sw.WriteLine(s)
        Call sw.Close()
        Application.DoEvents()
    End Sub

    Private Sub Initialise()
        Try
            If _initialised Then
                'Already loaded.
            Else
                'Load language-specific file.
                _applicationXML = New System.Xml.XmlDocument()
                Dim applicationLanguagePath As String = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".Language.xml"
                If System.IO.File.Exists(applicationLanguagePath) Then
                    Try
                        Call _applicationXML.Load(applicationLanguagePath)
                        _applicationXMLFound = True
                    Catch
                        Throw New Exception("Error loading the application language XML file, which was determined to be """ & applicationLanguagePath & """ The error returned was """ & Err.Description & """")
                    End Try
                Else
                    'No language file!
                    _applicationXMLFound = False
                End If

                'Load common file.
                _commonXML = New System.Xml.XmlDocument()
                Dim commonLanguagePath As String = My.Application.Info.DirectoryPath & "\Common.Language.xml"
                If System.IO.File.Exists(commonLanguagePath) Then
                    Try
                        Call _commonXML.Load(commonLanguagePath)
                        _commonXMLFound = True
                    Catch
                        Throw New Exception("Error loading the application language XML file, which was determined to be """ & applicationLanguagePath & """ The error returned was """ & Err.Description & """")
                    End Try
                Else
                    'No common language file!
                    _commonXMLFound = False
                End If

                'Load the locale information.
                Try
                    Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo(System.Globalization.CultureInfo.CurrentUICulture.LCID) ' Thread.CurrentThread.CurrentCulture.LCID);
                    _languageCode = ci.TwoLetterISOLanguageName.ToLowerInvariant() 'en or pt or fr or whatever.
                Catch
                    _languageCode = "en"
                End Try
                _initialised = True
            End If
        Catch
            _initialised = True
            _languageCode = "en"
        End Try
    End Sub

    Public Function GetLanguage() As String
        Try
            Call Initialise()
            Return _languageCode
        Catch
            Return "en"
        End Try
    End Function

    Public Sub ShowHelp()
        Try
            Dim path As String
            path = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".Help-" & _languageCode & ".rtf"
            If Not System.IO.File.Exists(path) Then
                If _debug Then Call WriteDebug("Did not find help file: " & path)
                path = My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".Help-en.rtf"
            End If
            If System.IO.File.Exists(path) Then
                Call Shell("write """ & path & """", AppWinStyle.NormalFocus)
            End If
        Catch
        End Try
    End Sub

    Public Sub DoForm(f As Windows.Forms.Form)
        Try
            Call Initialise()
            f.Text = GetText(f.Name & ".Text")
            For Each c As Windows.Forms.Control In f.Controls
                Call DoControl(c, f.Name)
                System.Windows.Forms.Application.DoEvents()
            Next c
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' DoControl translates a control AND its children - usually the .Text value,
    ''' but the .AccessibleName if there is one.
    ''' </summary>
    ''' <param name="c"></param>
    ''' <param name="formName"></param>
    ''' <remarks></remarks>
    Private Sub DoControl(c As Windows.Forms.Control, formName As String)
        Try
            'Handle .Text property
            Call DoControlText(c, formName)

            'Handle .AccessibleName property
            Call DoControlAccessibleName(c, formName)

            ' Handle ToolStrip control - includes menus like the File menu. 
            If TypeOf (c) Is System.Windows.Forms.ToolStrip Then
                'Tool strip. Need to process items.
                For Each ti As System.Windows.Forms.ToolStripItem In CType(c, System.Windows.Forms.ToolStrip).Items
                    Call DoControlItem(ti, formName)
                Next ti
            End If

            'This child iteration handles things like Panel controls hosting TextBox 
            'controls. It doesn't handle things like ToolstripMenuItem controls in MenuStrips. See the 
            'code above for that.
            If c.HasChildren Then
                For Each cChild As System.Windows.Forms.Control In c.Controls
                    Call DoControl(cChild, formName)
                Next cChild
            End If
        Catch
        End Try
    End Sub

    Private Sub DoControlText(c As Windows.Forms.Control, formName As String)
        Try
            Dim doText As Boolean
            Dim typeName As String = c.GetType().ToString()
            'Some controls we do not look up: notably TextBox
            If typeName = "System.Windows.Forms.TextBox" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.RichTextBox" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.ComboBox" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.ListBox" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.MenuStrip" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.WebBrowser" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.ToolStrip" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.Panel" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.PictureBox" Then
                doText = False
            ElseIf typeName = "System.Windows.Forms.StatusBar" Then
                doText = False
            ElseIf c.Text Is Nothing Then
                doText = False
            Else
                doText = True
            End If
            If doText Then
                Dim useTag As Boolean
                If c.Tag Is Nothing Then
                    useTag = False
                ElseIf c.Tag.ToString = "" Then
                    useTag = False
                Else
                    useTag = True
                End If
                Dim key As String
                If useTag Then
                    key = c.Tag.ToString()
                Else
                    key = formName & "." & c.Name
                End If

                Dim textKey As String = key & ".Text"
                Dim text As String = GetText(textKey)
                If text <> textKey Then
                    'Found an entry.
                    c.Text = text
                Else
                    'Failed to find anything when using the tag or the full formName.controlName
                    'Try falling back now to our Common file, which should save me having to 
                    'duplicate lots and lots of entries.
                    textKey = c.Name & ".Text"
                    text = GetText(textKey)
                    If text <> textKey Then
                        c.Text = text
                    End If
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub DoControlAccessibleName(c As Windows.Forms.Control, formName As String)
        Try
            Dim doAccessibleName As Boolean = False
            If c.AccessibleName Is Nothing Then
            ElseIf c.AccessibleName.Length = 0 Then
            Else
                Dim key As String = formName & "." & c.Name
                Dim accKey As String = key & ".AccessibleName"
                Dim accName As String = GetText(accKey)
                If accName <> accKey Then
                    c.AccessibleName = accName
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub DoControlItem(mi As System.Windows.Forms.ToolStripItem, formName As String)
        Try
            'Notice no .tag support. That's because in WebbIE there are lots of bookmark menu items
            'that have url tags. Don't want to I18N them. So skip tags for menus for now.
            Dim key As String
            Dim doText As Boolean = True
            If mi.Name.Length = 0 Then
                'No name: like, a menu element that is a divider, or empty - like the favorites in 
                'the WebbIE favorites menu.
                key = ""
                doText = False
            Else
                key = formName & "." & mi.Name
            End If
            If doText Then
                Dim textKey As String = key & ".Text"
                Dim text As String = GetText(textKey)
                If text <> textKey Then
                    mi.Text = text
                Else
                    'Failed to find anything when using the tag or the full formName.controlName
                    'Try falling back now to our Common file, which should save me having to 
                    'duplicate lots and lots of entries.
                    textKey = mi.Name & ".Text"
                    text = GetText(textKey)
                    If text <> textKey Then
                        mi.Text = text
                    End If
                End If
            End If
            'Now do drop-down items.
            If TypeOf (mi) Is Windows.Forms.ToolStripMenuItem And Not (TypeOf (mi) Is Windows.Forms.ToolStripSeparator) Then
                For Each miChild As Object In CType(mi, Windows.Forms.ToolStripMenuItem).DropDownItems
                    If TypeOf (miChild) Is System.Windows.Forms.ToolStripSeparator Then
                    Else
                        Call DoControlItem(CType(miChild, Windows.Forms.ToolStripMenuItem), formName)
                    End If
                Next miChild
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Returns the internationalised version of the string provided according to the current
    ''' language (and the availability of the translation.)
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If there is no "item" node in AssemblyName.Language.xml that has a "key" node
    ''' containing the argument text then the argument text is returned. This means that
    ''' if the code calls GetText("Hello world") and no translation is provided then then
    ''' function will return "Hello world". On the assumption that calling code is English
    ''' there is therefore an implicit English default.
    ''' If the "item" node that matches the argument text has a "leaveBlank" child then
    ''' the empty string is returned.
    ''' </remarks>
    Public Function GetText(ByRef text As String) As String
        Try
            Call Initialise()
            Dim key As String
            key = text.Replace("""", "'")
            Dim n As System.Xml.XmlNode = Nothing
            If _applicationXMLFound Then
                n = _applicationXML.DocumentElement.SelectSingleNode("contents/item[key=""" & key & """]")
            End If
            If n Is Nothing Then
                'Nothing found in the application translation file: try our common language file.
                If _commonXMLFound Then
                    n = _commonXML.DocumentElement.SelectSingleNode("item[key=""" & key & """]")
                End If
                If n Is Nothing Then
                    If _debug Then
                        Call WriteDebug("<item><key>" & key & "</key><content language=""en"">" & text & "</content><content language=""xx"">Xxxxx</content></item>")
                    End If
                    Return text
                Else
                    Dim item As System.Xml.XmlNode = n.SelectSingleNode("content[@language=""" & _languageCode & """]")
                    If item Is Nothing Then
                        Return text
                    Else
                        Return item.InnerText
                    End If
                End If
            Else
                If n.SelectSingleNode("leaveBlank") Is Nothing Then
                    Dim item As System.Xml.XmlNode = n.SelectSingleNode("content[@language=""" & _languageCode & """]")
                    If item Is Nothing Then
                        Return text
                    Else
                        Return item.InnerText
                    End If
                Else
                    Return ""
                End If
            End If
        Catch
            Return text
        End Try
    End Function
End Module