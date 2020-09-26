Option Explicit On
Option Strict On

' FREE code from CODECENTRIX
' http://www.codecentrix.com/
' http://codecentrix.blogspot.com/
' Collected from here Feb 2010:
' http://codecentrix.blogspot.com/2007/10/when-ihtmlwindow2getdocument-returns.html
Imports System
Imports System.Runtime.InteropServices

Friend Class CrossFrameIE

    Public Function GetIAccessibleFromNode(node As mshtml.IHTMLDOMNode) As Accessibility.IAccessible
        If node Is Nothing Then
            Return Nothing
        End If
        ' Convert IHTMLDOMNode to IAccessible using IServiceProvider.
        Try
            Dim sp As IServiceProvider = CType(node, IServiceProvider)

            ' Use IServiceProvider.QueryService to get IAccessible object.
            Dim brws As Object = Nothing
            Call sp.QueryService(IID_IAccessible, IID_IAccessible, brws)

            Return CType(brws, Accessibility.IAccessible)
        Catch ex As Exception
            Call Debug.Print("Exception in trying to get IFRAME: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Get the COM HTML Document from the IHTMLWindow2 you've got. 
    ''' </summary>
    ''' <param name="htmlWindow"></param>
    ''' <returns>Nothing, or the mshtml.IHTMLDocument2 object.</returns>
    Public Function GetDocumentFromWindow(htmlWindow As mshtml.IHTMLWindow2) As mshtml.IHTMLDocument2
        If htmlWindow Is Nothing Then
            Return Nothing
        End If
        Dim doc As mshtml.IHTMLDocument2

        If IsNothing(htmlWindow) Then
            Throw New ArgumentException("Don't pass null to GetDocumentFromWindow, it's not going to work!")
        End If

        '"First try the usual way to get the document."
        'Wait, this is not right: I have to cast DomDocument. Ah, no, that's from Windows Forms.
        Try
            doc = CType(htmlWindow.document, mshtml.IHTMLDocument2)
            Return doc
        Catch comEx As COMException
            Call Debug.Print("COMEX")
            ' I think COMException won't be ever fired but just to be sure ...
            If comEx.ErrorCode <> E_ACCESSDENIED Then
                Return Nothing
            End If
        Catch uae As System.UnauthorizedAccessException
            'Ah, this is the one!
            Call Debug.Print("UAE")
        Catch
            ' Any other error.
            Return Nothing
        End Try

        ' At this point the error was E_ACCESSDENIED because the frame contains a document from another domain.
        ' IE tries to prevent a cross frame scripting security issue.
        '(You know, if this fails, maybe we should just say "Good, it's an advert!" and bin the frame?
        ' That might actually be better...)
        'Return Nothing 'TODO test this out. 
        ' Convert IHTMLWindow2 to IWebBrowser2 using IServiceProvider.
        Try
            Dim sp As IServiceProvider = CType(htmlWindow, IServiceProvider)

            ' Use IServiceProvider.QueryService to get IWebBrowser2 object.
            Dim brws As Object = Nothing
            sp.QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, brws)

            ' Get the document from IWebBrowser2.
            Dim browser As SHDocVw.IWebBrowser2 = CType(brws, SHDocVw.IWebBrowser2)
            Return CType(browser.Document, mshtml.IHTMLDocument2)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Const E_ACCESSDENIED As Integer = &H80070005
    Private IID_IWebBrowserApp As Guid = New Guid("0002DF05-0000-0000-C000-000000000046")
    Private IID_IWebBrowser2 As Guid = New Guid("D30C1661-CDAF-11D0-8A3E-00C04FC9E26E")
    Private IID_IAccessible As Guid = New Guid("618736E0-3C3D-11CF-810C-00AA00389B71")
    Private IID_IHTMLDOMNode As Guid = New Guid("3050F5DA-98B5-11CF-BB82-00AA00BDCE0B")


    <ComImport(),
Guid("6d5140c1-7436-11ce-8034-00aa006009fa"),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
    Interface IServiceProvider
        Sub QueryService(ByRef guidService As Guid, ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppvObject As Object)
    End Interface

End Class
