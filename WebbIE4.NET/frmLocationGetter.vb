Option Strict On
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmLocationGetter
	Inherits System.Windows.Forms.Form
	
	Private mWininetState As Short
	Private mStopGetting As Boolean

    Private mHTTPGetter As System.Net.HttpWebRequest


	Public Sub StartGetting()

        'TODO !
        '        Dim i As Integer
        '        Dim gotAtLeastOne As Boolean
        '        Dim url As String
        '        Dim name_Renamed As String = ""
        '        Dim got As String
        '        Dim doc As String

        '        url = NextLocationToGet()
        '        Debug.Print("url: " & url)
        '        If url <> "" Then
        '            mHTTPGetter = CType(System.Net.WebRequest.Create(url), System.Net.HttpWebRequest)
        '            mHTTPGetter.BeginGetRequestStream(AddressOf GotData, null)
        '        End If

        '        If mWinInet.StillExecuting Then
        '            Debug.Print("Still working...")
        '        Else
        '            mStopGetting = False
        '            url = NextLocationToGet()
        '            Debug.Print("url: " & url)
        '            While Len(url) > 0 And Not mStopGetting And Not gExiting
        '                On Error GoTo SkipThisLocation
        '                Call mWinInet.Execute(url, "GET", "", "")

        '                For i = 1 To 100 : System.Windows.Forms.Application.DoEvents() : Next i
        '                On Error GoTo SkipResponse
        '                While mWininetState <> InetCtlsObjects.StateConstants.icResponseCompleted
        '                    If gExiting Then Exit Sub
        '                    For i = 1 To 100 : System.Windows.Forms.Application.DoEvents() : Next i
        '                End While
        'SkipResponse:

        '                If mWinInet.ResponseCode = 0 Then
        '                    'UPGRADE_WARNING: Couldn't resolve default property of object got. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '                    got = ""
        '                    doc = ""
        '                    On Error GoTo SkipWhile
        '                    got = CType(mWinInet.GetChunk(1024, InetCtlsObjects.DataTypeConstants.icString), String)
        '                    'Do I need more than 1024 bytes? To get the title? I don't think I do.
        '                    'And much quicker to get 1024 bytes.
        '                    doc = doc & CType(got, String)
        '                    '                While Len(got) <> 0
        '                    '                    If gExiting Then Exit Sub
        '                    '                    got = mWinInet.GetChunk(1024, icString)
        '                    '                    doc = doc & got
        '                    '                    For i = 1 To 100: DoEvents: Next i
        '                    '                Wend
        'SkipWhile:

        '                    If InStr(1, doc, "<title>", CompareMethod.Text) > 0 Then
        '                        doc = VB.Right(doc, Len(doc) - InStr(1, doc, "<title>", CompareMethod.Text) - Len("<title>") + 1)
        '                        If InStr(1, doc, "<") > 0 Then
        '                            name_Renamed = VB.Left(doc, InStr(1, doc, "<") - 1)
        '                            Debug.Print("NAME: " & name_Renamed)
        '                        End If
        '                    End If
        '                    If name_Renamed <> "" Then
        '                        Call AddLocationTitle(url, name_Renamed)
        '                        gotAtLeastOne = True
        '                    End If
        '                End If
        'SkipThisLocation:

        '                System.Windows.Forms.Application.DoEvents()
        '                url = NextLocationToGet()
        '            End While
        '            'Debug.Print "Finished getting!"
        '            If Not gExiting And Not mStopGetting And gotAtLeastOne Then
        '                Call PlayLocationRefreshSound()
        '                Call frmMain.RefreshCurrentPage()
        '            End If
        '        End If
    End Sub

    Public Sub StopGetting()

        mStopGetting = True
    End Sub


    'Private Sub mWinInet_StateChanged(ByVal eventSender As System.Object, ByVal eventArgs As AxInetCtlsObjects.DInetEvents_StateChangedEvent) Handles mWinInet.StateChanged

    '    mWininetState = eventArgs.state
    '    If gExiting Then
    '        'UPGRADE_ISSUE: VBControlExtender property mWinInet.Cancel was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '        Call mWinInet.Cancel()
    '    End If
    'End Sub
End Class