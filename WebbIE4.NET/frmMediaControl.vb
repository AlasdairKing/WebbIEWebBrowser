Public Class frmMediaControl

    'Control audio and video elements through a UI.
    'http://msdn.microsoft.com/en-us/library/system.type.invokemember%28v=vs.71%29.aspx

    Private Enum MediaType
        Audio
        Video
    End Enum

    Private Structure MediaElement
        Dim type As MediaType
        Dim index As Integer
    End Structure

    Private mMediaElement As MediaElement

    Private Sub SetElement(index As Integer)
        Try
            mMediaElement.index = index
            Me.btnOpen.Enabled = Not AllSrcAttributesAreData(CurrentMediaItem())
        Catch
        End Try
    End Sub

    Public Sub SetVideoElement(index As Integer)
        Try
            mMediaElement.type = MediaType.Video
            Call SetElement(index)
        Catch
        End Try
    End Sub

    Public Sub SetAudioElement(index As Integer)
        Try
            mMediaElement.type = MediaType.Audio
            Call SetElement(index)
        Catch
        End Try
    End Sub

    Private Sub btnPlay_Click(sender As System.Object, e As System.EventArgs) Handles btnPlay.Click
        Try
            Call CurrentMediaItem().InvokeMember("play")
        Catch
        End Try
    End Sub

    Private Sub frmMediaControl_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            mMediaElement.type = MediaType.Audio
        Catch
        End Try
    End Sub

    Private Function CurrentMediaItem() As HtmlElement
        Try
            Dim tagName As String
            If mMediaElement.type = MediaType.Audio Then
                tagName = "AUDIO"
            Else 'Video by default
                tagName = "VIDEO"
            End If
            Return gWebHost.webMain.Document.GetElementsByTagName(tagName).Item(mMediaElement.index)
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' If the current audio of video element and its source children are all data elements, 
    ''' return false.
    ''' </summary>
    ''' <param name="e"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function AllSrcAttributesAreData(e As HtmlElement) As Boolean
        On Error GoTo errorHandler
        Dim src As String = e.GetAttribute("src").Trim().ToUpperInvariant()
        If Not src.StartsWith("DATA:") Then
            Return False
        End If
        For Each ce As HtmlElement In e.Children
            If ce.TagName.ToUpperInvariant() = "SOURCE" Then
                If Not ce.GetAttribute("src").Trim.ToUpperInvariant().StartsWith("DATA:") Then
                    Return False
                End If
            End If
        Next ce
        'If we got here, we didn't find a single src attribute that wasn't a data uri element.
        Return True
errorHandler:
        Return False ' Don't know, something went wrong, so let it be tried!
    End Function

    Private Sub btnStop_Click(sender As Object, e As System.EventArgs) Handles btnStop.Click
        Try
            Call CurrentMediaItem().InvokeMember("stop")
        Catch
        End Try
    End Sub

    Private Sub tmrMedia_Tick(sender As System.Object, e As System.EventArgs) Handles tmrMedia.Tick
        Try
            If Me.Visible Then
                Dim element As HtmlElement = CurrentMediaItem()
                Dim o As Object = element.DomElement
                Dim el As mshtml.IHTMLElement = CType(o, mshtml.IHTMLElement)
                Dim currentTime As Double
                Dim duration As Double
                Dim gotDuration As Boolean
                Dim gotCurrentTime As Boolean
                If el.getAttribute("duration") Is Nothing Then
                    gotDuration = False
                Else
                    gotDuration = Double.TryParse(el.getAttribute("duration").ToString(), duration)
                    If Double.IsNaN(duration) Then ' experienced on facebook.com: audio, duration set, -1 value.
                        gotDuration = False
                    End If
                End If

                If el.getAttribute("currentTime") Is Nothing Then
                    gotCurrentTime = False
                Else
                    gotCurrentTime = Double.TryParse(el.getAttribute("currentTime").ToString(), currentTime)
                    If Double.IsNaN(currentTime) Then
                        gotCurrentTime = False
                    Else
                        If currentTime > duration Then
                            currentTime = duration
                        End If
                    End If
                End If
                'http://stackoverflow.com/questions/136195/trying-to-set-get-a-javascript-variable-in-an-activex-webbrowser-from-c-sharp?

                If gotDuration Then
                    trbMedia.Enabled = True
                    trbMedia.Maximum = CInt(duration * 100)
                    If gotCurrentTime Then
                        trbMedia.Value = CInt(currentTime * 100)
                        lblStatus.Text = Math.Round(currentTime) & " / " & Math.Round(duration)
                    Else
                        lblStatus.Text = Math.Round(duration).ToString()
                    End If
                Else
                    lblStatus.Text = ""
                    trbMedia.Enabled = False
                End If
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Launch the media content in the machine's default player.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        On Error Resume Next
        Dim mediaItem As HtmlElement = CurrentMediaItem()
        'Check for some error, do a beep if found.
        If IsNothing(mediaItem) Then
            Call Beep()
            Return
        End If
        'OK, now try to get the url: the element itself may have a src, in which case 
        '(this indicating the only source video) use it.
        Dim url As String = mediaItem.GetAttribute("src")
        If url = "" Then
            'Nope, check for child SOURCE elements with video:
            For Each s As HtmlElement In mediaItem.Children
                If s.TagName.ToUpperInvariant() = "SOURCE" Then
                    url = s.GetAttribute("src").Trim().ToUpperInvariant()
                    If url.Length > 0 Then
                        Dim sourceType As String = s.GetAttribute("type").Trim().ToUpperInvariant()
                        If sourceType.Contains("VIDEO/MP4") Then
                            'Use this one!
                            Exit For
                        ElseIf sourceType.Contains("AUDIO/MP3") Then
                            'Use this one!
                            Exit For
                        Else
                            url = ""
                        End If
                    End If
                End If
            Next s
        End If
        'Got anything?
        If url = "" Then
            'Failed to get the url when looking for video/mp4 or audio/mp3, just take the first child source...
            For Each s As HtmlElement In mediaItem.Children
                If s.TagName.ToUpperInvariant() = "SOURCE" Then
                    url = s.GetAttribute("src").Trim().ToUpperInvariant()
                    If url.Length > 0 Then
                        Exit For
                    End If
                End If
            Next s
        End If
        'OK, got anything now?
        If url = "" Then
            'Nope, give up.
            Call Beep()
            Return
        End If
        'Yes! Pass it to the default media player, which is likely to be Windows Media Player or iTunes.
        Call Process.Start(url)
    End Sub

    Private Sub frmMediaControl_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        If e.Control And e.KeyCode = Keys.P Then
            'Play
            Call Me.btnPlay_Click(Nothing, Nothing)
            e.Handled = True
        ElseIf e.KeyCode = Keys.Space Then
            'Pause
            Call Me.btnStop_Click(Nothing, Nothing)
            e.Handled = True
        ElseIf e.Control And e.KeyCode = Keys.O Then
            'Open
            Call btnOpen_Click(Nothing, Nothing)
            e.Handled = True
        ElseIf e.KeyCode = Keys.Escape Then
            'Close dialog
            e.Handled = True
            Call Me.Close()
        End If
    End Sub

    Private Sub trbMedia_Scroll(sender As Object, e As EventArgs) Handles trbMedia.Scroll

    End Sub
End Class