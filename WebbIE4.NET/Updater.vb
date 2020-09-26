''' <summary>
''' Version: 5 January 2013.
''' This is a VB.Net copy of Updater.cs. It should be identical in funcationality. See that for the
''' canonical comments.
''' </summary>
''' <remarks></remarks>
Public Class Updater

    Public Shared Sub CheckForUpdates(urlToCheck As String)
        Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Try
            xmlDoc.Load(urlToCheck)
        Catch

            ' Not online or problem with XML. Fail silently.
            Return
        End Try

        Try
            Dim version As String = xmlDoc.DocumentElement.SelectSingleNode("Version").InnerText
            If version <> System.Windows.Forms.Application.ProductVersion Then
                ' Need to update.
                Dim wc As System.Net.WebClient = New System.Net.WebClient()
                Dim filename As String = xmlDoc.DocumentElement.SelectSingleNode("Filename").InnerText
                Dim path As String = System.IO.Path.GetTempPath() & filename
                If System.IO.File.Exists(path) Then
                    Try
                        System.IO.File.Delete(path)
                        System.Windows.Forms.Application.DoEvents()
                    Catch
                        ' Probably already in use. Fail silently.
                        Return
                    End Try
                    Try
                        ' Download the new installer.
                        Call wc.DownloadFileAsync(New System.Uri(xmlDoc.DocumentElement.SelectSingleNode("URL").InnerText), path)
                        While wc.IsBusy

                            Call System.Windows.Forms.Application.DoEvents()
                        End While

                        ' Create an installer batch file.
                        ' Get the path
                        Dim batchPath As String = System.IO.Path.GetTempPath() + filename & ".bat"
                        If (System.IO.File.Exists(batchPath)) Then
                            Call System.IO.File.Delete(batchPath)
                            Call System.Windows.Forms.Application.DoEvents()
                        End If
                        ' Write the batch file.
                        Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(batchPath)
                        Call sw.WriteLine("@title Updating " + System.Windows.Forms.Application.ProductName)
                        Call sw.WriteLine("@echo Updating " + System.Windows.Forms.Application.ProductName)
                        Call sw.WriteLine("@cd """ & System.IO.Path.GetTempPath() & """")
                        ' Put in a pause to allow this application to close.
                        Call sw.WriteLine("@choice /D:Y /T:2 /N")
                        Call sw.WriteLine("@msiexec /I """ & filename & """ /passive")
                        'sw.WriteLine("pause")
                        Call sw.Close()

                        ' Now execute the batch file.
                        Dim proc As System.Diagnostics.Process = New System.Diagnostics.Process()
                        ' Turns out we probably want to show the window so the user can get an
                        ' idea something is happening.
                        'proc.StartInfo.CreateNoWindow = true
                        'proc.StartInfo.RedirectStandardOutput = true
                        'proc.StartInfo.RedirectStandardError = true
                        proc.StartInfo.FileName = batchPath
                        proc.StartInfo.UseShellExecute = False
                        Call proc.Start()

                        ' And close this application!
                        Call System.Windows.Forms.Application.Exit()
                        Return
                    Catch
                        ' Failed to download for some reason, fail silently.
                        Return
                    End Try

                End If
            End If
        Catch
            Return
        End Try
    End Sub
End Class
