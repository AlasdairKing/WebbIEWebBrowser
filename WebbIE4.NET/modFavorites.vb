Module modFavorites

    Public Function GenerateHomepageHTML() As String
        Try
            Dim output As System.Text.StringBuilder = New System.Text.StringBuilder()

            Call output.AppendLine("<html><head><title>" & GetText("WebbIE Homepage") & "</title><body>")
            Call output.AppendLine("<p><a href='websearch_89789798'>" & GetText("Websearch") & "</a></p>")
            Call output.AppendLine("<p><a href='goto_89789798'>" & GetText("Go to...") & "</a></p>")
            Call output.AppendLine(ParseFavoritesFolder(Environment.GetFolderPath(Environment.SpecialFolder.Favorites)))
            Call output.AppendLine("</body></html>")
            Return output.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Function ParseFavoritesFolder(path As String) As String
        Try
            Dim f As System.IO.FileInfo
            Dim fo As System.IO.DirectoryInfo
            Dim url As String
            Dim name As String
            Dim html As System.Text.StringBuilder = New System.Text.StringBuilder()
            'Do urls
            For Each f In New System.IO.DirectoryInfo(path).EnumerateFiles()
                name = ""
                name = f.Name
                url = modIniFile.GetString("InternetShortcut", "URL", "", f.FullName)
                If name <> "" And url <> "" Then
                    name = Replace(name, ".url", "", , , CompareMethod.Text)
                    html.AppendLine("<li><a href=""" & url & """>" & name & "</a></li>")
                End If
            Next f
            'Do subfolders. I'm going to do subfolders after links, otherwise stuff on the top menu will never get seen.
            For Each fo In New System.IO.DirectoryInfo(path).EnumerateDirectories
                html.Append("<li>" & fo.Name & "</li>")
                html.Append("<ul>" & ParseFavoritesFolder(fo.FullName) & "</ul>")
            Next fo
            Return html.ToString()
        Catch
            Return ""
        End Try
    End Function



End Module
