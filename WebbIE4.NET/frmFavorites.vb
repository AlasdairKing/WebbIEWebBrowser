Public Class frmFavorites

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        On Error Resume Next
        Call Me.Hide()
    End Sub

    Private Sub cmdOK_Click(sender As Object, e As System.EventArgs) Handles cmdOK.Click
        On Error Resume Next
        If tvwFavorites.SelectedNode Is Nothing Then
        ElseIf tvwFavorites.SelectedNode.Tag Is Nothing Then
        ElseIf tvwFavorites.SelectedNode.Tag.ToString().Length = 0 Then
        Else
            frmMain.StartNavigating(tvwFavorites.SelectedNode.Tag.ToString())
        End If
        Call Me.Hide()
    End Sub

    Private Sub frmFavorites_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
        cmdCancel.Left = 0 - cmdCancel.Width - 100
        cmdOK.Left = 0 - cmdOK.Width - 100
        tvwFavorites.Nodes.Clear()
        Call LoadFavorites(Nothing, New System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Favorites)))
    End Sub

    Private Sub LoadFavorites(ByVal currentNode As System.Windows.Forms.TreeNode, ByVal dir As System.IO.DirectoryInfo)
        'Iterate through dir's files, adding to currentMenu.
        Try
            Dim f As System.IO.FileInfo
            Dim childDir As System.IO.DirectoryInfo
            Dim url As String
            Dim name As String
            'Do subfolders. 
            For Each childDir In dir.EnumerateDirectories
                name = ""
                name = childDir.Name
                Dim newTreeNode As System.Windows.Forms.TreeNode
                If currentNode Is Nothing Then
                    newTreeNode = tvwFavorites.Nodes.Add(name)
                Else
                    newTreeNode = currentNode.Nodes.Add(name)
                End If
                Call LoadFavorites(newTreeNode, childDir)
            Next childDir
            'Do favorites. It seems more standard to put folders first, I think? Though
            'going to your favorites doesn't think that way: it does links first. TODO.
            For Each f In dir.EnumerateFiles()
                name = ""
                name = f.Name
                url = modIniFile.GetString("InternetShortcut", "URL", "", f.FullName)
                If name <> "" And url <> "" Then
                    name = Replace(name, ".url", "", , , CompareMethod.Text)
                    'tsi.Tag = url
                    Dim newNode As System.Windows.Forms.TreeNode
                    If currentNode Is Nothing Then
                        newNode = tvwFavorites.Nodes.Add(name)
                    Else
                        newNode = currentNode.Nodes.Add(name)
                    End If
                    newNode.Tag = url
                End If
            Next f
        Catch
            'System IO error, probably: just fail, don't continue.
        End Try
    End Sub
End Class