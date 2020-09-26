Option Strict On
Option Explicit On
Friend Class frmLinks
    Inherits System.Windows.Forms.Form

    Private Sub cmdCloseLinks_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error Resume Next
        cmdGo.Enabled = False
        Call Me.Hide()
    End Sub

    Private Sub cmdGo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGo.Click
        On Error Resume Next
        Call DoGo()
    End Sub

    Private Sub DoGo()
        On Error Resume Next
        Dim selection As Integer

        selection = lstLinks.SelectedIndex
        If selection = -1 Then
            'nothing selected - don't do anything
        Else
            If lstLinks.Text = GetText("RSS News Feed") Then
                Call frmMain.RSS()
            ElseIf (optSortPageOrder.Checked) Then
                Call frmMain.StartNavigating(gLinks(selection).address)
            Else
                Call frmMain.StartNavigating(gSortedLink(selection).address)
            End If
            Call Me.Hide()
        End If
    End Sub


    Private Sub frmLinks_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error Resume Next
        lstLinks.Font = frmMain.txtText.Font
        If lstLinks.SelectedIndex > -1 Then
            cmdGo.Enabled = True
        End If
        Call lstLinks.Focus()
    End Sub


    Private Sub lstLinks_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstLinks.SelectedIndexChanged
        On Error Resume Next
        cmdGo.Enabled = (lstLinks.SelectedIndex > -1)
    End Sub

    Private Sub lstLinks_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstLinks.DoubleClick
        Call DoGo()
    End Sub

    ''' <summary>
    ''' Populates the list of links from the gLinks or gSortedLinks array.
    ''' </summary>
    ''' <param name="newList"></param>
    ''' <remarks></remarks>
    Public Sub PopulateList(Optional ByRef newList As Boolean = True)
        Try
            Dim i As Integer
            Dim address As String
            Dim description As String
            Dim linkText As String
            Dim currentAddress As String

            ' Remember the currently-selected link, if there is one.
            If lstLinks.SelectedIndex = -1 Then
                'no selected link
                currentAddress = ""
            Else
                'we have a selected link: remember what it is for when we change link
                'order so it can remain the selected link.
                currentAddress = lstLinks.Text
            End If
            Call lstLinks.Items.Clear()

            Dim pageURI As System.Uri = modGlobals.gWebHost.webMain.Url

            'Iterate through sorted links or unsorted links array, adding to list.
            For i = 0 To gNumLinks - 1
                If optSortPageOrder.Checked Then
                    'page order
                    address = gLinks(i).address
                    description = gLinks(i).description
                Else
                    'alphabetical
                    address = gSortedLink(i).address
                    description = gSortedLink(i).description
                End If
                If address = "FAKEA-RSS" Then
                    linkText = GetText("RSS News Feed")
                Else
                    Dim url As System.Uri
                    Try
                        url = New System.Uri(address)
                        address = url.Host & "/" & url.LocalPath
                    Catch
                        url = Nothing
                        address = address
                    End Try
                    If address.StartsWith("/C:\") Then address = address.Replace("/C:\", "")
                    While address.EndsWith("/")
                        address = address.Substring(0, address.Length - 1)
                    End While
                    address = address.Replace("//", "/")
                    If url Is Nothing Then
                    ElseIf pageURI.Host = url.Host Then
                    Else
                        address = GetText("(out)") & " " & address
                    End If
                    If description = "" Then
                        linkText = address
                    Else
                        linkText = description & " - " & address
                    End If
                End If
                Call lstLinks.Items.Add(linkText)
                'check to see if this link we have just added was the current link
                'If so, make it the selected link
                If linkText = currentAddress Then
                    'found it!
                    lstLinks.SelectedIndex = lstLinks.Items.Count - 1
                End If
            Next i
            ' If we failed to find a link to select as we went through, or we are loading
            ' a new list, then put focus on item 9.
            If (newList Or lstLinks.SelectedIndex = -1) And (lstLinks.Items.Count > 0) Then
                lstLinks.SelectedIndex = 0
            End If
            Call lstLinks.Focus()
        Catch
        End Try
    End Sub

    Private Sub frmLinks_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
        Call lstLinks.Focus()
    End Sub

    Private Sub optSortPageOrder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSortPageOrder.CheckedChanged
        On Error Resume Next
        Call PopulateList()
    End Sub

    Private Sub optSortAlphabetical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSortAlphabetical.CheckedChanged
        On Error Resume Next
        Call PopulateList()
    End Sub

End Class