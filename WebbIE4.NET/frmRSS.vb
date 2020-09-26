Option Strict On
Option Explicit On
Friend Class frmRSS
    Inherits System.Windows.Forms.Form
    '   This file is part of WebbIE.
    '
    '    WebbIE is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General Public License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.
    '
    '    WebbIE is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General Public License for more details.
    '
    '    You should have received a copy of the GNU General Public License
    '    along with WebbIE.  If not, see <http://www.gnu.org/licenses/>.


    Private WithEvents mRSSFeed As System.Xml.XmlDocument
    Private mRSSURL As String
    Private mLinks As Collection = New Collection()
    Private ReadOnly mRSSThread As Threading.Thread

    Private Sub Initialise()
        Static initialised As Boolean
        Try
            If initialised Then
            Else
                mRSSFeed = New System.Xml.XmlDocument
                Call lstItems.Items.Add(modI18N.GetText("No RSS News feed found."))
                lstItems.SelectedIndex = 0
            End If
            initialised = True
        Catch
            initialised = True
        End Try
    End Sub

    Private Sub DisplayRSS(ByRef url As String)
        Try
            Call Initialise()
            Call lstItems.Items.Clear()
            mLinks = New Collection
            'Call frmMain.cboAddress.Clear
            If Len(url) = 0 Then
                'no RSS feed
                Call lstItems.Items.Add(modI18N.GetText("No RSS News feed found."))
                btnGo.Enabled = False
            Else
                'Threading, which sadly I don't understand yet!
                mRSSURL = url
                'http://msdn.microsoft.com/en-us/library/ms171728(v=vs.100).aspx?appId=Dev10IDEF1&l=EN-US&k=k(EHINVALIDOPERATION.WINFORMS.ILLEGALCROSSTHREADCALL)%3bk(TargetFrameworkMoniker-".NETFRAMEWORK,VERSION%3dV4.0")%3bk(DevLang-VB)&rd=true&cs-save-lang=1&cs-lang=vb#code-snippet-6
                'mRSSThread = New System.Threading.Thread(AddressOf GetRSS)
                'mRSSThread.IsBackground = True
                'mRSSThread.Start()
                Call GetRSS()
            End If
        Catch
        End Try
    End Sub

    Private Sub GetRSS()
        Try
            Dim nodeIterator As System.Xml.XmlNode
            Dim itemsFound As Boolean = False
            Dim text_Renamed As String
            Dim description As String
            Dim gotFeed As Boolean = False

            Call Initialise()
            'Get RSS feed
            Try
                Call mRSSFeed.Load(mRSSURL)
                gotFeed = True
            Catch webEx As System.Net.WebException
                'Couldn't reach that RSS URL, give up.
            Catch inOp As System.InvalidOperationException
                'Invalid XML, give up.
            Catch xmlEx As System.Xml.XmlException
                'Invalid XML, give up
                Call lstItems.Items.Add(modI18N.GetText("RSS Feed broken!"))
                btnGo.Enabled = False
            End Try
            If gotFeed Then
                'parsed okay: display.
                Me.btnGo.Enabled = True
                'Tried processing channels and items, but because of the varied
                'RSS formats, didn't work on some sites. 21 Feb 2007.
                If mRSSFeed.DocumentElement.SelectSingleNode("channel/title") Is Nothing Then
                    Me.Text = modI18N.GetText("RSS News Feed")
                Else
                    Me.Text = modI18N.GetText("RSS News Feed") & " - " & mRSSFeed.DocumentElement.SelectSingleNode("channel/title").InnerText
                End If
                For Each nodeIterator In mRSSFeed.SelectNodes("//item")
                    'For Each channelIterator In mRSSFeed.documentElement.selectNodes("channel")
                    '    For Each nodeIterator In channelIterator.selectNodes("item")
                    Call mLinks.Add(nodeIterator.SelectSingleNode("link").InnerText)
                    text_Renamed = nodeIterator.SelectSingleNode("title").InnerText
                    description = nodeIterator.SelectSingleNode("description").InnerText
                    If Len(description) > 0 Then
                        'okay, we have a description: at least one site makes this a link,
                        'not text, which isn't helpful. So reject any links
                        If InStr(1, description, "<") = 0 Then
                            text_Renamed = text_Renamed & " - " & description
                        End If
                    End If
                    text_Renamed = Replace(text_Renamed, "&amp;", "&")
                    text_Renamed = Replace(text_Renamed, "&nbsp;", " ")
                    text_Renamed = Replace(text_Renamed, "&rsquo;", "'")
                    Call lstItems.Items.Add(text_Renamed)
                    'Call frmMain.cboAddress.AddItem(text)
                    itemsFound = True
                    '    Next nodeIterator
                    'Next channelIterator
                Next nodeIterator
            End If

            If itemsFound Then
                btnGo.Enabled = True
            Else
                btnGo.Enabled = False
                Call lstItems.Items.Add(modI18N.GetText("No items in RSS News feed."))
            End If
            lstItems.SelectedIndex = 0
        Catch
        End Try
    End Sub

    Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnGo.Click
        On Error Resume Next
        Call Initialise()
        Call GotoItem(lstItems.SelectedIndex)
    End Sub

    Public Sub GotoItem(ByRef index As Integer)
        Try
            Call Initialise()
            If index >= 0 And mLinks.Count() > 0 Then
                Call frmMain.StartNavigating(mLinks.Item(index + 1).ToString)
                Call Me.Hide()
                'Call frmMain.SetFocus Gets into a nasty race with frmMain to have the focus!
            End If
        Catch
        End Try
    End Sub

    Private Sub frmRSS_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error Resume Next
        Call DisplayRSS(gRSSFeedURL)
        Call lstItems.Focus()
    End Sub

    Private Sub frmRSS_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Try
            Call Initialise()
            Dim KeyCode As Integer = eventArgs.KeyCode
            Dim Shift As Integer = eventArgs.KeyData \ &H10000

            If KeyCode = System.Windows.Forms.Keys.E And eventArgs.Control Then
                Call Me.Hide()
                'Call frmMain.SetFocus There's a nasty race with frmMain to get focus, might be this.
            End If
        Catch
        End Try
    End Sub

    Private Sub lstItems_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstItems.DoubleClick
        Try
            Call Initialise()
            If btnGo.Enabled Then
                Call cmdOpen_Click(btnGo, New System.EventArgs())
            End If
        Catch
        End Try
    End Sub

    Private Sub lstItems_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lstItems.KeyDown
        Try
            Dim KeyCode As Integer = eventArgs.KeyCode
            Dim Shift As Integer = eventArgs.KeyData \ &H10000

            Call Initialise()
            If KeyCode = System.Windows.Forms.Keys.Up And lstItems.SelectedIndex = 0 Then
                Call PlayErrorSound()
            ElseIf KeyCode = System.Windows.Forms.Keys.Down And lstItems.SelectedIndex = lstItems.Items.Count - 1 Then
                Call PlayErrorSound()
            End If
        Catch
        End Try
    End Sub

    Private Sub frmRSS_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        On Error Resume Next
        Call modI18N.DoForm(Me)
        Call frmColourSelect.SetColourScheme(Me, CType(My.Settings.ColourScheme, ColourScheme))
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        On Error Resume Next
        Call Me.Hide()
    End Sub
End Class