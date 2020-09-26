Option Strict On
Option Explicit On
Module modHTTPGetter
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
	
	'This module uses Winsock to get files from an internet server
	
	
	Private mTarget As String
	Private mDataReceived As String ' the data we've been sent so far
	Private mDataToReceive As Integer 'how much data to receive in total
	Private mParsedHeader As Boolean 'whether we've already parsed the header and have all its information
	Private mFileName As String 'the name of the target file
	Private mLocalFilename As String ' the name the user wants to save it to disk under
	Private mContentStart As Integer 'stores the start of the post-header content
    Private ReadOnly COMMAND_LINE As String ' the command line for the conversion utility
    Private Const QUOTATION_MARKS As String = """"

    ''' <summary>
    ''' gets the file determined by url, saves it to path
    ''' </summary>
    ''' <param name="url"></param>
    ''' <param name="Path"></param>
    ''' <remarks></remarks>
    Public Sub GetFile(ByRef url As String, ByRef Path As String)
        Try
            'TODO replace this with a .Net library...
            Call DetermineWinsockSettings(url) 'Get settings for Winsock
            mParsedHeader = False
            'work out the filename
            mFileName = Right(url, Len(url) - InStrRev(url, "/", Len(url)))
            mLocalFilename = Path
            mDataReceived = ""
            'Debug.Print "FileName: " & mFileName
            mTarget = url
            frmDownloadProgress.Visible = True
            frmDownloadProgress.Text = ""
            frmDownloadProgress.lblDownloading.Text = ""
            frmDownloadProgress.progressBar.Value = 0
            'Call gTargetForm.Winsock.Connect() 'Get ready to go. When Winsock has connected,
            'it will fire Winsock_Connect, which must call WinsockConnect() below.
        Catch
        End Try
    End Sub

    Private Sub DetermineWinsockSettings(ByRef url As String)
        Try
            'works out the protocol, remote host and other settings needed for
            'Winsock to function. This will probably require some API and registry calls

            Dim hostName As String
            Dim proxy As String

            'look at the address to get to work out the target, if needed
            If InStr(1, url, "http://") = 1 Or InStr(1, url, "ftp://") = 1 Then
                hostName = Mid(url, 8, InStr(8, url, "/") - 8) 'start at 8 to avoid http://
            Else
                hostName = Mid(url, 7, InStr(7, url, "/") - 7) 'start at 7 to avoid ftp://
            End If
            'how are we connecting to the internet? Check out the registry settings.
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Internet Settings", False)
            If regKey.GetValue("ProxyEnable").ToString = "1" Then
                'we're using the proxy
                proxy = regKey.GetValue("ProxyServer").ToString
                'gTargetForm.Winsock.RemoteHost = Left(proxy, InStr(1, proxy, ":") - 1)
                'Debug.Print "RH:" & frmMain.Winsock.RemoteHost
                'gTargetForm.Winsock.RemotePort = CInt(Val(Right(proxy, Len(proxy) - InStr(1, proxy, ":"))))
                '        Debug.Print "RP:" & gTargetForm.Winsock.RemotePort
            Else
                'we're directly connected: use the hostname
                'gTargetForm.Winsock.RemoteHost = hostName
                'gTargetForm.Winsock.RemotePort = 80
            End If
            Call regKey.Close()
        Catch
        End Try
    End Sub

    Public Sub WinsockConnect()
        Try
            'Must be called by Winsock_Connect. Creates the http header and sends it.

            Dim httpRequest As String
            'build request line that contains the HTTP method, 
            'path to the file to retrieve,
            'and HTTP version info. Each line of the request 
            'must be completed by the vbCrLf
            httpRequest = "GET " & mTarget & " HTTP/1.1" & vbCrLf

            'add HTTP headers to the request
            'add required header - "Host", that contains the remote host name
            'httpRequest = httpRequest & "Host: " & gTargetForm.Winsock.RemoteHost & vbCrLf
            'add the "Connection" header to force the server to close the connection
            httpRequest &= "Connection: close" & vbCrLf
            'add optional header "Accept"
            httpRequest &= "Accept: */*" & vbCrLf
            'add other optional headers
            'add a blank line that indicates the end of the request
            httpRequest &= vbCrLf
            'send the request
            'Call gTargetForm.Winsock.SendData(httpRequest)
            'Good, now we wait for the data to arrive through
        Catch
        End Try
    End Sub

    Private Function ParseHeader(ByRef header As String) As Boolean
            'works out all the file information from the header: returns false iff an HTTP
            'error code (value greater than 400, according to http://www.w3.org/Protocols/rfc2616/rfc2616-sec6.html#sec6)
            'is found.
            On Error GoTo tryLater
            Dim contentTypeStart As Integer
            Dim contentLengthStart As Integer
            Dim response() As String
            Dim errorMessage As String
            Dim errorMessageStarts As Integer

            'first check for an error message
            response = Split(header, " ")
            If Val(response(1)) > 399 Then
                'got an error message! Abort file acquisition
                errorMessageStarts = InStr(1, header, response(1)) + Len(response(1))
                errorMessage = Mid(header, errorMessageStarts, InStr(1, header, vbNewLine) - errorMessageStarts)
                Call MsgBox(modI18N.GetText("The website reported the following problem so WebbIE could not download the file:") & " " & response(1) & errorMessage, MsgBoxStyle.Exclamation, Application.ProductName)
                ParseHeader = False
            Else
                'okay, not an error
                ParseHeader = True
                'now check it's complete
                If InStr(1, header, vbCrLf & vbCrLf) > 0 Then
                    'okay, we've got a full header: process it
                    contentTypeStart = InStr(1, header, "Content-Type: ")
                    contentLengthStart = InStr(1, header, "Content-Length: ")
                    mContentStart = InStr(1, header, vbCrLf & vbCrLf) + 4
                    'Debug.Print "Content: [" & Mid(header, mContentStart, 5)
                    mDataToReceive = CInt(Val(Mid(header, contentLengthStart + Len("Content-Length: "), InStr(contentLengthStart, header, vbNewLine) - contentLengthStart - Len("Content-Length: "))))
                    'MsgBox "dtr:" & mDataToReceive
                    mParsedHeader = True
                Else
                    'nope, not complete yet. Hang around until it is.
                End If
            End If
            Exit Function
tryLater:
            Exit Function
    End Function
End Module