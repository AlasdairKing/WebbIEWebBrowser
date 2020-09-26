Imports System.Net
Imports System.IO

Module modTelemetry

    ''' <summary>
    ''' Registers with the www.webbie.org.uk server that this program has been used.
    ''' </summary>
    ''' <param name="programId">An identifier for the program, like "WebbIE Web Browser" or "BBC Live Radio"</param>
    ''' <remarks>Calls some PHP on the WebbIE server, which should write to a SQL database.</remarks>
    Public Sub RegisterUse(programId As String)
        programId = programId.Replace(" ", "_")
        If My.Settings.TelemetryId = "NOTDEFINED" Then
            My.Settings.TelemetryId = System.Guid.NewGuid().ToString()
        End If
        Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create("http://www.webbie.org.uk/telemetry.php?Product=" & programId & "&Guid=" & My.Settings.TelemetryId), HttpWebRequest)
        Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
    End Sub

End Module
