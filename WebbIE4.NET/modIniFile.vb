Option Explicit On
Option Strict On

Module modIniFile

    Public Function GetString(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As String, ByVal strFilename As String) As String
        Try
            ' Returns a string from your INI file
            Dim intCharCount As Integer
            Dim objResult As New System.Text.StringBuilder(256)
            intCharCount = NativeMethods.GetPrivateProfileString(Section, Key,
               [Default], objResult, objResult.Capacity, strFilename)
            If intCharCount > 0 Then
                GetString = Left(objResult.ToString, intCharCount)
            Else
                GetString = ""
            End If
        Catch
            Return ""
        End Try
    End Function

    Public Function GetInteger(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Integer, ByVal strFilename As String) As Integer
        Try
            ' Returns an integer from your INI file
            Return NativeMethods.GetPrivateProfileInt(Section, Key,
               [Default], strFilename)
        Catch
            Return [Default]
        End Try
    End Function

    Public Function GetBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Boolean, ByVal strFilename As String) As Boolean
        ' Returns a boolean from your INI file
        Try
            Return (NativeMethods.GetPrivateProfileInt(Section, Key,
               CInt([Default]), strFilename) = 1)
        Catch
            Return [Default]
        End Try
    End Function

    Public Sub WriteString(ByVal Section As String, _
      ByVal Key As String, ByVal Value As String, ByVal strFilename As String)
        Try
            ' Writes a string to your INI file
            NativeMethods.WritePrivateProfileString(Section, Key, Value, strFilename)
            Flush(strFilename)
        Catch
        End Try
    End Sub

    Public Sub WriteInteger(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Integer, ByVal strFilename As String)
        Try
            ' Writes an integer to your INI file
            WriteString(strFilename, Section, Key, CStr(Value))
            Flush(strFilename)
        Catch
        End Try
    End Sub

    Public Sub WriteBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Boolean, ByVal strFilename As String)
        Try
            ' Writes a boolean to your INI file
            WriteString(strFilename, Section, Key, CStr(CInt(Value)))
            Flush(strFilename)
        Catch
        End Try
    End Sub

    Private Sub Flush(ByVal strFilename As String)
        Try
            ' Stores all the cached changes to your INI file
            NativeMethods.FlushPrivateProfileString(0, 0, 0, strFilename)
        Catch
        End Try
    End Sub

End Module
