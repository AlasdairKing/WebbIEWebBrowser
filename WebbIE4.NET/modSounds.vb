Option Strict On
Option Explicit On
Module modSounds
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

    'Sounds
    ' Contains functions for calling the different navigation and interface sounds
    ' in WebbIE

    ''' <summary>
    ''' Plays the progress sound. 
    ''' </summary>
    ''' <param name="progress">A value from 0 to 9.</param>
    Public Sub PlayProgressSound(ByVal progress As Integer)
        Try
            If progress < 0 Then progress = 0
            If progress > 9 Then progress = 9
            Select Case progress
                Case 0
                    Call PlayResourceWav2(My.Resources.progress0)
                Case 1
                    Call PlayResourceWav2(My.Resources.progress1)
                Case 2
                    Call PlayResourceWav2(My.Resources.progress2)
                Case 3
                    Call PlayResourceWav2(My.Resources.progress3)
                Case 4
                    Call PlayResourceWav2(My.Resources.progress4)
                Case 5
                    Call PlayResourceWav2(My.Resources.progress5)
                Case 6
                    Call PlayResourceWav2(My.Resources.progress6)
                Case 7
                    Call PlayResourceWav2(My.Resources.progress7)
                Case 8
                    Call PlayResourceWav2(My.Resources.progress8)
                Case 9
                    Call PlayResourceWav2(My.Resources.progress9)
            End Select
        Catch
        End Try
    End Sub

    Public Sub PlayResourceWav2(r As System.IO.Stream)
        Try
            If My.Settings.NavigationSounds Then
                Dim sp As System.Media.SoundPlayer = New System.Media.SoundPlayer(r)
                Call sp.Play()
            End If
        Catch
        End Try
    End Sub

    Public Sub PlayDoneSound()
        Try
            'plays the sound when the program has finished - a "dong"
            Call PlayResourceWav2(My.Resources.done)
        Catch
        End Try
    End Sub

    Public Sub PlayErrorSound()
        On Error Resume Next
        'plays sound to indicate error
        'In fact, can just use Beep
        Call Beep()
    End Sub

    Public Sub PlayWorkingSound()
        Try
            'plays a sound to indicate WebbIE is working
            Call PlayResourceWav2(My.Resources.working)
        Catch
        End Try
    End Sub

    Public Sub PlayLocationRefreshSound()
        On Error Resume Next
        'plays the sound when an RSS feed is discovered
        Call PlayResourceWav2(My.Resources.locationrefresh)
    End Sub

    Public Sub PlayStartSound()
        On Error Resume Next
        'plays the sound when an RSS feed is discovered
        Call PlayResourceWav2(My.Resources.start)
    End Sub

    Public Sub PlayRSSSound()
        On Error Resume Next
        'plays the sound when an RSS feed is discovered
        Call PlayResourceWav2(My.Resources.rss_audio)
    End Sub

    Public Sub PlayAjaxSound()
        On Error Resume Next
        'plays the sound when an Ajax-inspired focus change occurs
        'plays the sound when an RSS feed is discovered
        Call PlayResourceWav2(My.Resources.ajaxfocus)
    End Sub
End Module