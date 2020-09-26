Option Strict On
Option Explicit On
Friend Class clsBackForwardsHandler
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
	
	'Manages the line that should be used for the caret by the Back and Forwards function
	
	
    Private mLineNumbers() As Integer 'contains the line numbers to use
    Private mCurrentPosition As Integer ' the index to the mLineNumbers collection of
    'the current page
    Private mGoingBack As Boolean 'whether the current navigation is back
    Private mGoingForwards As Boolean 'whether the current navigation is forwards

    Public ReadOnly Property goingSomewhere() As Boolean
        Get

            goingSomewhere = (mGoingBack Or mGoingForwards)
        End Get
    End Property

    Public Sub New()
        MyBase.New()
        'setup the array we're going to use
        ReDim mLineNumbers(1000)
        mCurrentPosition = 0
    End Sub

    'store a line number on the onset of a navigation happening
    Public Sub NavigationBegin(ByRef newLineNumber As Integer)

        'okay, store this line number as the current location line number
        Debug.Print("Saved " & newLineNumber & " at position " & mCurrentPosition)
        mLineNumbers(mCurrentPosition) = newLineNumber
        
    End Sub

    'clears the going state, moves the current pointer along
    Public Function GetLineNumberAndClear() As Integer

        'change the pointer location
        Dim s As String = ""
        For i As Integer = 0 To 10
            s = s & mLineNumbers(i)
        Next i
        Debug.Print("Contains: [" & s & "]")
        If mGoingBack Then
            If mCurrentPosition > 0 Then
                Debug.Print("Back!")
                mCurrentPosition = mCurrentPosition - 1
            End If
        ElseIf mGoingForwards Then
            'should always be a value here - but I'll check anyway
            If mCurrentPosition <= UBound(mLineNumbers) Then
                Debug.Print("Forwards!")
                mCurrentPosition = mCurrentPosition + 1
            End If
        Else
            'normal navigation: we've reached somewhere new
            mCurrentPosition = mCurrentPosition + 1
            Call CheckForLineNumberArraySize()
            mLineNumbers(mCurrentPosition) = 0
        End If
        'return the new line number
        GetLineNumberAndClear = mLineNumbers(mCurrentPosition)
        Debug.Print("Returning line: " & GetLineNumberAndClear)
        s = ""
        For i As Integer = 0 To 10
            s = s & mLineNumbers(i)
        Next i
        Debug.Print("Contains: [" & s & "]")
        'clear the state variables
        mGoingBack = False
        mGoingForwards = False
    End Function

    'tell the object the user is going back
    Public Sub goingBack()
        mGoingBack = True
    End Sub

    'makes the lineNumber array bigger if necessary
    Private Sub CheckForLineNumberArraySize()

        If mCurrentPosition > UBound(mLineNumbers) Then
            'need to resize mLineNumbers
            ReDim Preserve mLineNumbers(UBound(mLineNumbers) + 500)
        End If
    End Sub

    'tell the object the user is going forwards
    Public Sub goingForwards()

        mGoingForwards = True
    End Sub
End Class