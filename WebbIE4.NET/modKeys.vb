Option Strict On
Option Explicit On
Module modKeys
    '   This file is part of WebbIE.
    '
    '    WebbIE is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General private License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.
    '
    '    WebbIE is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General private License for more details.
    '
    '    You should have received a copy of the GNU General private License
    '    along with WebbIE.  If not, see <http://www.gnu.org/licenses/>.

    Public Function IsShiftPressed() As Boolean
        On Error Resume Next
        Return (NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) < 0)
    End Function

    Public Function IsCapslockPressed() As Boolean
        On Error Resume Next
        Return (NativeMethods.GetKeyState(NativeMethods.VK_CapsLock) < 0)
    End Function

End Module