Module modSendKeys

    Public Sub SendPaste()
        On Error Resume Next
        Call NativeMethods.keybd_event(NativeMethods.VK_CONTROL, 0, 0, IntPtr.Zero)
        Call NativeMethods.keybd_event(NativeMethods.VK_V, 0, 0, IntPtr.Zero)

        Call NativeMethods.keybd_event(NativeMethods.VK_V, 0, NativeMethods.KEYEVENTF_KEYUP, IntPtr.Zero)
        Call NativeMethods.keybd_event(NativeMethods.VK_CONTROL, 0, NativeMethods.KEYEVENTF_KEYUP, IntPtr.Zero)
    End Sub

    Public Sub SendSelectAll()
        On Error Resume Next
        Call NativeMethods.keybd_event(NativeMethods.VK_CONTROL, 0, 0, IntPtr.Zero)
        Call NativeMethods.keybd_event(NativeMethods.VK_A, 0, 0, IntPtr.Zero)

        Call NativeMethods.keybd_event(NativeMethods.VK_A, 0, NativeMethods.KEYEVENTF_KEYUP, IntPtr.Zero)
        Call NativeMethods.keybd_event(NativeMethods.VK_CONTROL, 0, NativeMethods.KEYEVENTF_KEYUP, IntPtr.Zero)
    End Sub

End Module
