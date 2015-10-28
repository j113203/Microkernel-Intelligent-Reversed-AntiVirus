Imports System.Runtime.InteropServices

Public Class 程式檢查
    Private Sub 程式攔截_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(Screen.PrimaryScreen.WorkingArea.Width, 11)
        Me.Location = New Point(0, 0)
        Dim CallBack As New Win32_API.WinEventProc(AddressOf EventCallBack)
        Win32_API.SetWinEventHook(Win32_API.EVENT_SYSTEM_FOREGROUND, Win32_API.EVENT_SYSTEM_MOVESIZEEND, IntPtr.Zero, Marshal.GetFunctionPointerForDelegate(CallBack), 0, 0, Win32_API.WINEVENT_OUTOFCONTEXT Or Win32_API.WINEVENT_SKIPOWNPROCESS)
    End Sub
    Public Sub EventCallBack(hWinEventHook As IntPtr, [Event] As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
        Select Case [Event]
            Case Win32_API.EVENT_SYSTEM_FOREGROUND
                Label1.Text = Win32_API.GetWindowTitle(hwnd)
                UpdateLocation(hwnd)
            Case Win32_API.EVENT_SYSTEM_MOVESIZEEND
                UpdateLocation(hwnd)
        End Select
    End Sub
    Sub UpdateLocation(hwnd As IntPtr)
        Dim Location As New RECT
        Win32_API.GetWindowRect(hwnd, Location)
        Me.Size = New Size(Location.Right - Location.Left, 11)
        Me.Location = New Point(Location.Left, Location.Top)
    End Sub
End Class