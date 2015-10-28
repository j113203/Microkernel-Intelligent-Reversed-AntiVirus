Imports System.Runtime.InteropServices
Imports System.Text
Public Class Win32_API
    <DllImport("user32.dll")>
    Public Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function
    <DllImport("user32.dll", EntryPoint:="GetWindowText", SetLastError:=True)>
    Public Shared Function GetWindowText(hwnd As IntPtr, title As StringBuilder, cp As Integer) As Integer
    End Function
    Public Shared Function GetWindowTitle(hwnd As IntPtr) As String
        Dim Buff As New StringBuilder(500)
        GetWindowText(hwnd, Buff, Buff.Capacity)
        Return Buff.ToString()
    End Function
    <DllImport("user32.dll", EntryPoint:="SetWinEventHook")>
    Public Shared Function SetWinEventHook(ByVal eventMin As UInteger, ByVal eventMax As UInteger, ByVal hmodWinEventProc As IntPtr, ByVal lpfnWinEventProc As IntPtr, ByVal idProcess As UInteger, ByVal idThread As UInteger, ByVal dwFlags As UInteger) As IntPtr
    End Function
    Public Delegate Sub WinEventProc(hWinEventHook As IntPtr, [Event] As UInteger, hwnd As IntPtr, idObject As Integer, idChild As Integer, dwEventThread As UInteger, dwmsEventTime As UInteger)
    Public Const WINEVENT_OUTOFCONTEXT As UInteger = &H0
    Public Const WINEVENT_SKIPOWNTHREAD As UInteger = &H1
    Public Const WINEVENT_SKIPOWNPROCESS As UInteger = &H2
    Public Const WINEVENT_INCONTEXT As UInteger = &H4
    Public Const EVENT_MIN As UInteger = &H1
    Public Const EVENT_MAX As UInteger = &H7FFFFFFF
    Public Const EVENT_SYSTEM_SOUND As UInteger = &H1
    Public Const EVENT_SYSTEM_ALERT As UInteger = &H2
    Public Const EVENT_SYSTEM_FOREGROUND As UInteger = &H3
    Public Const EVENT_SYSTEM_MENUSTART As UInteger = &H4
    Public Const EVENT_SYSTEM_MENUEND As UInteger = &H5
    Public Const EVENT_SYSTEM_MENUPOPUPSTART As UInteger = &H6
    Public Const EVENT_SYSTEM_MENUPOPUPEND As UInteger = &H7
    Public Const EVENT_SYSTEM_CAPTURESTART As UInteger = &H8
    Public Const EVENT_SYSTEM_CAPTUREEND As UInteger = &H9
    Public Const EVENT_SYSTEM_MOVESIZESTART As UInteger = &HA
    Public Const EVENT_SYSTEM_MOVESIZEEND As UInteger = &HB
    Public Const EVENT_SYSTEM_CONTEXTHELPSTART As UInteger = &HC
    Public Const EVENT_SYSTEM_CONTEXTHELPEND As UInteger = &HD
    Public Const EVENT_SYSTEM_DRAGDROPSTART As UInteger = &HE
    Public Const EVENT_SYSTEM_DRAGDROPEND As UInteger = &HF
    Public Const EVENT_SYSTEM_DIALOGSTART As UInteger = &H10
    Public Const EVENT_SYSTEM_DIALOGEND As UInteger = &H11
    Public Const EVENT_SYSTEM_SCROLLINGSTART As UInteger = &H12
    Public Const EVENT_SYSTEM_SCROLLINGEND As UInteger = &H13
    Public Const EVENT_SYSTEM_SWITCHSTART As UInteger = &H14
    Public Const EVENT_SYSTEM_SWITCHEND As UInteger = &H15
    Public Const EVENT_SYSTEM_MINIMIZESTART As UInteger = &H16
    Public Const EVENT_SYSTEM_MINIMIZEEND As UInteger = &H17
    Public Const EVENT_SYSTEM_DESKTOPSWITCH As UInteger = &H20
    Public Const EVENT_SYSTEM_END As UInteger = &HFF
    Public Const EVENT_OEM_DEFINED_START As UInteger = &H101
    Public Const EVENT_OEM_DEFINED_END As UInteger = &H1FF
    Public Const EVENT_UIA_EVENTID_START As UInteger = &H4E00
    Public Const EVENT_UIA_EVENTID_END As UInteger = &H4EFF
    Public Const EVENT_UIA_PROPID_START As UInteger = &H7500
    Public Const EVENT_UIA_PROPID_END As UInteger = &H75FF
    Public Const EVENT_CONSOLE_CARET As UInteger = &H4001
    Public Const EVENT_CONSOLE_UPDATE_REGION As UInteger = &H4002
    Public Const EVENT_CONSOLE_UPDATE_SIMPLE As UInteger = &H4003
    Public Const EVENT_CONSOLE_UPDATE_SCROLL As UInteger = &H4004
    Public Const EVENT_CONSOLE_LAYOUT As UInteger = &H4005
    Public Const EVENT_CONSOLE_START_APPLICATION As UInteger = &H4006
    Public Const EVENT_CONSOLE_END_APPLICATION As UInteger = &H4007
    Public Const EVENT_CONSOLE_END As UInteger = &H40FF
    Public Const EVENT_OBJECT_CREATE As UInteger = &H8000
    Public Const EVENT_OBJECT_DESTROY As UInteger = &H8001
    Public Const EVENT_OBJECT_SHOW As UInteger = &H8002
    Public Const EVENT_OBJECT_HIDE As UInteger = &H8003
    Public Const EVENT_OBJECT_REORDER As UInteger = &H8004
    Public Const EVENT_OBJECT_FOCUS As UInteger = &H8005
    Public Const EVENT_OBJECT_SELECTION As UInteger = &H8006
    Public Const EVENT_OBJECT_SELECTIONADD As UInteger = &H8007
    Public Const EVENT_OBJECT_SELECTIONREMOVE As UInteger = &H8008
    Public Const EVENT_OBJECT_SELECTIONWITHIN As UInteger = &H8009
    Public Const EVENT_OBJECT_STATECHANGE As UInteger = &H800A
    Public Const EVENT_OBJECT_LOCATIONCHANGE As UInteger = &H800B
    Public Const EVENT_OBJECT_NAMECHANGE As UInteger = &H800C
    Public Const EVENT_OBJECT_DESCRIPTIONCHANGE As UInteger = &H800D
    Public Const EVENT_OBJECT_VALUECHANGE As UInteger = &H800E
    Public Const EVENT_OBJECT_PARENTCHANGE As UInteger = &H800F
    Public Const EVENT_OBJECT_HELPCHANGE As UInteger = &H8010
    Public Const EVENT_OBJECT_DEFACTIONCHANGE As UInteger = &H8011
    Public Const EVENT_OBJECT_ACCELERATORCHANGE As UInteger = &H8012
    Public Const EVENT_OBJECT_INVOKED As UInteger = &H8013
    Public Const EVENT_OBJECT_TEXTSELECTIONCHANGED As UInteger = &H8014
    Public Const EVENT_OBJECT_CONTENTSCROLLED As UInteger = &H8015
    Public Const EVENT_SYSTEM_ARRANGMENTPREVIEW As UInteger = &H8016
    Public Const EVENT_OBJECT_END As UInteger = &H80FF
    Public Const EVENT_AIA_START As UInteger = &HA000
    Public Const EVENT_AIA_END As UInteger = &HAFFF
End Class
