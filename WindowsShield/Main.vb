Imports System.Management

Public Class Main
    Dim 程式攔截_查詢 As New WqlEventQuery("__InstanceCreationEvent", New TimeSpan(0, 0, 1), "TargetInstance isa ""Win32_Process""")
    Dim 程式攔截_監視器 As New ManagementEventWatcher
    Private Sub 程式_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        程式攔截_監視器.Query = 程式攔截_查詢
        AddHandler 程式攔截_監視器.EventArrived, AddressOf 程式攔截
        程式攔截_監視器.Start()
    End Sub
    Dim Whitelist As String() = {
        ""
    }

    Public Sub 程式攔截(sender As Object, e As EventArrivedEventArgs)
        Dim a As ManagementBaseObject = e.NewEvent("TargetInstance")
        Dim output(1) As String
        Dim c As String = ""
        For Each b In a.Properties
            c += b.Name + vbNewLine
        Next
    End Sub
End Class
