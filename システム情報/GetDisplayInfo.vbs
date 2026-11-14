Option Explicit

Dim service, controllers, controller
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set controllers = service.ExecQuery( _
    "SELECT Name, CurrentHorizontalResolution, CurrentVerticalResolution " & _
    "FROM Win32_VideoController")

For Each controller In controllers
    WScript.Echo "Display adapter: " & controller.Name
    If Not IsNull(controller.CurrentHorizontalResolution) And _
        Not IsNull(controller.CurrentVerticalResolution) Then
        WScript.Echo "Reported resolution: " & _
            controller.CurrentHorizontalResolution & " x " & _
            controller.CurrentVerticalResolution
    Else
        WScript.Echo "Current resolution is not reported by this adapter."
    End If
Next
