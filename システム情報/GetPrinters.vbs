Option Explicit

Dim service, printers, printer
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set printers = service.ExecQuery( _
    "SELECT Name, Default, Network, Local FROM Win32_Printer")

For Each printer In printers
    WScript.Echo "Printer: " & printer.Name
    WScript.Echo "Default: " & printer.Default
    WScript.Echo "Network printer: " & printer.Network
    WScript.Echo "Local printer: " & printer.Local
    WScript.Echo ""
Next
