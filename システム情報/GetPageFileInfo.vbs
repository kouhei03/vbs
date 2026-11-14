Option Explicit

Dim service, settings, setting, usages, usage, found
Set service = GetObject("winmgmts:\\.\root\cimv2")

Set settings = service.ExecQuery( _
    "SELECT Name, InitialSize, MaximumSize FROM Win32_PageFileSetting")
found = False
For Each setting In settings
    found = True
    WScript.Echo "Configured page file: " & setting.Name
    WScript.Echo "Initial size (MB): " & setting.InitialSize
    WScript.Echo "Maximum size (MB): " & setting.MaximumSize
Next
If Not found Then WScript.Echo "No explicit page file settings returned (possibly system-managed)."

Set usages = service.ExecQuery( _
    "SELECT Name, AllocatedBaseSize, CurrentUsage, PeakUsage FROM Win32_PageFileUsage")
For Each usage In usages
    WScript.Echo "Active page file: " & usage.Name
    WScript.Echo "Allocated size (MB): " & usage.AllocatedBaseSize
    WScript.Echo "Current usage (MB): " & usage.CurrentUsage
    WScript.Echo "Peak usage (MB): " & usage.PeakUsage
Next
