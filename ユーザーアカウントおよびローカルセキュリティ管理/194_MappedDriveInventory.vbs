Option Explicit

Dim service, drives, drive
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drives = service.ExecQuery( _
    "SELECT DeviceID, ProviderName, FreeSpace, Size FROM Win32_LogicalDisk WHERE DriveType=4")
WScript.Echo "Drive" & vbTab & "Provider" & vbTab & "FreeBytes" & vbTab & "SizeBytes"
For Each drive In drives
    WScript.Echo drive.DeviceID & vbTab & drive.ProviderName & vbTab & _
        drive.FreeSpace & vbTab & drive.Size
Next
