Option Explicit
Dim service, drives, drive, usedBytes, usedPercent
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drives = service.ExecQuery("SELECT DeviceID, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType = 3")
For Each drive In drives
    If Not IsNull(drive.Size) And Not IsNull(drive.FreeSpace) Then
        usedBytes = CDbl(drive.Size) - CDbl(drive.FreeSpace)
        usedPercent = 0
        If CDbl(drive.Size) > 0 Then usedPercent = (usedBytes / CDbl(drive.Size)) * 100
        WScript.Echo drive.DeviceID & " FileSystem=" & drive.FileSystem
        WScript.Echo "Total bytes=" & drive.Size & " Used bytes=" & usedBytes & " Free bytes=" & drive.FreeSpace
        WScript.Echo "Used percent=" & FormatNumber(usedPercent, 2)
    End If
Next
