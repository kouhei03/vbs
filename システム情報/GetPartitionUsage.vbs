Option Explicit

Dim service, drives, drive, usedBytes, usedPercent
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drives = service.ExecQuery( _
    "SELECT DeviceID, Size, FreeSpace FROM Win32_LogicalDisk WHERE DriveType=3")

For Each drive In drives
    If Not IsNull(drive.Size) Then
        If Not IsNull(drive.FreeSpace) Then
            If CDbl(drive.Size) > 0 Then
                usedBytes = CDbl(drive.Size) - CDbl(drive.FreeSpace)
                usedPercent = Round((usedBytes / CDbl(drive.Size)) * 100, 1)
                WScript.Echo drive.DeviceID & " used (bytes): " & usedBytes
                WScript.Echo drive.DeviceID & " free (bytes): " & drive.FreeSpace
                WScript.Echo drive.DeviceID & " usage (%): " & usedPercent
            End If
        End If
    End If
Next
