Option Explicit
Dim service, drives, drive, driveTypeName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drives = service.ExecQuery("SELECT DeviceID, DriveType, ProviderName FROM Win32_LogicalDisk")
For Each drive In drives
    Select Case drive.DriveType
        Case 0: driveTypeName = "Unknown"
        Case 1: driveTypeName = "No root directory"
        Case 2: driveTypeName = "Removable"
        Case 3: driveTypeName = "Local fixed disk"
        Case 4: driveTypeName = "Network"
        Case 5: driveTypeName = "CD-ROM"
        Case 6: driveTypeName = "RAM disk"
        Case Else: driveTypeName = "Unrecognized type " & drive.DriveType
    End Select
    WScript.Echo drive.DeviceID & ": " & driveTypeName & " Provider=" & drive.ProviderName
Next
