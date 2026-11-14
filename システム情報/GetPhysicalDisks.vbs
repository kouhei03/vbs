Option Explicit

Dim service, disks, disk
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set disks = service.ExecQuery( _
    "SELECT Model, Size, MediaType FROM Win32_DiskDrive")

For Each disk In disks
    WScript.Echo "Model: " & disk.Model
    WScript.Echo "Media type: " & disk.MediaType
    WScript.Echo "Capacity (bytes): " & disk.Size
    WScript.Echo ""
Next
