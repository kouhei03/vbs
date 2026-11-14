' サービスごとのServiceGroupレジストリ値を読み取ります。レジストリは変更しません。
Option Explicit
Dim service, services, item, shell, registryPath, groupName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
Set shell = CreateObject("WScript.Shell")
For Each item In services
    registryPath = "HKLM\SYSTEM\CurrentControlSet\Services\" & item.Name & "\Group"
    On Error Resume Next
    groupName = shell.RegRead(registryPath)
    If Err.Number <> 0 Then
        groupName = "(not set or unreadable)"
        Err.Clear
    End If
    On Error GoTo 0
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & groupName & vbTab & item.State
Next
