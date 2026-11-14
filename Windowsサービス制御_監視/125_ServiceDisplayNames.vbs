' サービス名と表示名を列挙します。表示名の変更はしません。
Option Explicit
Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
For Each item In services
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
Next
