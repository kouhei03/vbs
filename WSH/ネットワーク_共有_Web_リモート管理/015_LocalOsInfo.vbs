' ローカルコンピューターのOS名、バージョン、アーキテクチャを表示します。
Option Explicit
Dim service, systems, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT Caption, Version, OSArchitecture, BuildNumber FROM Win32_OperatingSystem")
For Each item In systems
    WScript.Echo "Caption=" & item.Caption
    WScript.Echo "Version=" & item.Version & " Build=" & item.BuildNumber & " Architecture=" & item.OSArchitecture
Next
