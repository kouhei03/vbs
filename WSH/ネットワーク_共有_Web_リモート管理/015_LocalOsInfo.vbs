' 概要: ローカルコンピューターのOS名、バージョン、アーキテクチャを表示します。
' 実行方法: cscript //nologo 015_LocalOsInfo.vbs
' 処理結果: ローカルOS情報表示。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, systems, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set systems = service.ExecQuery("SELECT Caption, Version, OSArchitecture, BuildNumber FROM Win32_OperatingSystem")
For Each item In systems
    WScript.Echo "Caption=" & item.Caption
    WScript.Echo "Version=" & item.Version & " Build=" & item.BuildNumber & " Architecture=" & item.OSArchitecture
Next
