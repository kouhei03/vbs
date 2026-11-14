' ファイル概要: BIOSの製造元・バージョン・公開日を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetBIOSInfo.vbs
' 実行結果: BIOSの製造元・バージョン・公開日を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, bioses, bios
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set bioses = service.ExecQuery( _
    "SELECT Manufacturer, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS")

For Each bios In bioses
    WScript.Echo "Manufacturer: " & bios.Manufacturer
    WScript.Echo "Version: " & bios.SMBIOSBIOSVersion
    WScript.Echo "Release date (WMI format): " & bios.ReleaseDate
Next
