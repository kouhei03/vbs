' ファイル概要: OS名・バージョン・ビルド番号を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetOSInfo.vbs
' 実行結果: OS名・バージョン・ビルド番号を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim wmiService
Dim operatingSystems
Dim operatingSystem

On Error Resume Next
Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
If Err.Number <> 0 Then
    WScript.Echo "Failed to connect to WMI: " & Err.Description
    WScript.Quit 1
End If

Err.Clear
' [処理解説] WMIの Win32_OperatingSystem を照会し、Caption, Version, BuildNumber を取得します。
Set operatingSystems = wmiService.ExecQuery( _
    "SELECT Caption, Version, BuildNumber FROM Win32_OperatingSystem")
If Err.Number <> 0 Then
    WScript.Echo "Failed to query operating system information: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

If operatingSystems.Count = 0 Then
    WScript.Echo "No operating system information was returned."
    WScript.Quit 1
End If

For Each operatingSystem In operatingSystems
    WScript.Echo "OS Name: " & operatingSystem.Caption
    WScript.Echo "Version: " & operatingSystem.Version
    WScript.Echo "Build Number: " & operatingSystem.BuildNumber
Next
