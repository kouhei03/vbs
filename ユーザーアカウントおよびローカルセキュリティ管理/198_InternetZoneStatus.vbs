' ファイル概要: 現在のInternetゾーンのセキュリティレベルを照会する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 198_InternetZoneStatus.vbs
' 実行結果: 現在のInternetゾーンのセキュリティレベルを照会する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim registry, result, zoneValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetDWORDValue(&H80000001, _
    "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", _
    "CurrentLevel", zoneValue)
If result = 0 Then
    WScript.Echo "Current Internet zone security level: " & zoneValue
ElseIf result = 2 Then
    WScript.Echo "Current Internet zone level is not explicitly set for this user."
Else
    WScript.Echo "Zone setting unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Read-only query; zone or trusted-site settings were not changed."
