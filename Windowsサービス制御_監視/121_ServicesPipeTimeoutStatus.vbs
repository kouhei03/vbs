' ファイル概要: ServicesPipeTimeout値の読み取り。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 121_ServicesPipeTimeoutStatus.vbs
' 実行結果: ServicesPipeTimeout値の読み取り処理の結果やエラーをコンソールに表示します。

' ServicesPipeTimeoutの現在値を読み取ります。レジストリ変更はしません。
Option Explicit
Dim shell, value, registryPath
Set shell = CreateObject("WScript.Shell")
registryPath = "HKLM\SYSTEM\CurrentControlSet\Control\ServicesPipeTimeout"
On Error Resume Next
value = shell.RegRead(registryPath)
If Err.Number <> 0 Then
    WScript.Echo "ServicesPipeTimeout is not explicitly set or cannot be read."
    Err.Clear
Else
    WScript.Echo "ServicesPipeTimeout (milliseconds)=" & value
End If
On Error GoTo 0
