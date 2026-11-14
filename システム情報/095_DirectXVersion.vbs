' ファイル概要: DirectXバージョン。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 095_DirectXVersion.vbs
' 実行結果: DirectXバージョン処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, value
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\DirectX", "Version", value)
If result = 0 Then
    WScript.Echo "Version=" & value
ElseIf result = 2 Then
    WScript.Echo "Version is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
