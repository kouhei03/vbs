' ファイル概要: Secure Boot状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 127_SecureBootStatus.vbs
' 実行結果: Secure Boot状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, value
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Control\SecureBoot\State", "UEFISecureBootEnabled", value)
If result = 0 Then
    WScript.Echo "UEFISecureBootEnabled=" & value
ElseIf result = 2 Then
    WScript.Echo "UEFISecureBootEnabled is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
