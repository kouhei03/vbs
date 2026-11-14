' ファイル概要: Windowsビルドとリビジョン。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 117_WindowsBuildRevision.vbs
' 実行結果: Windowsビルドとリビジョン処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, systems, system, registry, result, revision
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_OperatingSystem を照会し、Caption, BuildNumber, Version を取得します。
Set systems = service.ExecQuery("SELECT Caption, BuildNumber, Version FROM Win32_OperatingSystem")
For Each system In systems
    WScript.Echo system.Caption & " Build=" & system.BuildNumber & " Version=" & system.Version
Next
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR", revision)
If result = 0 Then
    WScript.Echo "UBR=" & revision
Else
    WScript.Echo "UBR unavailable; WMI code=" & result
End If
