' ファイル概要: レジストリと現在のプロセスのPATHの長さを比較する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo CompareEnvironmentRefreshState.vbs
' 実行結果: レジストリと現在のプロセスのPATHの長さを比較する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, registry, processPath, systemPath, userPath, result
Set shell = CreateObject("WScript.Shell")
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
processPath = shell.Environment("PROCESS")("PATH")
result = registry.GetExpandedStringValue(&H80000002, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Path", systemPath)
If result = 0 Then
    WScript.Echo "System registry PATH length=" & Len(systemPath)
Else
    WScript.Echo "System registry PATH unavailable; WMI return code: " & result
End If
result = registry.GetStringValue(&H80000001, "Environment", "Path", userPath)
If result = 0 Then
    WScript.Echo "User registry PATH length=" & Len(userPath)
Else
    WScript.Echo "User registry PATH is not set or unavailable; WMI return code: " & result
End If
WScript.Echo "Current process PATH length=" & Len(processPath)
WScript.Echo "This is a diagnostic comparison only; no environment broadcast was sent."
