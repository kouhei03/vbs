' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: 現在のプロセスPATHと永続設定の値を比較するオブジェクトを宣言します。
Dim shell, registry, processPath, systemPath, userPath, result
Set shell = CreateObject("WScript.Shell")
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
processPath = shell.Environment("PROCESS")("PATH")
' 解説: システムとユーザーのレジストリPATHを読み、現在のプロセス値との違いを診断します。
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
