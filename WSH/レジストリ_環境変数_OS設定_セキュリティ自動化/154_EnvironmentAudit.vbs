' ユーザー、システム、プロセス環境変数を表示します。設定変更はしません。
Option Explicit
Dim shell, processEnv, userEnv, systemEnv, item
Set shell = CreateObject("WScript.Shell")
Set processEnv = shell.Environment("PROCESS")
Set userEnv = shell.Environment("USER")
Set systemEnv = shell.Environment("SYSTEM")
WScript.Echo "=== USER ==="
For Each item In userEnv
    WScript.Echo RedactSensitive(item)
Next
WScript.Echo "=== SYSTEM ==="
For Each item In systemEnv
    WScript.Echo RedactSensitive(item)
Next
WScript.Echo "=== PROCESS ==="
For Each item In processEnv
    WScript.Echo RedactSensitive(item)
Next
Function RedactSensitive(entry)
    Dim separator, name
    separator = InStr(entry, "=")
    If separator < 2 Then
        RedactSensitive = "(unrecognized environment entry)"
        Exit Function
    End If
    name = Left(entry, separator - 1)
    If InStr(1, name, "PASSWORD", 1) > 0 Or InStr(1, name, "SECRET", 1) > 0 Or _
        InStr(1, name, "TOKEN", 1) > 0 Or InStr(1, name, "CREDENTIAL", 1) > 0 Or _
        InStr(1, name, "API_KEY", 1) > 0 Then
        RedactSensitive = name & "=(redacted)"
    Else
        RedactSensitive = entry
    End If
End Function
