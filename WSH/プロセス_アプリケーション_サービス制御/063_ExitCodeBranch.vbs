' 固定コマンドの終了コードを取得し、成功／失敗の分岐例を表示します。
Option Explicit
Dim shell, exitCode
Set shell = CreateObject("WScript.Shell")
exitCode = shell.Run("cmd.exe /c ver", 1, True)
If exitCode = 0 Then
    WScript.Echo "Success: Windows version query completed."
Else
    WScript.Echo "Failure: fixed command returned " & exitCode
    WScript.Quit exitCode
End If
