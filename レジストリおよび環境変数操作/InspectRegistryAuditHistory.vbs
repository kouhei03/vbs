' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: Windowsイベントログ照会を実行し、出力ストリームを順に読み取ります。
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
' 解説: Securityログから直近20件のレジストリ値変更監査イベントを読み取ります。
Set process = shell.Exec("wevtutil.exe qe Security /q:" & Chr(34) & "*[System[(EventID=4657)]]" & Chr(34) & " /c:20 /rd:true /f:text")
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then
    WScript.Echo "The Security log query failed; check audit configuration and permissions."
    WScript.Quit process.ExitCode
End If
WScript.Echo "An empty result does not prove that registry auditing is configured."
