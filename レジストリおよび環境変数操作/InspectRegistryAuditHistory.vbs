' ファイル概要: セキュリティログのレジストリ変更監査イベントを照会する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectRegistryAuditHistory.vbs
' 実行結果: セキュリティログのレジストリ変更監査イベントを照会する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line
Set shell = CreateObject("WScript.Shell")
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
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
