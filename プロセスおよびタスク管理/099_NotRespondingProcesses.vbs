' ファイル概要: 応答なしプロセス検出のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 099_NotRespondingProcesses.vbs
' 実行結果: 応答なしプロセス検出のみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line, found
Set shell = CreateObject("WScript.Shell")
' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
Set process = shell.Exec("tasklist.exe /v /fo csv")
found = False
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    If InStr(1, line, "Not Responding", vbTextCompare) > 0 Or InStr(line, "応答なし") > 0 Then
        WScript.Echo line
        found = True
    End If
Loop
If Not found Then WScript.Echo "No not-responding process row found."
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
