' 概要: DNSリゾルバーキャッシュを表示する固定コマンド例です。消去はしません。
' 実行方法: cscript //nologo 005_DnsCacheSummary.vbs
' 処理結果: DNSキャッシュの件数概要（詳細表示・消去なし）。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim shell, runner, line, count
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("ipconfig.exe /displaydns")
count = 0
Do While Not runner.StdOut.AtEndOfStream
    line = runner.StdOut.ReadLine
    If InStr(1, line, "Record Name", 1) > 0 Then count = count + 1
Loop
Do While Not runner.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & runner.StdErr.ReadLine
Loop
If runner.ExitCode <> 0 Then WScript.Quit runner.ExitCode
WScript.Echo "DNS cached-record sections=" & count & " (record details omitted)."
