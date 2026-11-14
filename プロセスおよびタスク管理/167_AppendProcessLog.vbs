' ファイル概要: プロセス数のローカルログ追記。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 167_AppendProcessLog.vbs <log-file>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: プロセス数のローカルログ追記処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, service, processes, logFile, path
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 167_AppendProcessLog.vbs <log-file>": WScript.Quit 2
path=WScript.Arguments(0)
Set fso=CreateObject("Scripting.FileSystemObject")
Set service=GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、ProcessId を取得します。
Set processes=service.ExecQuery("SELECT ProcessId FROM Win32_Process")
Set logFile=fso.OpenTextFile(path,8,True)
logFile.WriteLine Now & ",process-count=" & processes.Count
logFile.Close
WScript.Echo "A local summary line was appended."
