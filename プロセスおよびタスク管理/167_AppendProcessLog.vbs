' ファイル概要: プロセス数のローカルログ追記。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 167_AppendProcessLog.vbs <log-file>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: プロセス数のローカルログ追記処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, service, processes, log, path
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 167_AppendProcessLog.vbs <log-file>": WScript.Quit 2
path=WScript.Arguments(0)
Set fso=CreateObject("Scripting.FileSystemObject")
Set service=GetObject("winmgmts:\\.\root\cimv2")
Set processes=service.ExecQuery("SELECT ProcessId FROM Win32_Process")
Set log=fso.OpenTextFile(path,8,True)
log.WriteLine Now & ",process-count=" & processes.Count
log.Close
WScript.Echo "A local summary line was appended."
