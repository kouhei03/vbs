' 解説: ローカルプロセス数と時刻だけを指定ログへ追記します。実行ファイル名やコマンドラインは記録しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, service, processes, log, path
Dim fso, service, processes, log, path
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count<>1
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 167_AppendProcessLog.vbs <log-file>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
path=WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso=CreateObject("Scripting.FileSystemObject")
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service=GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes=service.ExecQuery("SELECT ProcessId FROM Win32_Process")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set log=fso.OpenTextFile(path,8,True)
' 解説: 取得値または計算結果を変数に格納します。
log.WriteLine Now & ",process-count=" & processes.Count
' 解説: この文でローカル診断処理を進めます: log.Close
log.Close
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "A local summary line was appended."
