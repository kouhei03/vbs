' 解説: 現在のローカルプロセス数を数え、同時実行数の監視データを表示します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, processes
Dim service, processes
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes = service.ExecQuery("SELECT ProcessId FROM Win32_Process")
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Current local process count=" & processes.Count
