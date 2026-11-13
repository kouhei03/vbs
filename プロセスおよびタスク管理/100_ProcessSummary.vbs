' 解説: ローカルプロセス数と合計Working Setを集計し、監視用サマリーを表示します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, processes, process, count, bytes
Dim service, processes, process, count, bytes
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes = service.ExecQuery("SELECT ProcessId,WorkingSetSize FROM Win32_Process")
' 解説: 取得値または計算結果を変数に格納します。
count = 0
' 解説: 取得値または計算結果を変数に格納します。
bytes = CDbl(0)
' 解説: コレクション内の各項目を順番に処理します: For Each process In processes
For Each process In processes
    ' 解説: 取得値または計算結果を変数に格納します。
    count = count + 1
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not IsNull(process.WorkingSetSize)
    If Not IsNull(process.WorkingSetSize) Then bytes = bytes + CDbl(process.WorkingSetSize)
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Process count=" & count
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Total working set bytes=" & FormatNumber(bytes,0)
