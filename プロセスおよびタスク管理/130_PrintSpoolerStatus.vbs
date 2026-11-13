' 解説: Print Spoolerサービスに対応するローカルプロセス状態を読み取ります。サービスは操作しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, items, item
Dim service, items, item
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set items = service.ExecQuery("SELECT Name,State,ProcessId,StartMode FROM Win32_Service WHERE Name='Spooler'")
' 解説: 条件を判定し、必要な分岐を選択します: If items.Count = 0
If items.Count = 0 Then WScript.Echo "Spooler service not found.": WScript.Quit 1
' 解説: コレクション内の各項目を順番に処理します: For Each item In items
For Each item In items
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "State=" & item.State & " PID=" & item.ProcessId & " StartMode=" & item.StartMode
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
