' 解説: ローカルPC上のRemote Desktop Connectionクライアント状態を照会します。接続先は指定/起動しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, processes, process
Dim service, processes, process
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes = service.ExecQuery("SELECT Name,ProcessId,CreationDate FROM Win32_Process WHERE Name='mstsc.exe'")
' 解説: 条件を判定し、必要な分岐を選択します: If processes.Count = 0
If processes.Count = 0 Then WScript.Echo "mstsc.exe is not running."
' 解説: コレクション内の各項目を順番に処理します: For Each process In processes
For Each process In processes
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo process.Name & " PID=" & process.ProcessId & " Created=" & process.CreationDate
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
