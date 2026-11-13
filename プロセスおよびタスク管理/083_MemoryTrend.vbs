' 解説: 指定プロセスのWorking Setを最大5回記録し、増加傾向を確認するためのデータを表示します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, records, process, processName, validator, i, found
Dim service, records, process, processName, validator, i, found
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 083_MemoryTrend.vbs <process.exe>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
processName = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 条件を判定し、必要な分岐を選択します: If Not validator.Test(processName)
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定範囲の要素を順番に処理します: For i = 1 To 5
For i = 1 To 5
    ' 解説: 指定条件に合う情報を読み取り専用で照会します。
    Set records = service.ExecQuery("SELECT ProcessId,WorkingSetSize FROM Win32_Process WHERE Name = '" & processName & "'")
    ' 解説: 取得値または計算結果を変数に格納します。
    found = False
    ' 解説: コレクション内の各項目を順番に処理します: For Each process In records
    For Each process In records
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo "Sample=" & i & " PID=" & process.ProcessId & " WorkingSetBytes=" & process.WorkingSetSize
        ' 解説: 取得値または計算結果を変数に格納します。
        found = True
    ' 解説: 現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not found
    If Not found Then WScript.Echo "Process not running at sample " & i
    ' 解説: 条件を判定し、必要な分岐を選択します: If i < 5
    If i < 5 Then WScript.Sleep 2000
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Samples are diagnostic only and do not assert a memory leak."
