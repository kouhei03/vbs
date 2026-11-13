' 解説: ローカルプロセスの親PIDをたどり、親子関係をテキスト形式で表示します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, records, process, names, parents, pid, parentPid, chain, depth
Dim service, records, process, names, parents, pid, parentPid, chain, depth
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set records = service.ExecQuery("SELECT Name,ProcessId,ParentProcessId FROM Win32_Process")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set names = CreateObject("Scripting.Dictionary")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set parents = CreateObject("Scripting.Dictionary")
' 解説: コレクション内の各項目を順番に処理します: For Each process In records
For Each process In records
    ' 解説: 取得値または計算結果を変数に格納します。
    names(CStr(process.ProcessId)) = process.Name
    ' 解説: 取得値または計算結果を変数に格納します。
    parents(CStr(process.ProcessId)) = CStr(process.ParentProcessId)
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: コレクション内の各項目を順番に処理します: For Each process In records
For Each process In records
    ' 解説: 取得値または計算結果を変数に格納します。
    pid = CStr(process.ProcessId)
    ' 解説: 取得値または計算結果を変数に格納します。
    parentPid = parents(pid)
    ' 解説: 取得値または計算結果を変数に格納します。
    chain = process.Name & " (PID " & pid & ")"
    ' 解説: 取得値または計算結果を変数に格納します。
    depth = 0
    ' 解説: 条件が成立している間、繰り返します: Do While names.Exists(parentPid) And depth < 16
    Do While names.Exists(parentPid) And depth < 16
        ' 解説: 取得値または計算結果を変数に格納します。
        chain = names(parentPid) & " (PID " & parentPid & ") > " & chain
        ' 解説: 取得値または計算結果を変数に格納します。
        parentPid = parents(parentPid)
        ' 解説: 取得値または計算結果を変数に格納します。
        depth = depth + 1
    ' 解説: 現在のDoループ反復を終了します。
    Loop
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo chain
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
