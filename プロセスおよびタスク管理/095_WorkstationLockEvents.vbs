' 解説: Securityログから直近のワークステーションWorkstation lockイベントを読み取ります。ログは変更しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, events, eventRecord, query, count
Dim service, events, eventRecord, query, count
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 取得値または計算結果を変数に格納します。
query = "SELECT EventCode,TimeGenerated,RecordNumber,SourceName FROM Win32_NTLogEvent WHERE Logfile='Security' AND EventCode='4800'"
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set events = service.ExecQuery(query)
' 解説: 取得値または計算結果を変数に格納します。
count = 0
' 解説: コレクション内の各項目を順番に処理します: For Each eventRecord In events
For Each eventRecord In events
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Event=" & eventRecord.EventCode & " Time=" & eventRecord.TimeGenerated & " Record=" & eventRecord.RecordNumber
    ' 解説: 取得値または計算結果を変数に格納します。
    count = count + 1
    ' 解説: 条件を判定し、必要な分岐を選択します: If count >= 20
    If count >= 20 Then Exit For
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を判定し、必要な分岐を選択します: If count = 0
If count = 0 Then WScript.Echo "No event returned; log permissions/audit policy may limit results."
