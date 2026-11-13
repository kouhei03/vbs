' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process、count、totalWorkingSet）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報、一致件数。
Dim service, processes, process, count, totalWorkingSet
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、ProcessId, WorkingSetSizeを取得します。
Set processes = service.ExecQuery( _
    "SELECT ProcessId, WorkingSetSize FROM Win32_Process")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count = 0
count = 0
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: totalWorkingSet = 0
totalWorkingSet = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: count = count + 1
    count = count + 1
    ' 解説: 条件「Not IsNull(process.WorkingSetSize) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If Not IsNull(process.WorkingSetSize) Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: totalWorkingSet = totalWorkingSet + CDbl(process.WorkingSetSize)
        totalWorkingSet = totalWorkingSet + CDbl(process.WorkingSetSize)
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process count: " & count
WScript.Echo "Process count: " & count
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Combined working set (bytes): " & FormatNumber(totalWorkingSet, 0)
WScript.Echo "Combined working set (bytes): " & FormatNumber(totalWorkingSet, 0)
