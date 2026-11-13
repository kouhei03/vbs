' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process、processName、found、matchCount）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報、対象プロセス名、対象が見つかったかを示すフラグ。
Dim service, processes, process, processName, found, matchCount

' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo TestProcessRunning.vbs <process-name>"
    WScript.Echo "Usage: cscript //nologo TestProcessRunning.vbs <process-name>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: processName = Trim(WScript.Arguments(0))
processName = Trim(WScript.Arguments(0))
' 解説: 条件「Len(processName) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(processName) = 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name must not be empty."
    WScript.Echo "Process name must not be empty."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、Name, ProcessIdを取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: matchCount = 0
matchCount = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 条件「StrComp(process.Name, processName, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If StrComp(process.Name, processName, vbTextCompare) = 0 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
        found = True
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: matchCount = matchCount + 1
        matchCount = matchCount + 1
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & " is running (PID " & process.ProcessId & ")."
        WScript.Echo process.Name & " is running (PID " & process.ProcessId & ")."
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 条件「Not found Then」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: processName & " is not running."
    WScript.Echo processName & " is not running."
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Matching process count: " & matchCount
WScript.Echo "Matching process count: " & matchCount
