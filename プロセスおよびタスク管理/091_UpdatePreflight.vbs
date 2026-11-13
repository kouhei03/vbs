' 解説: 指定したアプリ名が実行中かを更新前確認として報告します。終了操作はしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, processName, validator, records, i
Dim service, processName, validator, records, i
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count < 1
If WScript.Arguments.Count < 1 Then WScript.Echo "Usage: cscript //nologo 091_UpdatePreflight.vbs <process.exe> [process.exe ...]": WScript.Quit 2
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定範囲の要素を順番に処理します: For i = 0 To WScript.Arguments.Count - 1
For i = 0 To WScript.Arguments.Count - 1
    ' 解説: 取得値または計算結果を変数に格納します。
    processName = WScript.Arguments(i)
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not validator.Test(processName)
    If Not validator.Test(processName) Then WScript.Echo "Invalid name: " & processName: WScript.Quit 2
    ' 解説: 指定条件に合う情報を読み取り専用で照会します。
    Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo processName & " running instances=" & records.Count
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Preflight only; no application was closed."
