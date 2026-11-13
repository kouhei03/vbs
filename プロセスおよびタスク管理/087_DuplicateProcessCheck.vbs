' 解説: 指定した実行ファイル名の同時起動数を数え、重複状態を報告します。起動のブロックはしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, records, processName, validator
Dim service, records, processName, validator
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 087_DuplicateProcessCheck.vbs <process.exe>": WScript.Quit 2
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
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Running instances: " & records.Count
' 解説: 条件を判定し、必要な分岐を選択します: If records.Count > 1
If records.Count > 1 Then WScript.Echo "Duplicate instances detected; no process was blocked or terminated."
