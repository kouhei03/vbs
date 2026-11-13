' 解説: ローカルプロセス名/PIDのHTML一覧を新規ファイルに出力します。既存ファイルは上書きしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, service, processes, process, report, path
Dim fso, service, processes, process, report, path
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count<>1
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 170_ProcessDashboard.vbs <new-html-path>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
path=WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso=CreateObject("Scripting.FileSystemObject")
' 解説: 条件を判定し、必要な分岐を選択します: If fso.FileExists(path)
If fso.FileExists(path) Then WScript.Echo "Refusing to overwrite existing file.": WScript.Quit 1
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service=GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes=service.ExecQuery("SELECT Name,ProcessId FROM Win32_Process")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set report=fso.CreateTextFile(path,False,True)
' 解説: 取得値または計算結果を変数に格納します。
report.WriteLine "<!doctype html><html><head><meta charset=""utf-8""><title>Process inventory</title></head><body>"
' 解説: この文でローカル診断処理を進めます: report.WriteLine "<h1>Local process inventory</h1><p>Generated " & Html(Now) & "</p><table><tr><th>Name</th><th>PID</th></tr>"
report.WriteLine "<h1>Local process inventory</h1><p>Generated " & Html(Now) & "</p><table><tr><th>Name</th><th>PID</th></tr>"
' 解説: コレクション内の各項目を順番に処理します: For Each process In processes
For Each process In processes
    ' 解説: この文でローカル診断処理を進めます: report.WriteLine "<tr><td>" & Html(process.Name) & "</td><td>" & process.ProcessId & "</td></tr>"
    report.WriteLine "<tr><td>" & Html(process.Name) & "</td><td>" & process.ProcessId & "</td></tr>"
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: この文でローカル診断処理を進めます: report.WriteLine "</table></body></html>"
report.WriteLine "</table></body></html>"
' 解説: この文でローカル診断処理を進めます: report.Close
report.Close
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Dashboard written: " & path

' 解説: 再利用する補助関数を定義します: Function Html(value)
Function Html(value)
    ' 解説: このスクリプトで使う変数を宣言します: text
    Dim text
    ' 解説: 取得値または計算結果を変数に格納します。
    text=CStr(value)
    ' 解説: 取得値または計算結果を変数に格納します。
    text=Replace(text,"&","&amp;")
    ' 解説: 取得値または計算結果を変数に格納します。
    text=Replace(text,"<","&lt;")
    ' 解説: 取得値または計算結果を変数に格納します。
    text=Replace(text,">","&gt;")
    ' 解説: 取得値または計算結果を変数に格納します。
    text=Replace(text,Chr(34),"&quot;")
    ' 解説: 取得値または計算結果を変数に格納します。
    Html=text
' 解説: 補助関数の定義を終了します。
End Function
