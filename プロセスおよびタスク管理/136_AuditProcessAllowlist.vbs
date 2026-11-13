' 解説: 読み取り専用の許可リストと実行中プロセス名を比較し、未登録名を報告します。起動ブロックはしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, allowFile, allowed, line, service, processes, process, found
Dim fso, allowFile, allowed, line, service, processes, process, found
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 136_AuditProcessAllowlist.vbs <allowlist-file>": WScript.Quit 2
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso = CreateObject("Scripting.FileSystemObject")
' 解説: 条件を判定し、必要な分岐を選択します: If Not fso.FileExists(WScript.Arguments(0))
If Not fso.FileExists(WScript.Arguments(0)) Then WScript.Echo "Allowlist file not found.": WScript.Quit 1
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set allowed = CreateObject("Scripting.Dictionary")
' 解説: 取得値または計算結果を変数に格納します。
allowed.CompareMode = vbTextCompare
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set allowFile = fso.OpenTextFile(WScript.Arguments(0),1,False)
' 解説: 条件が成立している間、繰り返します: Do Until allowFile.AtEndOfStream
Do Until allowFile.AtEndOfStream
    ' 解説: 取得値または計算結果を変数に格納します。
    line = Trim(allowFile.ReadLine)
    ' 解説: 条件を判定し、必要な分岐を選択します: If Len(line)>0
    If Len(line)>0 Then allowed(line)=True
' 解説: 現在のDoループ反復を終了します。
Loop
' 解説: この文でローカル診断処理を進めます: allowFile.Close
allowFile.Close
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes = service.ExecQuery("SELECT Name,ProcessId FROM Win32_Process")
' 解説: コレクション内の各項目を順番に処理します: For Each process In processes
For Each process In processes
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not allowed.Exists(process.Name)
    If Not allowed.Exists(process.Name) Then WScript.Echo "Not in allowlist: " & process.Name & " PID=" & process.ProcessId
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Audit only; no process was blocked."
