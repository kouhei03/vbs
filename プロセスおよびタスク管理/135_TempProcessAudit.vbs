' 解説: TEMP配下から起動した実行中プロセスを列挙し、確認用情報を表示します。隔離や終了はしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, processes, process, shell, tempPath, pathValue, found
Dim service, processes, process, shell, tempPath, pathValue, found
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set processes = service.ExecQuery("SELECT Name,ProcessId,ExecutablePath,CreationDate FROM Win32_Process")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
tempPath = LCase(shell.ExpandEnvironmentStrings("%TEMP%"))
' 解説: 取得値または計算結果を変数に格納します。
found = False
' 解説: コレクション内の各項目を順番に処理します: For Each process In processes
For Each process In processes
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not IsNull(process.ExecutablePath)
    If Not IsNull(process.ExecutablePath) Then
        ' 解説: 取得値または計算結果を変数に格納します。
        pathValue = LCase(process.ExecutablePath)
        ' 解説: 条件を判定し、必要な分岐を選択します: If Left(pathValue, Len(tempPath)) = tempPath
        If Left(pathValue, Len(tempPath)) = tempPath Then
            ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
            WScript.Echo process.Name & " PID=" & process.ProcessId & " Path=" & process.ExecutablePath
            ' 解説: 取得値または計算結果を変数に格納します。
            found = True
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を判定し、必要な分岐を選択します: If Not found
If Not found Then WScript.Echo "No process image under the current TEMP path was found."
