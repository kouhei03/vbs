' 解説: 指定プロセスのI/O読み書き量をローカルWMIから読み取り、対象プロセスだけを表示します。プロセスは変更しません。
Option Explicit
' 解説: WMI接続、引数、照会結果、およびプロパティ表示に使う変数を宣言します。
Dim service, processName, validator, records, process, fieldName, value
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 075_ProcessIoBytes.vbs <process.exe>"
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: プロセス名を英数字等に制限し、WQL条件への不正な入力混入を防ぎます。
processName = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 条件を判定し、必要な分岐を選択します: If Not validator.Test(processName)
If Not validator.Test(processName) Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Use a simple executable name such as notepad.exe."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: ローカルWMIへ接続し、名前を限定してプロセス情報を照会します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set records = service.ExecQuery("SELECT Name, ProcessId, ReadOperationCount, WriteOperationCount, ReadTransferCount, WriteTransferCount FROM Win32_Process WHERE Name = '" & processName & "'")
' 解説: 条件を判定し、必要な分岐を選択します: If records.Count = 0
If records.Count = 0 Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Process not found: " & processName
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 各一致プロセスのPIDと要求されたプロパティを表示します。
For Each process In records
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Name=" & process.Name & " PID=" & process.ProcessId
    ' 解説: コレクション内の各項目を順番に処理します: For Each fieldName In Array({', '.join(q(p) for p in props.split(','))})
    For Each fieldName In Array({', '.join(q(p) for p in props.split(','))})
        ' 解説: 取得値または計算結果を変数に格納します。
        value = process.Properties_.Item(fieldName).Value
        ' 解説: 条件を判定し、必要な分岐を選択します: If IsNull(value)
        If IsNull(value) Then
            ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
            WScript.Echo fieldName & "=(not reported)"
        ' 解説: 先行条件が成立しない場合の代替処理へ進みます。
        Else
            ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
            WScript.Echo fieldName & "=" & value
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: 現在の反復を終え、次の要素へ進みます。
    Next
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
