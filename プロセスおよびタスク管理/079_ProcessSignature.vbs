' 解説: 指定プロセスの実行ファイルに対するAuthenticode署名状態をローカルで検証します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, records, process, processName, validator, shell, commandLine, psCommand, line, runner
Dim service, records, process, processName, validator, shell, commandLine, psCommand, line, runner
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 079_ProcessSignature.vbs <process.exe>": WScript.Quit 2
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
Set records = service.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name = '" & processName & "'")
' 解説: 条件を判定し、必要な分岐を選択します: If records.Count = 0
If records.Count = 0 Then WScript.Echo "Process not found.": WScript.Quit 1
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: コレクション内の各項目を順番に処理します: For Each process In records
For Each process In records
    ' 解説: 条件を判定し、必要な分岐を選択します: If IsNull(process.ExecutablePath) Or Len(process.ExecutablePath) = 0
    If IsNull(process.ExecutablePath) Or Len(process.ExecutablePath) = 0 Then
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo "Executable path unavailable."
    ' 解説: 先行条件が成立しない場合の代替処理へ進みます。
    Else
        ' 解説: 取得値または計算結果を変数に格納します。
        psCommand = "Get-AuthenticodeSignature -LiteralPath '" & Replace(process.ExecutablePath, "'", "''") & "' | Select-Object Status,@{Name='Publisher';Expression={$_.SignerCertificate.Subject}} | Format-List"
        ' 解説: 取得値または計算結果を変数に格納します。
        commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & psCommand & Chr(34)
        ' 解説: 固定されたローカル照会コマンドを実行します。
        Set runner = shell.Exec(commandLine)
        ' 解説: 条件が成立している間、繰り返します: Do While Not runner.StdOut.AtEndOfStream
        Do While Not runner.StdOut.AtEndOfStream
            ' 解説: 取得値または計算結果を変数に格納します。
            line = runner.StdOut.ReadLine
            ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
            WScript.Echo line
        ' 解説: 現在のDoループ反復を終了します。
        Loop
        ' 解説: 条件が成立している間、繰り返します: Do While Not runner.StdErr.AtEndOfStream
        Do While Not runner.StdErr.AtEndOfStream
            ' 解説: 取得値または計算結果を変数に格納します。
            line = runner.StdErr.ReadLine
            ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
            WScript.Echo "ERROR: " & line
        ' 解説: 現在のDoループ反復を終了します。
        Loop
        ' 解説: 条件を判定し、必要な分岐を選択します: If runner.ExitCode <> 0
        If runner.ExitCode <> 0 Then WScript.Echo "Signature query exit code=" & runner.ExitCode
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
