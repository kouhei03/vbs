' 解説: BitLocker回復キーIDをWindowsの読み取り専用PowerShellコマンドレットで照会します。
Option Explicit
' 解説: PowerShellプロセスと標準出力・標準エラーを扱う変数を宣言します。
Dim shell, process, line, commandLine
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 固定された照会式を一つのコマンド引数として組み立て、設定変更を含まずに実行します。
commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & "Get-BitLockerVolume | ForEach-Object { $mount = $_.MountPoint; foreach ($kp in $_.KeyProtector) { if ($kp.KeyProtectorType -eq 'RecoveryPassword') { [pscustomobject]@{ MountPoint = $mount; KeyProtectorId = $kp.KeyProtectorId } } } } | Format-Table -AutoSize" & Chr(34)
' 解説: 読み取り用コマンドを実行し、結果取得用オブジェクトを設定します。
Set process = shell.Exec(commandLine)
' 解説: 取得結果とエラーをそれぞれ表示します。
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 取得値または計算結果を変数へ格納します。
    line = process.StdOut.ReadLine
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo line
' 解説: Doループの現在の反復を終了します。
Loop
' 解説: 条件が成立している間、繰り返し処理します: Do While Not process.StdErr.AtEndOfStream
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 取得値または計算結果を変数へ格納します。
    line = process.StdErr.ReadLine
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "ERROR: " & line
' 解説: Doループの現在の反復を終了します。
Loop
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: process.ExitCode <> 0
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
