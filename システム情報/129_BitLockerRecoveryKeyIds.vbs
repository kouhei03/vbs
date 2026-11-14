' ファイル概要: BitLocker回復キーID。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 129_BitLockerRecoveryKeyIds.vbs
' 実行結果: BitLocker回復キーID処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, process, line, commandLine
Set shell = CreateObject("WScript.Shell")
commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & "Get-BitLockerVolume | ForEach-Object { $mount = $_.MountPoint; foreach ($kp in $_.KeyProtector) { if ($kp.KeyProtectorType -eq 'RecoveryPassword') { [pscustomobject]@{ MountPoint = $mount; KeyProtectorId = $kp.KeyProtectorId } } } } | Format-Table -AutoSize" & Chr(34)
Set process = shell.Exec(commandLine)
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    WScript.Echo line
Loop
Do While Not process.StdErr.AtEndOfStream
    line = process.StdErr.ReadLine
    WScript.Echo "ERROR: " & line
Loop
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
