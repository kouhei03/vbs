' ファイル概要: プロセス実行ファイルのAuthenticode署名状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 079_ProcessSignature.vbs <process.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: プロセス実行ファイルのAuthenticode署名状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, process, processName, validator, shell, commandLine, psCommand, line, runner
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 079_ProcessSignature.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name = '" & processName & "'")
If records.Count = 0 Then WScript.Echo "Process not found.": WScript.Quit 1
Set shell = CreateObject("WScript.Shell")
For Each process In records
    If IsNull(process.ExecutablePath) Or Len(process.ExecutablePath) = 0 Then
        WScript.Echo "Executable path unavailable."
    Else
        psCommand = "Get-AuthenticodeSignature -LiteralPath '" & Replace(process.ExecutablePath, "'", "''") & "' | Select-Object Status,@{Name='Publisher';Expression={$_.SignerCertificate.Subject}} | Format-List"
        commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & psCommand & Chr(34)
        Set runner = shell.Exec(commandLine)
        Do While Not runner.StdOut.AtEndOfStream
            line = runner.StdOut.ReadLine
            WScript.Echo line
        Loop
        Do While Not runner.StdErr.AtEndOfStream
            line = runner.StdErr.ReadLine
            WScript.Echo "ERROR: " & line
        Loop
        If runner.ExitCode <> 0 Then WScript.Echo "Signature query exit code=" & runner.ExitCode
    End If
Next
