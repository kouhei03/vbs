' ファイル概要: サービス実行ファイルのローカル署名確認。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 170_ServiceBinarySignatureStatus.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サービス実行ファイルのローカル署名確認処理の結果やエラーをコンソールに表示します。

' 指定サービスの実行ファイル署名状態をローカルで検証します。サービス設定は変更しません。
Option Explicit
Dim serviceName, validator, service, services, item, found, rawPath, executablePath, exePosition, quotePosition
Dim shell, psCommand, commandLine, process, line
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 170_ServiceBinarySignatureStatus.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Invalid service name."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, PathName を取得します。
Set services = service.ExecQuery("SELECT Name, PathName FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        found = True
        If IsNull(item.PathName) Then
            rawPath = ""
        Else
            rawPath = Trim(CStr(item.PathName))
        End If
        executablePath = ""
        If Len(rawPath) > 0 And Left(rawPath, 1) = Chr(34) Then
            quotePosition = InStr(2, rawPath, Chr(34))
            If quotePosition > 2 Then executablePath = Mid(rawPath, 2, quotePosition - 2)
        Else
            exePosition = InStr(1, LCase(rawPath), ".exe", vbBinaryCompare)
            If exePosition > 0 Then executablePath = Left(rawPath, exePosition + 3)
        End If
        If Len(executablePath) = 0 Then
            WScript.Echo "Could not extract an executable .exe path from: " & rawPath
            Exit For
        Else
            Set shell = CreateObject("WScript.Shell")
            executablePath = shell.ExpandEnvironmentStrings(executablePath)
            psCommand = "Get-AuthenticodeSignature -LiteralPath '" & Replace(executablePath, "'", "''") & "' | Select-Object Status,@{Name='Publisher';Expression={$_.SignerCertificate.Subject}} | Format-List"
            commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & psCommand & Chr(34)
            ' [処理解説] 外部プロセスを起動し、出力ストリームと終了コードを取得できるオブジェクトを受け取ります。
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
        End If
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found: " & serviceName
    WScript.Quit 3
End If
