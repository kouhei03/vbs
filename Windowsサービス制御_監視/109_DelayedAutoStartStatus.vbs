' ファイル概要: 遅延自動開始レジストリ値の読み取り。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 109_DelayedAutoStartStatus.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 遅延自動開始レジストリ値の読み取り処理の結果やエラーをコンソールに表示します。

' 指定サービスのDelayedAutoStartレジストリ値を読み取り専用で表示します。
Option Explicit
Dim serviceName, validator, shell, registryPath, delayedValue
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 109_DelayedAutoStartStatus.vbs <service-name>"
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
Set shell = CreateObject("WScript.Shell")
registryPath = "HKLM\SYSTEM\CurrentControlSet\Services" & serviceName & "\DelayedAutoStart"
On Error Resume Next
' [処理解説] レジストリの現在値を読み取ります。設定値は変更しません。
delayedValue = shell.RegRead(registryPath)
If Err.Number <> 0 Then
    WScript.Echo "DelayedAutoStart value is absent or unreadable; this alone does not establish startup mode."
    Err.Clear
Else
    WScript.Echo "DelayedAutoStart=" & delayedValue & " (1 usually indicates delayed automatic start)"
End If
On Error GoTo 0
