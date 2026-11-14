' 概要: 指定したHKCU\Software配下の値を読み取ります。レジストリ変更はしません。
' 実行方法: Usage: cscript //nologo 151_RegistryValueAudit.vbs <HKCU\\Software\\...\\ValueName>
' 処理結果: HKCU\Software配下の値を読み取り。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim path, re, shell, value
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 151_RegistryValueAudit.vbs <HKCU\\Software\\...\\ValueName>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^HKEY_CURRENT_USER\\Software\\[A-Za-z0-9 _.-]+(\\[A-Za-z0-9 _.-]+)*$"
If Not re.Test(path) Then
    WScript.Echo "Only a value under HKCU\\Software is accepted."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Read failed or value does not exist: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo path & "=" & value
