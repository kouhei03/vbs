' 概要: UACのEnableLUA値を照会します。通知レベルやUACは変更しません。
' 実行方法: cscript //nologo 167_UacPolicyAudit.vbs
' 処理結果: UAC EnableLUA値の照会。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "EnableLUA unavailable: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "EnableLUA=" & value & " (read-only)"
