' 概要: 指定拡張子の現在のUserChoice ProgIdを照会します。関連付けは変更しません。
' 実行方法: Usage: cscript //nologo 164_FileAssociationAudit.vbs <extension-without-dot>
' 処理結果: 拡張子のUserChoice ProgId照会。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim extension, re, shell, path, value
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 164_FileAssociationAudit.vbs <extension-without-dot>"
    WScript.Quit 2
End If
extension = LCase(WScript.Arguments(0))
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9]{1,10}$"
If Not re.Test(extension) Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\." & extension & "\UserChoice\ProgId"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "UserChoice ProgId not found or protected by Windows."
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "." & extension & " -> " & value
