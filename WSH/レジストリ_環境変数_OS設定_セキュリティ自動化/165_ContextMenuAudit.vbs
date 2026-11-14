' 概要: 独自のHKCUコンテキストメニュー項目の有無だけを確認します。項目は追加しません。
' 実行方法: cscript //nologo 165_ContextMenuAudit.vbs
' 処理結果: ユーザーコンテキストメニュー値の照会。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, roots, root, value, found
Set shell = CreateObject("WScript.Shell")
roots = Array( _
    "HKEY_CURRENT_USER\Software\Classes\*\shell\", _
    "HKEY_CURRENT_USER\Software\Classes\Directory\shell\", _
    "HKEY_CURRENT_USER\Software\Classes\Directory\Background\shell\")
found = False
For Each root In roots
    On Error Resume Next
    ' 主処理: 対象のレジストリ値を読み取ります。
    value = shell.RegRead(root)
    If Err.Number = 0 Then
        WScript.Echo "Default item at " & root & ": " & value
        found = True
    End If
    Err.Clear
    On Error GoTo 0
Next
If Not found Then WScript.Echo "No default context-menu labels reported at the checked paths."
WScript.Echo "Read-only audit."
