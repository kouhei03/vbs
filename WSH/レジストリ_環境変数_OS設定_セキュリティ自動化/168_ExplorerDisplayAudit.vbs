' 概要: Explorerの隠しファイル・拡張子表示設定を読み取ります。設定変更はしません。
' 実行方法: cscript //nologo 168_ExplorerDisplayAudit.vbs
' 処理結果: 隠しファイル・拡張子表示設定確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, names, name, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\"
names = Array("Hidden", "HideFileExt", "ShowSuperHidden")
For Each name In names
    On Error Resume Next
    ' 主処理: 対象のレジストリ値を読み取ります。
    value = shell.RegRead(path & name)
    If Err.Number = 0 Then
        WScript.Echo name & "=" & value
    Else
        WScript.Echo name & "=(not set)"
    End If
    Err.Clear
    On Error GoTo 0
Next
