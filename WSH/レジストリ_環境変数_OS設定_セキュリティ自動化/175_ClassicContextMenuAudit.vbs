' 概要: Windows 11クラシックコンテキストメニューのユーザー設定有無を確認します。
' 実行方法: cscript //nologo 175_ClassicContextMenuAudit.vbs
' 処理結果: クラシックメニュー上書き値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32\"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Classic context-menu override not found."
Else
    WScript.Echo "Classic context-menu override exists."
End If
On Error GoTo 0
WScript.Echo "Read-only audit."
