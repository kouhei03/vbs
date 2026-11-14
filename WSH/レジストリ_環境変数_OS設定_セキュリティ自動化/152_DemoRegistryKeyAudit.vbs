' 概要: このサンプルが使うHKCUのデモキーの存在だけを確認します。作成はしません。
' 実行方法: cscript //nologo 152_DemoRegistryKeyAudit.vbs
' 処理結果: デモ用レジストリキーの存在確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, value
Set shell = CreateObject("WScript.Shell")
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead("HKEY_CURRENT_USER\Software\WSH\SafeExamples\Marker")
If Err.Number = 0 Then
    WScript.Echo "Demo marker exists: " & value
Else
    WScript.Echo "Demo marker not found."
End If
On Error GoTo 0
WScript.Echo "Read-only check; no key or value was created."
