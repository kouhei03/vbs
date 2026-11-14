' 概要: デモ専用HKCU値が存在するか検証します。値の作成・修復はしません。
' 実行方法: cscript //nologo 185_DemoRegistryVerification.vbs
' 処理結果: デモ用レジストリ値の事後確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\WSH\SafeExamples\Demo\ExampleValue"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Demo value not present."
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "Demo value is present and readable."
