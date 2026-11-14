' 概要: サンプル専用キーの存在を確認します。キー・値の削除は行いません。
' 実行方法: cscript //nologo 153_RegistryCleanupPreview.vbs
' 処理結果: デモキーの削除候補表示（削除なし）。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, value
Set shell = CreateObject("WScript.Shell")
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead("HKEY_CURRENT_USER\Software\WSH\SafeExamples\Marker")
If Err.Number = 0 Then
    WScript.Echo "Cleanup candidate: HKCU\Software\WSH\SafeExamples\Marker"
Else
    WScript.Echo "No sample cleanup candidate found."
End If
On Error GoTo 0
WScript.Echo "Preview only; registry data was not deleted."
