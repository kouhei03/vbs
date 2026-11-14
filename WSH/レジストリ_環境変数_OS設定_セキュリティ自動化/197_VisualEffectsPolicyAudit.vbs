' 概要: 視覚効果プリセットの値だけを読み取ります。デザイン設定は変更しません。
' 実行方法: cscript //nologo 197_VisualEffectsPolicyAudit.vbs
' 処理結果: 視覚効果プリセット値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects\VisualFXSetting"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "VisualFXSetting=(Windows default or not set)"
Else
    WScript.Echo "VisualFXSetting=" & value & " (read-only)"
End If
On Error GoTo 0
