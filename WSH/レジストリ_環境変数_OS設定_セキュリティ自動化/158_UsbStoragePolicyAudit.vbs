' 概要: USBSTORサービスの起動値を照会します。ポリシー変更・USB履歴列挙はしません。
' 実行方法: cscript //nologo 158_UsbStoragePolicyAudit.vbs
' 処理結果: USBSTORポリシー値の確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, value, path
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\USBSTOR\Start"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "USBSTOR policy unavailable: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "USBSTOR Start=" & value & " (read-only)"
