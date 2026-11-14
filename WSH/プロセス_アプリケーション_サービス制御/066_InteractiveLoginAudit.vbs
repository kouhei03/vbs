' 概要: 対話ログオン中のユーザーを照会します（セッション変更なし）。
' 実行方法: cscript //nologo 066_InteractiveLoginAudit.vbs
' 処理結果: 対話ログオンユーザー確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, systems, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set systems = service.ExecQuery("SELECT Name, UserName FROM Win32_ComputerSystem")
For Each item In systems
    If IsNull(item.UserName) Or item.UserName = "" Then
        WScript.Echo item.Name & ": no interactive user reported"
    Else
        WScript.Echo item.Name & ": " & item.UserName
    End If
Next
