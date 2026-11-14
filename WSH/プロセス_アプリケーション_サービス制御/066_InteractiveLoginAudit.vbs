' 対話ログオン中のユーザーを照会します（セッション変更なし）。
Option Explicit
Dim service, systems, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT Name, UserName FROM Win32_ComputerSystem")
For Each item In systems
    If IsNull(item.UserName) Or item.UserName = "" Then
        WScript.Echo item.Name & ": no interactive user reported"
    Else
        WScript.Echo item.Name & ": " & item.UserName
    End If
Next
