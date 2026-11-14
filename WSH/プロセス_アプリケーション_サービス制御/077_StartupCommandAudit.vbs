' 概要: WMIに公開されるログオン時スタートアップ項目を監査します（変更なし）。
' 実行方法: cscript //nologo 077_StartupCommandAudit.vbs
' 処理結果: WMIスタートアップ項目の監査。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, entries, entry
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set entries = service.ExecQuery("SELECT Name, Command, Location, User FROM Win32_StartupCommand")
For Each entry In entries
    WScript.Echo entry.Name & vbTab & entry.User & vbTab & entry.Location & vbTab & entry.Command
Next
