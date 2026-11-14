' WMIに公開されるログオン時スタートアップ項目を監査します（変更なし）。
Option Explicit
Dim service, entries, entry
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set entries = service.ExecQuery("SELECT Name, Command, Location, User FROM Win32_StartupCommand")
For Each entry In entries
    WScript.Echo entry.Name & vbTab & entry.User & vbTab & entry.Location & vbTab & entry.Command
Next
