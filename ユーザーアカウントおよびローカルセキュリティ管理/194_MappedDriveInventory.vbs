' ファイル概要: 割り当て済みネットワークドライブを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 194_MappedDriveInventory.vbs
' 実行結果: 割り当て済みネットワークドライブを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, drives, drive
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drives = service.ExecQuery( _
    "SELECT DeviceID, ProviderName, FreeSpace, Size FROM Win32_LogicalDisk WHERE DriveType=4")
WScript.Echo "Drive" & vbTab & "Provider" & vbTab & "FreeBytes" & vbTab & "SizeBytes"
For Each drive In drives
    WScript.Echo drive.DeviceID & vbTab & drive.ProviderName & vbTab & _
        drive.FreeSpace & vbTab & drive.Size
Next
