' ファイル概要: 論理ドライブとその種類を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetLogicalDriveTypes.vbs
' 実行結果: 論理ドライブとその種類を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, drives, drive, driveType
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set drives = service.ExecQuery( _
    "SELECT DeviceID, DriveType FROM Win32_LogicalDisk")

For Each drive In drives
    Select Case drive.DriveType
        Case 0: driveType = "Unknown"
        Case 1: driveType = "No root directory"
        Case 2: driveType = "Removable"
        Case 3: driveType = "Fixed"
        Case 4: driveType = "Network"
        Case 5: driveType = "CD-ROM"
        Case 6: driveType = "RAM disk"
        Case Else: driveType = "Other"
    End Select
    WScript.Echo drive.DeviceID & ": " & driveType
Next
