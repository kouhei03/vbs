' ファイル概要: 物理ディスクの型番・種類・容量を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetPhysicalDisks.vbs
' 実行結果: 物理ディスクの型番・種類・容量を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, disks, disk
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set disks = service.ExecQuery( _
    "SELECT Model, Size, MediaType FROM Win32_DiskDrive")

For Each disk In disks
    WScript.Echo "Model: " & disk.Model
    WScript.Echo "Media type: " & disk.MediaType
    WScript.Echo "Capacity (bytes): " & disk.Size
    WScript.Echo ""
Next
