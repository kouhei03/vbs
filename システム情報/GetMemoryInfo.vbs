' ファイル概要: 物理メモリの総量と空き容量を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetMemoryInfo.vbs
' 実行結果: 物理メモリの総量と空き容量を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, systems, system
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery( _
    "SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem")

For Each system In systems
    WScript.Echo "Total physical memory (bytes): " & _
        (CDbl(system.TotalVisibleMemorySize) * 1024)
    WScript.Echo "Free physical memory (bytes): " & _
        (CDbl(system.FreePhysicalMemory) * 1024)
Next
