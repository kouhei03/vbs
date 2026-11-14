' ファイル概要: CPUのモデル・クロック・コア数を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetCPUInfo.vbs
' 実行結果: CPUのモデル・クロック・コア数を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processors, processor
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processors = service.ExecQuery( _
    "SELECT Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors " & _
    "FROM Win32_Processor")

For Each processor In processors
    WScript.Echo "Model: " & processor.Name
    WScript.Echo "Maximum clock speed (MHz): " & processor.MaxClockSpeed
    WScript.Echo "Cores: " & processor.NumberOfCores
    WScript.Echo "Logical processors: " & processor.NumberOfLogicalProcessors
    WScript.Echo ""
Next
