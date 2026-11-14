' ファイル概要: メモリフォームファクター。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 060_MemoryFormFactor.vbs
' 実行結果: メモリフォームファクター処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, modules, module, formName
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PhysicalMemory を照会し、DeviceLocator, FormFactor を取得します。
Set modules = service.ExecQuery("SELECT DeviceLocator, FormFactor FROM Win32_PhysicalMemory")
For Each module In modules
    Select Case module.FormFactor
        Case 8: formName = "DIMM"
        Case 12: formName = "SODIMM"
        Case 9: formName = "RIMM"
        Case 13: formName = "SRIMM"
        Case Else: formName = "Other/unknown code " & module.FormFactor
    End Select
    WScript.Echo module.DeviceLocator & ": " & formName
Next
