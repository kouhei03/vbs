' ファイル概要: メモリエラー訂正情報。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 062_MemoryEcc.vbs
' 実行結果: メモリエラー訂正情報処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, modules, module, dataWidth, totalWidth
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set modules = service.ExecQuery("SELECT DeviceLocator, DataWidth, TotalWidth FROM Win32_PhysicalMemory")
If modules.Count = 0 Then
    WScript.Echo "No physical memory modules reported."
    WScript.Quit 0
End If
For Each module In modules
    dataWidth = module.DataWidth
    totalWidth = module.TotalWidth
    If IsNull(dataWidth) Or IsNull(totalWidth) Then
        WScript.Echo module.DeviceLocator & " DataWidth/TotalWidth unavailable."
        WScript.Echo "ECC indication unavailable."
    Else
        WScript.Echo module.DeviceLocator & " DataWidth=" & dataWidth & " TotalWidth=" & totalWidth
        If CLng(totalWidth) > CLng(dataWidth) Then
            WScript.Echo "Additional check-bit width reported; ECC is likely."
        Else
            WScript.Echo "No additional ECC width reported."
        End If
    End If
Next
