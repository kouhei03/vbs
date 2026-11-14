' ファイル概要: UEFI/レガシーブート方式。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 041_FirmwareBootType.vbs
' 実行結果: UEFI/レガシーブート方式処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, systems, system, bootType
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT FirmwareType FROM Win32_ComputerSystem")
If systems.Count = 0 Then
    WScript.Echo "FirmwareType is unavailable."
    WScript.Quit 1
End If
For Each system In systems
    If IsNull(system.FirmwareType) Then
        bootType = "Unknown; firmware type was not reported"
    Else
        Select Case system.FirmwareType
        Case 2: bootType = "Legacy BIOS"
        Case 3: bootType = "UEFI"
        Case Else: bootType = "Unknown or unsupported: " & system.FirmwareType
        End Select
    End If
    WScript.Echo "Firmware boot type: " & bootType
Next
