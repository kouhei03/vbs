' ファイル概要: 筐体タイプ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 042_ChassisType.vbs
' 実行結果: 筐体タイプ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, enclosures, enclosure, code, label
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set enclosures = service.ExecQuery("SELECT ChassisTypes FROM Win32_SystemEnclosure")
If enclosures.Count = 0 Then
    WScript.Echo "No chassis information returned."
    WScript.Quit 1
End If
For Each enclosure In enclosures
    If IsArray(enclosure.ChassisTypes) Then
        For Each code In enclosure.ChassisTypes
            Select Case CLng(code)
                Case 3: label = "Desktop"
                Case 8, 9, 10, 14: label = "Portable/Laptop/Notebook"
                Case 11: label = "Handheld"
                Case 13: label = "All-in-One"
                Case 30: label = "Tablet"
                Case 31: label = "Convertible"
                Case 32: label = "Detachable"
                Case 7: label = "Tower"
                Case 23: label = "Rack Mount"
                Case Else: label = "Other SMBIOS chassis type"
            End Select
            WScript.Echo "Chassis type code " & code & ": " & label
        Next
    Else
        WScript.Echo "Chassis type data is unavailable."
    End If
Next
