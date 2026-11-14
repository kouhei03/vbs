' ファイル概要: Windows Updateポリシーの値名と型を列挙する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectWindowsUpdatePolicy.vbs
' 実行結果: Windows Updateポリシーの値名と型を列挙する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, names, types, i, value
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.EnumValues(&H80000002, "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", names, types)
If result <> 0 Then
    WScript.Echo "Windows Update policy key unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Windows Update policy values (names and types only):"
If IsArray(names) Then
    For i = 0 To UBound(names)
        WScript.Echo names(i) & " (type " & types(i) & ")"
    Next
End If
