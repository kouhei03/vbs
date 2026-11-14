' 影響範囲が限定されたHKCUデモ用.regテンプレートを新規ファイルに出力します。インポートはしません。
Option Explicit
Dim fso, path, output
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 184_RegistryTemplatePreview.vbs <new-file.reg>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If LCase(fso.GetExtensionName(path)) <> "reg" Or fso.FileExists(path) Then
    WScript.Echo "Provide a new .reg file path."
    WScript.Quit 2
End If
Set output = fso.CreateTextFile(path, False, True)
output.WriteLine "Windows Registry Editor Version 5.00"
output.WriteLine
output.WriteLine "[HKEY_CURRENT_USER\Software\WSH\SafeExamples\Demo]"
output.WriteLine """ExampleValue""=""Preview only"""
output.Close
WScript.Echo "Created template only; it was not imported: " & path
