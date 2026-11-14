' 概要: 影響範囲が限定されたHKCUデモ用.regテンプレートを新規ファイルに出力します。インポートはしません。
' 実行方法: Usage: cscript //nologo 184_RegistryTemplatePreview.vbs <new-file.reg>
' 処理結果: HKCUデモ用.regテンプレートを新規作成（インポートなし）。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, output
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
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
' 主処理: 結果ファイルを作成します。
Set output = fso.CreateTextFile(path, False, True)
output.WriteLine "Windows Registry Editor Version 5.00"
output.WriteLine
output.WriteLine "[HKEY_CURRENT_USER\Software\WSH\SafeExamples\Demo]"
output.WriteLine """ExampleValue""=""Preview only"""
output.Close
WScript.Echo "Created template only; it was not imported: " & path
