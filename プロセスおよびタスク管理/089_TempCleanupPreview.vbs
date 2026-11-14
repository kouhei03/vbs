' ファイル概要: TEMP古いファイルの削除プレビューのみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 089_TempCleanupPreview.vbs [age-days]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: TEMP古いファイルの削除プレビューのみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, shell, tempPath, folder, file, ageDays, cutoff, found
If WScript.Arguments.Count > 1 Then WScript.Echo "Usage: cscript //nologo 089_TempCleanupPreview.vbs [age-days]": WScript.Quit 2
ageDays = 30
If WScript.Arguments.Count = 1 Then
    If Not IsNumeric(WScript.Arguments(0)) Then WScript.Echo "Age must be a number.": WScript.Quit 2
    ageDays = CLng(WScript.Arguments(0))
End If
If ageDays < 1 Or ageDays > 3650 Then WScript.Echo "Age must be 1..3650 days.": WScript.Quit 2
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
tempPath = shell.ExpandEnvironmentStrings("%TEMP%")
Set folder = fso.GetFolder(tempPath)
' [処理解説] 指定日数より古いTEMP直下のファイルを候補として表示し、削除はしません。
cutoff = DateAdd("d", -ageDays, Now)
found = False
For Each file In folder.Files
    If file.DateLastModified < cutoff Then
        WScript.Echo file.Path & " modified=" & file.DateLastModified & " bytes=" & file.Size
        found = True
    End If
Next
If Not found Then WScript.Echo "No old files found in TEMP root. No files were deleted."
