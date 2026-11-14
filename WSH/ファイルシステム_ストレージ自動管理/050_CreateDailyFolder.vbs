' 概要: 指定親フォルダーの下に今日の日付(yyyy-MM-dd)のフォルダーを作成します。
' 実行方法: Usage: cscript //nologo 050_CreateDailyFolder.vbs <parent-folder>
' 処理結果: 日付名フォルダー作成。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, parent, folderPath, dateName
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 050_CreateDailyFolder.vbs <parent-folder>"
    WScript.Quit 2
End If
' 主処理: 日付名フォルダー作成。
Set fso = CreateObject("Scripting.FileSystemObject")
parent = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FolderExists(parent) Then
    WScript.Echo "Parent folder not found: " & parent
    WScript.Quit 1
End If
dateName = Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2)
folderPath = fso.BuildPath(parent, dateName)
If fso.FolderExists(folderPath) Then
    WScript.Echo "Already exists: " & folderPath
Else
    fso.CreateFolder folderPath
    WScript.Echo "Created: " & folderPath
End If
