' 解説: TEMP直下で指定日数より古いファイルを列挙するだけの後片付けプレビューです。削除しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, shell, tempPath, folder, file, ageDays, cutoff, found
Dim fso, shell, tempPath, folder, file, ageDays, cutoff, found
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count > 1
If WScript.Arguments.Count > 1 Then WScript.Echo "Usage: cscript //nologo 089_TempCleanupPreview.vbs [age-days]": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
ageDays = 30
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count = 1
If WScript.Arguments.Count = 1 Then
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not IsNumeric(WScript.Arguments(0))
    If Not IsNumeric(WScript.Arguments(0)) Then WScript.Echo "Age must be a number.": WScript.Quit 2
    ' 解説: 取得値または計算結果を変数に格納します。
    ageDays = CLng(WScript.Arguments(0))
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 条件を判定し、必要な分岐を選択します: If ageDays < 1 Or ageDays > 3650
If ageDays < 1 Or ageDays > 3650 Then WScript.Echo "Age must be 1..3650 days.": WScript.Quit 2
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso = CreateObject("Scripting.FileSystemObject")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
tempPath = shell.ExpandEnvironmentStrings("%TEMP%")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set folder = fso.GetFolder(tempPath)
' 解説: 取得値または計算結果を変数に格納します。
cutoff = DateAdd("d", -ageDays, Now)
' 解説: 取得値または計算結果を変数に格納します。
found = False
' 解説: コレクション内の各項目を順番に処理します: For Each file In folder.Files
For Each file In folder.Files
    ' 解説: 条件を判定し、必要な分岐を選択します: If file.DateLastModified < cutoff
    If file.DateLastModified < cutoff Then
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo file.Path & " modified=" & file.DateLastModified & " bytes=" & file.Size
        ' 解説: 取得値または計算結果を変数に格納します。
        found = True
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を判定し、必要な分岐を選択します: If Not found
If Not found Then WScript.Echo "No old files found in TEMP root. No files were deleted."
