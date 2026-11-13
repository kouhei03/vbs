' 解説: ドラッグ&ドロップ等で渡されたファイルのメタデータだけを表示します。内容は開きません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, i, path, file
Dim fso, i, path, file
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso=CreateObject("Scripting.FileSystemObject")
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count=0
If WScript.Arguments.Count=0 Then WScript.Echo "Drop one or more files onto the script or pass file paths.": WScript.Quit 2
' 解説: 指定範囲の要素を順番に処理します: For i=0 To WScript.Arguments.Count-1
For i=0 To WScript.Arguments.Count-1
    ' 解説: 取得値または計算結果を変数に格納します。
    path=WScript.Arguments(i)
    ' 解説: 条件を判定し、必要な分岐を選択します: If fso.FileExists(path)
    If fso.FileExists(path) Then
        ' 解説: 後続処理で利用するオブジェクト参照を設定します。
        Set file=fso.GetFile(path)
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo file.Path & " bytes=" & file.Size & " modified=" & file.DateLastModified
    ' 解説: 先行条件が成立しない場合の代替処理へ進みます。
    Else
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo "Not a file: " & path
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
