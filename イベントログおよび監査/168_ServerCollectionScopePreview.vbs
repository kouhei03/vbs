' ファイル概要: サーバー一覧ファイルの名前を検証する（ネットワーク接続なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 168_ServerCollectionScopePreview.vbs <server-list-file>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サーバー一覧ファイルの名前を検証する（ネットワーク接続なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim inputPath, fso, inputFile, expression, line, count, uniqueNames
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 168_ServerCollectionScopePreview.vbs <server-list-file>"
    WScript.Quit 2
End If
inputPath = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(inputPath) Then
    WScript.Echo "Input file does not exist."
    WScript.Quit 1
End If
If fso.GetFile(inputPath).Size > 1048576 Then
    WScript.Echo "Input file exceeds the 1 MiB limit."
    WScript.Quit 2
End If
Set expression = New RegExp
expression.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
expression.IgnoreCase = True
Set uniqueNames = CreateObject("Scripting.Dictionary")
' [処理解説] 辞書のキーは大文字・小文字を区別せず照合します。
uniqueNames.CompareMode = vbTextCompare
Set inputFile = fso.OpenTextFile(inputPath, 1, False)
Do Until inputFile.AtEndOfStream
    line = Trim(inputFile.ReadLine)
    If Len(line) > 0 And Left(line, 1) <> "#" Then
        If expression.Test(line) Then uniqueNames(line) = True
    End If
Loop
inputFile.Close
count = uniqueNames.Count
WScript.Echo "Validated unique server names: " & count
WScript.Echo "No network access or remote log collection was performed."
