' ファイル概要: コンソール進捗表示デモ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 160_ConsoleProgressDemo.vbs
' 実行結果: コンソール進捗表示デモ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim percent
For percent=0 To 100 Step 10
    WScript.StdOut.Write "Progress: " & percent & "%" & vbCr
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    WScript.Sleep 100
Next
WScript.StdOut.WriteLine "Progress: 100%"
