' 概要: 固定回数のコンソール進捗表示をデモします。
' 実行方法: cscript //nologo 194_ConsoleProgressDemo.vbs
' 処理結果: コンソール進捗表示デモ。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsの状態または固定データを確認し、結果を表示します。
Option Explicit
Dim i, count, bar
' 主処理: コンソール進捗表示デモ。
count = 20
For i = 1 To count
    bar = String(i, "#") & String(count - i, "-")
    WScript.StdOut.Write vbCr & "[" & bar & "] " & CInt(i * 100 / count) & "%"
    WScript.Sleep 100
Next
' 取得した値または処理状態をコンソールへ表示します。
WScript.Echo
