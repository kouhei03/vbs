' 解説: コンソールに進捗率を表示するだけのローカルデモです。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: percent
Dim percent
' 解説: 指定範囲の要素を順番に処理します: For percent=0 To 100 Step 10
For percent=0 To 100 Step 10
    ' 解説: この文でローカル診断処理を進めます: WScript.StdOut.Write "Progress: " & percent & "%" & vbCr
    WScript.StdOut.Write "Progress: " & percent & "%" & vbCr
    ' 解説: この文でローカル診断処理を進めます: WScript.Sleep 100
    WScript.Sleep 100
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: この文でローカル診断処理を進めます: WScript.StdOut.WriteLine "Progress: 100%"
WScript.StdOut.WriteLine "Progress: 100%"
