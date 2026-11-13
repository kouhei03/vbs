' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（shell、environment、variableName）。主な用途: Windows Script Hostシェル。
Dim shell, environment, variableName
' 解説: WScript.Shell（Windows環境変数や固定のローカルコマンドを扱うオブジェクト）を生成し、変数に格納します。
Set shell = CreateObject("WScript.Shell")
' 解説: 現在のプロセス環境変数を読み取るためのEnvironmentコレクションを取得します。
Set environment = shell.Environment("PROCESS")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each variableName In Array("SystemRoot", "TEMP", "TMP")
For Each variableName In Array("SystemRoot", "TEMP", "TMP")
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: variableName & "=" & environment(variableName)
    WScript.Echo variableName & "=" & environment(variableName)
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
