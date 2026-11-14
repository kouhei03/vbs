' /verbose引数の解釈例です。サービス構成や状態は変更しません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: debugMode, argument
Dim debugMode, argument
' 解説: 入力値、取得値、または計算結果を変数に格納します。
 debugMode = False
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each argument In WScript.Arguments
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If LCase(argument) = "/verbose" Then debugMode = True
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If debugMode Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Verbose mode enabled for this local read-only demo."
' 解説: 先行条件に該当しない場合の処理です。
Else
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Normal mode. Pass /verbose for an extra diagnostic line."
' 解説: 条件分岐を終了します。
End If
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "No service operation was performed."
