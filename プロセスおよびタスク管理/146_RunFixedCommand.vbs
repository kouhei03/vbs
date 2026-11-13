' 解説: 固定された安全なWindowsバージョン表示コマンドを同期/非同期で呼び出す例です。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, waitForExit, result, mode
Dim shell, waitForExit, result, mode
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count<>1
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 146_RunFixedCommand.vbs <sync|async>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
mode=LCase(WScript.Arguments(0))
' 解説: 条件を判定し、必要な分岐を選択します: If mode="sync"
If mode="sync" Then
    ' 解説: 取得値または計算結果を変数に格納します。
    waitForExit=True
' 解説: 先行条件が不成立のとき追加条件を判定します: ElseIf mode="async"
ElseIf mode="async" Then
    ' 解説: 取得値または計算結果を変数に格納します。
    waitForExit=False
' 解説: 先行条件が成立しない場合の代替処理へ進みます。
Else
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Mode must be sync or async."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
result=shell.Run("cmd.exe /c ver",1,waitForExit)
' 解説: 条件を判定し、必要な分岐を選択します: If waitForExit
If waitForExit Then WScript.Echo "Exit code=" & result Else WScript.Echo "Command launched asynchronously."
