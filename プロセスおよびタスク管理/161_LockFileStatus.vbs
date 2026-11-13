' 解説: #154のロックファイル状態を確認するだけで、作成/削除は行いません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: fso, shell, lockPath
Dim fso, shell, lockPath
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set fso=CreateObject("Scripting.FileSystemObject")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
lockPath=shell.ExpandEnvironmentStrings("%TEMP%") & "\vbs-process-example.lock"
' 解説: 条件を判定し、必要な分岐を選択します: If fso.FileExists(lockPath)
If fso.FileExists(lockPath) Then
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "Lock file exists: " & lockPath
' 解説: 先行条件が成立しない場合の代替処理へ進みます。
Else
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo "No lock file exists."
' 解説: 条件分岐ブロックを終了します。
End If
