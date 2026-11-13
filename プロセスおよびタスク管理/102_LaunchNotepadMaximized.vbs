' 解説: 固定された標準アプリを最大化表示で起動するローカル例です。任意コマンドは受け取りません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, pid
Dim shell, pid
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 取得値または計算結果を変数に格納します。
pid = shell.Run("notepad.exe", 3, False)
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Notepad launch requested in maximized mode; return value=" & pid
