' 解説: PROCESSスコープの環境変数を一時設定する例です。ユーザー/システム環境変数は変更しません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: shell, environment
Dim shell, environment
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell=CreateObject("WScript.Shell")
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set environment=shell.Environment("PROCESS")
' 解説: 取得値または計算結果を変数に格納します。
environment("VBS_PROCESS_DEMO")="temporary-value"
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Process-scope value=" & environment("VBS_PROCESS_DEMO")
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "This value is limited to this script process and its children."
