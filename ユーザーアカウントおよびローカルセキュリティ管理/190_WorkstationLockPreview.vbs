' ファイル概要: 端末ロックのプレビューを表示する（ロックなし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 190_WorkstationLockPreview.vbs
' 実行結果: 端末ロックのプレビューを表示する（ロックなし）処理の結果やエラーをコンソールに表示します。

Option Explicit

WScript.Echo "Workstation lock preview."
WScript.Echo "No lock command was issued; review local policy and user impact before locking."
