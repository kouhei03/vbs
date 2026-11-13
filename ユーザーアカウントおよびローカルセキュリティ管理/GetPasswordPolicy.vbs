' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（shell、command、standardOutput、standardError）。主な用途: Windows Script Hostシェル、実行した固定コマンド、標準出力内容、標準エラー出力内容。
Dim shell, command, standardOutput, standardError
' 解説: WScript.Shell（Windows環境変数や固定のローカルコマンドを扱うオブジェクト）を生成し、変数に格納します。
Set shell = CreateObject("WScript.Shell")

' 解説: 既存のコメント行を補足として残し、直後の処理の意図を説明しています。
' This fixed command displays local account policy and does not change it.
' 解説: 固定されたローカルコマンドを実行し、標準出力/エラーや終了コードを取得します: Set command = shell.Exec("net.exe accounts")
Set command = shell.Exec("net.exe accounts")
' 解説: 標準出力または標準エラーの残りをすべて読み取り、デッドロックを避けて実行結果を回収します。
standardOutput = command.StdOut.ReadAll
' 解説: 標準出力または標準エラーの残りをすべて読み取り、デッドロックを避けて実行結果を回収します。
standardError = command.StdErr.ReadAll

' 解説: 条件「Len(standardOutput) > 0 Then WScript.Echo standardOutput」を評価し、成立した場合に続く分岐処理を実行します。
If Len(standardOutput) > 0 Then WScript.Echo standardOutput
' 解説: 条件「command.ExitCode <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If command.ExitCode <> 0 Then
    ' 解説: 条件「Len(standardError) > 0 Then WScript.Echo standardError」を評価し、成立した場合に続く分岐処理を実行します。
    If Len(standardError) > 0 Then WScript.Echo standardError
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "net accounts failed with exit code " & command.ExitCode
    WScript.Echo "net accounts failed with exit code " & command.ExitCode
    ' 解説: スクリプトを終了し、呼び出し元に終了コード command.ExitCode（終了コードを返して終了）を返します。
    WScript.Quit command.ExitCode
' 解説: 条件分岐ブロックを終了します。
End If
