' 概要: 指定タイトルの既存ウィンドウがあるか確認します。キー送信は行いません。
' 実行方法: Usage: cscript //nologo 091_WindowTitleAudit.vbs <window-title>
' 処理結果: ウィンドウタイトルの確認（SendKeysなし）。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellを使ってWindowsの設定や状態を照会します。
Option Explicit
Dim title, shell
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 091_WindowTitleAudit.vbs <window-title>"
    WScript.Quit 2
End If
' 主処理: ウィンドウタイトルの確認（SendKeysなし）。
title = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
If shell.AppActivate(title) Then
    WScript.Echo "Matching window exists (it may have been activated)."
Else
    WScript.Echo "No matching window found."
    WScript.Quit 1
End If
