' 概要: 指定タイトルの既存ウィンドウをアクティブにします。SendKeysや最小化は行いません。
' 実行方法: Usage: cscript //nologo 065_ActivateWindowTitle.vbs <window-title>
' 処理結果: 既存ウィンドウのアクティブ化。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellを使ってWindowsの設定や状態を照会します。
Option Explicit
Dim title, shell
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 065_ActivateWindowTitle.vbs <window-title>"
    WScript.Quit 2
End If
' 主処理: 既存ウィンドウのアクティブ化。
title = WScript.Arguments(0)
If title = "" Then WScript.Quit 2
Set shell = CreateObject("WScript.Shell")
If shell.AppActivate(title) Then
    WScript.Echo "Activated matching window."
Else
    WScript.Echo "No matching window found."
    WScript.Quit 1
End If
