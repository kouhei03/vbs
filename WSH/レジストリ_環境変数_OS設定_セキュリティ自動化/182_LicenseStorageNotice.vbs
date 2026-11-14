' 概要: ライセンス情報をレジストリへ保存・暗号化する処理はしません。
' 実行方法: cscript //nologo 182_LicenseStorageNotice.vbs
' 処理結果: ライセンス情報を保存しない安全案内。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsの状態または固定データを確認し、結果を表示します。
Option Explicit
' 取得した値または処理状態をコンソールへ表示します。
' 主処理: ライセンス情報を保存しない安全案内。
WScript.Echo "No license value was requested, stored, or transmitted."
WScript.Echo "Use an approved secret manager or the application's documented DPAPI-backed storage."
