' 概要: 資格情報のキャッシュ削除は行いません。Windows標準の資格情報管理機能を案内します。
' 実行方法: cscript //nologo 032_ProxyCredentialCacheNotice.vbs
' 処理結果: 資格情報を扱わない安全な案内。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsまたはネットワークの状態を確認し、結果を表示します。
Option Explicit
' 主処理: 資格情報を扱わない安全な案内。
WScript.Echo "No proxy credentials or cached credentials were read or deleted."
WScript.Echo "Review saved credentials manually in Windows Credential Manager."
