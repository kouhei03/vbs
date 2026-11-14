' 概要: 公開IP取得など、外部サービスへの追跡リクエストは行わない安全な案内例です。
' 実行方法: cscript //nologo 013_PublicIpLookupExcluded.vbs
' 処理結果: 外部IPチェッカーへの接続を行わない案内。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsまたはネットワークの状態を確認し、結果を表示します。
Option Explicit
' 主処理: 外部IPチェッカーへの接続を行わない案内。
WScript.Echo "No external IP-checker request was made."
WScript.Echo "For privacy, inspect the network gateway or an organization-approved service manually."
