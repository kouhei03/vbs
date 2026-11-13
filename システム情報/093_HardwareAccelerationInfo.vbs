' 解説: GPUドライバーと状態を照会します。WMIだけではアクセラレーションの全機能状態を断定できません。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: service, adapters, adapter
Dim service, adapters, adapter
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set adapters = service.ExecQuery("SELECT Name, Status, Availability, VideoModeDescription FROM Win32_VideoController")
' 解説: コレクション内の各項目を順に処理します: For Each adapter In adapters
For Each adapter In adapters
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo adapter.Name & " Status=" & adapter.Status & " Availability=" & adapter.Availability
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "VideoModeDescription=" & adapter.VideoModeDescription
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "WMI does not expose a definitive global hardware-acceleration on/off switch."
