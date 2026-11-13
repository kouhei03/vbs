' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、disks、disk）。主な用途: WMIサービス接続、物理ディスク一覧、物理ディスク情報。
Dim service, disks, disk
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの物理ディスク（Win32_DiskDrive）を読み取り専用で照会し、Model, Size, MediaTypeを取得します。
Set disks = service.ExecQuery( _
    "SELECT Model, Size, MediaType FROM Win32_DiskDrive")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each disk In disks
For Each disk In disks
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Model: " & disk.Model
    WScript.Echo "Model: " & disk.Model
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Media type: " & disk.MediaType
    WScript.Echo "Media type: " & disk.MediaType
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Capacity (bytes): " & disk.Size
    WScript.Echo "Capacity (bytes): " & disk.Size
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: ""
    WScript.Echo ""
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
