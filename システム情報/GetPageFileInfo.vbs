' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、settings、setting、usages、usage、found）。主な用途: WMIサービス接続、設定一覧、設定情報、対象が見つかったかを示すフラグ。
Dim service, settings, setting, usages, usage, found
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")

' 解説: WMIのページファイル設定（Win32_PageFileSetting）を読み取り専用で照会し、Name, InitialSize, MaximumSizeを取得します。
Set settings = service.ExecQuery( _
    "SELECT Name, InitialSize, MaximumSize FROM Win32_PageFileSetting")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each setting In settings
For Each setting In settings
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
    found = True
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Configured page file: " & setting.Name
    WScript.Echo "Configured page file: " & setting.Name
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Initial size (MB): " & setting.InitialSize
    WScript.Echo "Initial size (MB): " & setting.InitialSize
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Maximum size (MB): " & setting.MaximumSize
    WScript.Echo "Maximum size (MB): " & setting.MaximumSize
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件「Not found Then WScript.Echo "No explicit page file settings returned (possibly system-managed)."」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then WScript.Echo "No explicit page file settings returned (possibly system-managed)."

' 解説: WMIのページファイル使用状況（Win32_PageFileUsage）を読み取り専用で照会し、Name, AllocatedBaseSize, CurrentUsage, PeakUsageを取得します。
Set usages = service.ExecQuery( _
    "SELECT Name, AllocatedBaseSize, CurrentUsage, PeakUsage FROM Win32_PageFileUsage")
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each usage In usages
For Each usage In usages
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Active page file: " & usage.Name
    WScript.Echo "Active page file: " & usage.Name
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Allocated size (MB): " & usage.AllocatedBaseSize
    WScript.Echo "Allocated size (MB): " & usage.AllocatedBaseSize
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Current usage (MB): " & usage.CurrentUsage
    WScript.Echo "Current usage (MB): " & usage.CurrentUsage
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Peak usage (MB): " & usage.PeakUsage
    WScript.Echo "Peak usage (MB): " & usage.PeakUsage
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
