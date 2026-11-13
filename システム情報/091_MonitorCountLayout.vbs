' 解説: WMIが認識するディスプレイ数と識別情報を表示します。物理配置座標はこのWMIクラスでは取得できません。
Option Explicit
' 解説: WMI接続、モニター一覧、個々のモニター、および件数を保持します。
Dim service, monitors, monitor, monitorCount
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set monitors = service.ExecQuery("SELECT Name, PNPDeviceID, ScreenWidth, ScreenHeight, Status FROM Win32_DesktopMonitor")
' 解説: 取得値または計算結果を変数へ格納します。
monitorCount = monitors.Count
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "WMI-reported monitor count: " & monitorCount
' 解説: コレクション内の各項目を順に処理します: For Each monitor In monitors
For Each monitor In monitors
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo SafeText(monitor.Name) & " " & SafeText(monitor.ScreenWidth) & "x" & SafeText(monitor.ScreenHeight)
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "PNPDeviceID=" & SafeText(monitor.PNPDeviceID) & " Status=" & SafeText(monitor.Status)
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Monitor physical arrangement coordinates are not exposed by this query."

' 解説: 未報告のWMIプロパティを安全に空文字列へ変換します。
Function SafeText(value)
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsNull(value)
    If IsNull(value) Then
        ' 解説: 取得値または計算結果を変数へ格納します。
        SafeText = ""
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 取得値または計算結果を変数へ格納します。
        SafeText = CStr(value)
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 補助関数の定義を終了します。
End Function
