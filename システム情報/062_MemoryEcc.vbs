' 解説: DIMMごとのDataWidthとTotalWidthを比較し、ECC用の追加幅が報告されるか確認します。
Option Explicit
' 解説: WMI接続、メモリモジュール、および幅の比較結果を保持します。
Dim service, modules, module, dataWidth, totalWidth
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set modules = service.ExecQuery("SELECT DeviceLocator, DataWidth, TotalWidth FROM Win32_PhysicalMemory")
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: modules.Count = 0
If modules.Count = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "No physical memory modules reported."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 0
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: コレクション内の各項目を順に処理します: For Each module In modules
For Each module In modules
    ' 解説: 取得値または計算結果を変数へ格納します。
    dataWidth = module.DataWidth
    ' 解説: 取得値または計算結果を変数へ格納します。
    totalWidth = module.TotalWidth
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsNull(dataWidth) Or IsNull(totalWidth)
    If IsNull(dataWidth) Or IsNull(totalWidth) Then
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo module.DeviceLocator & " DataWidth/TotalWidth unavailable."
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "ECC indication unavailable."
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo module.DeviceLocator & " DataWidth=" & dataWidth & " TotalWidth=" & totalWidth
        ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: CLng(totalWidth) > CLng(dataWidth)
        If CLng(totalWidth) > CLng(dataWidth) Then
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo "Additional check-bit width reported; ECC is likely."
        ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
        Else
            ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
            WScript.Echo "No additional ECC width reported."
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
