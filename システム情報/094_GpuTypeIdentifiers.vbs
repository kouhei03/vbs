' 解説: GPU名やメーカー識別子から統合/外部GPUを推定します。分類は名称による参考情報で確定判定ではありません。
Option Explicit
' 解説: WMI接続、ビデオコントローラー、および推定表示を保持します。
Dim service, adapters, adapter, identity, classification
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set adapters = service.ExecQuery("SELECT Name, AdapterCompatibility, VideoProcessor, PNPDeviceID FROM Win32_VideoController")
' 解説: コレクション内の各項目を順に処理します: For Each adapter In adapters
For Each adapter In adapters
    ' 解説: 取得値または計算結果を変数へ格納します。
    identity = LCase(SafeText(adapter.Name) & " " & SafeText(adapter.AdapterCompatibility) & " " & SafeText(adapter.VideoProcessor))
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: InStr(identity, "intel") > 0 Or InStr(identity, "uhd graphics") > 0 Or InStr(identity, "iris") > 0
    If InStr(identity, "intel") > 0 Or InStr(identity, "uhd graphics") > 0 Or InStr(identity, "iris") > 0 Then
        ' 解説: 取得値または計算結果を変数へ格納します。
        classification = "Likely integrated (heuristic)"
    ' 解説: 先行条件が不成立の場合に追加条件を評価します: InStr(identity, "nvidia") > 0 Or InStr(identity, "geforce") > 0 Or InStr(identity, "radeon") > 0
    ElseIf InStr(identity, "nvidia") > 0 Or InStr(identity, "geforce") > 0 Or InStr(identity, "radeon") > 0 Then
        ' 解説: 取得値または計算結果を変数へ格納します。
        classification = "Likely discrete (heuristic; some Radeon GPUs are integrated)"
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 取得値または計算結果を変数へ格納します。
        classification = "Unknown; inspect device documentation"
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo SafeText(adapter.Name) & " / " & SafeText(adapter.AdapterCompatibility)
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Classification: " & classification
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "PNPDeviceID=" & SafeText(adapter.PNPDeviceID)
' 解説: 現在の反復を終え、次の要素へ進みます。
Next

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
