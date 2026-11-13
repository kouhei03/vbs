' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（timeFilter、service、events、eventItem、logonCount、logoffCount）。主な用途: WMI日時フィルター、WMIサービス接続、イベント一覧、イベントログ項目。
Dim timeFilter, service, events, eventItem, logonCount, logoffCount
' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' 解説: ローカル日時をWMIが扱うCIM日時形式に変換して日時オブジェクトへ設定します。
timeFilter.SetVarDate DateAdd("h", -24, Now), True
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsイベントログ（Win32_NTLogEvent）を読み取り専用で照会し、EventCodeを取得します。 絞り込み条件には、検証済みの実行時引数や計算した日時が挿入されます（文字列連結位置を含むため、ここでは空値として表示しません）。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='Security' " & _
    "AND (EventCode=4624 OR EventCode=4634) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: logonCount = 0
logonCount = 0
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: logoffCount = 0
logoffCount = 0
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each eventItem In events
For Each eventItem In events
    ' 解説: 条件「eventItem.EventCode = 4624 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If eventItem.EventCode = 4624 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: logonCount = logonCount + 1
        logonCount = logonCount + 1
    ' 解説: 先行する条件が不成立の場合に評価する追加条件です: eventItem.EventCode = 4634 Then
    ElseIf eventItem.EventCode = 4634 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: logoffCount = logoffCount + 1
        logoffCount = logoffCount + 1
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Successful logon events (4624), last 24 hours: " & logonCount
WScript.Echo "Successful logon events (4624), last 24 hours: " & logonCount
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Logoff events (4634), last 24 hours: " & logoffCount
WScript.Echo "Logoff events (4634), last 24 hours: " & logoffCount
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Counts only; account, session, and network details are omitted."
WScript.Echo "Counts only; account, session, and network details are omitted."
