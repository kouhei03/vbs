' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process、dateTime、startTime）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報、WMI日時変換オブジェクト、監視開始時刻。
Dim service, processes, process, dateTime, startTime
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、Name, ProcessId, CreationDateを取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, CreationDate FROM Win32_Process")
' 解説: WbemScripting.SWbemDateTime（WMI日時の変換/比較に使うオブジェクト）を生成し、変数に格納します。
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name" & vbTab & "PID" & vbTab & "Start time"
WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Start time"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 条件「Not IsNull(process.CreationDate) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If Not IsNull(process.CreationDate) Then
        ' 解説: WMIから得たCIM日時文字列を日時変換オブジェクトに設定し、VBScript日付型へ変換できるようにします。
        dateTime.Value = process.CreationDate
        ' 解説: CIM日時値をVBScriptの日付型へ変換し、表示または計算に使います。
        startTime = dateTime.GetVarDate(True)
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & vbTab & process.ProcessId & vbTab & CStr(startTime)
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            CStr(startTime)
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
