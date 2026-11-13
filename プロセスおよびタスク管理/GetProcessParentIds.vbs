' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報。
Dim service, processes, process
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、Name, ProcessId, ParentProcessIdを取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, ParentProcessId FROM Win32_Process")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name" & vbTab & "PID" & vbTab & "Parent PID"
WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Parent PID"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & vbTab & process.ProcessId & vbTab & process.ParentProcessId
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.ParentProcessId
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
