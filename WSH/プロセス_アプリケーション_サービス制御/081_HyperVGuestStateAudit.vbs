' 概要: Hyper-V仮想マシン名と電源状態を読み取ります。状態変更は行いません。
' 実行方法: cscript //nologo 081_HyperVGuestStateAudit.vbs
' 処理結果: Hyper-V仮想マシン状態照会（制御なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim virtualization, machines, machine
Set virtualization = GetObject("winmgmts:\\.\root\virtualization\v2")
' 主処理: WMIの照会結果を処理します。
Set machines = virtualization.ExecQuery("SELECT ElementName, EnabledState, HealthState FROM Msvm_ComputerSystem WHERE Caption='Virtual Machine'")
For Each machine In machines
    WScript.Echo machine.ElementName & vbTab & "EnabledState=" & machine.EnabledState & vbTab & "HealthState=" & machine.HealthState
Next
