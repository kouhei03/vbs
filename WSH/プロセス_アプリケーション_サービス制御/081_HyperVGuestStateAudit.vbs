' Hyper-V仮想マシン名と電源状態を読み取ります。状態変更は行いません。
Option Explicit
Dim virtualization, machines, machine
Set virtualization = GetObject("winmgmts:\\.\root\virtualization\v2")
Set machines = virtualization.ExecQuery("SELECT ElementName, EnabledState, HealthState FROM Msvm_ComputerSystem WHERE Caption='Virtual Machine'")
For Each machine In machines
    WScript.Echo machine.ElementName & vbTab & "EnabledState=" & machine.EnabledState & vbTab & "HealthState=" & machine.HealthState
Next
