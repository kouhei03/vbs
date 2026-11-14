Option Explicit
Dim service, processors, processor, architectureName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processors = service.ExecQuery("SELECT Architecture, AddressWidth, DataWidth FROM Win32_Processor")
For Each processor In processors
    Select Case processor.Architecture
        Case 0: architectureName = "x86"
        Case 5: architectureName = "ARM"
        Case 9: architectureName = "x64"
        Case 12: architectureName = "ARM64"
        Case 6: architectureName = "Itanium"
        Case Else: architectureName = "Unknown (" & processor.Architecture & ")"
    End Select
    WScript.Echo "Architecture=" & architectureName
    WScript.Echo "AddressWidth=" & processor.AddressWidth & " DataWidth=" & processor.DataWidth
Next
