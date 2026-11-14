Option Explicit
Dim service, modules, module, formName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set modules = service.ExecQuery("SELECT DeviceLocator, FormFactor FROM Win32_PhysicalMemory")
For Each module In modules
    Select Case module.FormFactor
        Case 8: formName = "DIMM"
        Case 12: formName = "SODIMM"
        Case 9: formName = "RIMM"
        Case 13: formName = "SRIMM"
        Case Else: formName = "Other/unknown code " & module.FormFactor
    End Select
    WScript.Echo module.DeviceLocator & ": " & formName
Next
