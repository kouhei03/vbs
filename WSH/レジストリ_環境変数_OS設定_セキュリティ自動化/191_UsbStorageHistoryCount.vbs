' USBSTOR配下のデバイスキー件数だけを数えます。デバイス識別子は表示しません。
Option Explicit
Const HKEY_LOCAL_MACHINE = &H80000002
Dim registry, result, inputParams, outputParams, names, count
Set registry = GetObject("winmgmts:\\.\root\default:StdRegProv")
Set inputParams = registry.Methods_("EnumKey").InParameters.SpawnInstance_()
inputParams.hDefKey = HKEY_LOCAL_MACHINE
inputParams.sSubKeyName = "SYSTEM\CurrentControlSet\Enum\USBSTOR"
Set outputParams = registry.ExecMethod_("EnumKey", inputParams)
If outputParams.ReturnValue <> 0 Then
    WScript.Echo "USBSTOR key is absent or inaccessible."
    WScript.Quit 1
End If
count = 0
names = outputParams.sNames
If Not IsNull(names) Then
    On Error Resume Next
    count = UBound(names) - LBound(names) + 1
    If Err.Number <> 0 Then count = 0: Err.Clear
    On Error GoTo 0
End If
WScript.Echo "USBSTOR device-class key count=" & count & " (identifiers omitted)."
