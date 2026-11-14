' 概要: USBSTOR配下のデバイスキー件数だけを数えます。デバイス識別子は表示しません。
' 実行方法: cscript //nologo 191_UsbStorageHistoryCount.vbs
' 処理結果: USBSTOR履歴キー件数（識別子非表示）。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsの状態または固定データを確認し、結果を表示します。
Option Explicit
Const HKEY_LOCAL_MACHINE = &H80000002
Dim registry, result, inputParams, outputParams, names, count
' 主処理: USBSTOR履歴キー件数（識別子非表示）。
Set registry = GetObject("winmgmts:\\.\root\default:StdRegProv")
Set inputParams = registry.Methods_("EnumKey").InParameters.SpawnInstance_()
inputParams.hDefKey = HKEY_LOCAL_MACHINE
inputParams.sSubKeyName = "SYSTEM\CurrentControlSet\Enum\USBSTOR"
Set outputParams = registry.ExecMethod_("EnumKey", inputParams)
If outputParams.ReturnValue <> 0 Then
    ' 取得した値または処理状態をコンソールへ表示します。
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
