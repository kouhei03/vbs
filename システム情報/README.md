# vbs
<!-- 解説: この見出しは「vbs」を説明する節のタイトルです。 -->

## System information scripts
<!-- 解説: この見出しは「System information scripts」を説明する節のタイトルです。 -->

The following scripts query local OS, hardware, storage, and peripheral
information through WMI or built-in Windows Script Host APIs. Run an
individual script from Command Prompt:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: The following scripts query local OS, hardware, storage, and peripheral -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: information through WMI or built-in Windows Script Host APIs. Run an -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: individual script from Command Prompt: -->

```text
cscript //nologo GetOSInfo.vbs
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はGetOSInfo.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

Other scripts in this directory cover architecture, computer name, join
status, service pack, uptime, CPU and memory, page file, motherboard, BIOS,
disks, volumes, display, audio, USB devices, printers, battery, input devices,
and time-zone information. Each script prints results locally; none uploads
data.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Other scripts in this directory cover architecture, computer name, join -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: status, service pack, uptime, CPU and memory, page file, motherboard, BIOS, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: disks, volumes, display, audio, USB devices, printers, battery, input devices, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: and time-zone information. Each script prints results locally; none uploads -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: data. -->

Scripts that collect clipboard contents, environment-variable values,
activation details, security-control status, network identifiers/settings,
installed application or update inventories, or system logs are intentionally
not included because they can expose sensitive information or security
posture. The computer-name script is read-only and does not rename the machine.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Scripts that collect clipboard contents, environment-variable values, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: activation details, security-control status, network identifiers/settings, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: installed application or update inventories, or system logs are intentionally -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: not included because they can expose sensitive information or security -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: posture. The computer-name script is read-only and does not rename the machine. -->
