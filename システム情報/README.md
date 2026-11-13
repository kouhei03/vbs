# vbs

## System information scripts

The following scripts query local OS, hardware, storage, and peripheral
information through WMI or built-in Windows Script Host APIs. Run an
individual script from Command Prompt:

```text
cscript //nologo GetOSInfo.vbs
```

Other scripts in this directory cover architecture, computer name, join
status, service pack, uptime, CPU and memory, page file, motherboard, BIOS,
disks, volumes, display, audio, USB devices, printers, battery, input devices,
and time-zone information. Each script prints results locally; none uploads
data.

Scripts that collect clipboard contents, environment-variable values,
activation details, security-control status, network identifiers/settings,
installed application or update inventories, or system logs are intentionally
not included because they can expose sensitive information or security
posture. The computer-name script is read-only and does not rename the machine.
