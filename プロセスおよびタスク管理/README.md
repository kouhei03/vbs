# Process information scripts

These scripts provide read-only information about processes on the local
computer. Run an individual script from Command Prompt:

```text
cscript //nologo GetRunningProcesses.vbs
cscript //nologo TestProcessRunning.vbs notepad.exe
cscript //nologo GetProcessMemoryUsage.vbs
```

The scripts list process names and PIDs, check for a process by exact name,
and report working-set and virtual-memory sizes. They do not terminate,
launch, modify, or remotely control processes, and they do not send or save
the collected information.

Process termination, bulk termination, automatic restart or scheduling,
arbitrary command execution, elevation, remote control, software
installation/removal, stress testing, hidden-process inspection, and memory
dumps are intentionally not included.

Additional read-only diagnostics include per-process CPU counters (which can
exceed 100% across cores) and process creation times. Service exit codes and
pending states can also be inspected. `ExportServiceInventory.vbs
<new-csv-path>` writes only a local CSV report and refuses to overwrite an
existing file; it does not change service configuration.

## Windows service information scripts

These service queries are read-only and operate on the local computer:

```text
cscript //nologo GetWindowsServices.vbs
cscript //nologo GetServiceStatus.vbs Spooler
cscript //nologo GetServiceStartupType.vbs Spooler
cscript //nologo CheckMiddlewareServices.vbs
cscript //nologo WaitForServiceRunning.vbs Spooler 60
```

They list services, query one service by its exact service name, check common
SQL Server/IIS service states, and wait up to 300 seconds for a service to
reach `Running`. Invalid arguments, missing services, and timeouts are
reported with non-zero exit codes. They do not start, stop, reconfigure, or
remotely control services.

Service startup/recovery/account changes, remote operations, service
installation/removal, security-service enforcement, and automatic recovery
or maintenance actions are intentionally not included.
