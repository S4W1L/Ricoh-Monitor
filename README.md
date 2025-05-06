# Ricoh Printer SNMP Monitoring Tool

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue.svg)  
![License](https://img.shields.io/badge/License-MIT-green.svg)

A compiled PowerShell executable that monitors Ricoh MPC printers via SNMP and sends scheduled email reports with toner levels and page counts.

## Features

- Compiled EXE for easy deployment
- Collects printer metrics via SNMP (v2c)
- Tracks toner levels (CMYK) and page counts
- Scheduled email reports (HTML format)
- Central configuration via JSON file
- Automatic retry mechanism (3 attempts)
- Custom branding (HPZ)

## System Requirements

- Windows 10/11 or Windows Server 2016+
- .NET Framework 4.7.2 or later
- Network access to printers on SNMP port (161)
- SMTP server access for email notifications

## Deployment

1. **Compile the Script** (Admin PowerShell):
   ```powershell
   Invoke-PS2EXE Invoke-PS2EXE -InputFile "NAME_OF_SCRIPT.ps1" -OutputFile "APP_NAME.exe" -IconFile "ICON.ico" -Title "TITLE" -Company "COMPANY" -Product "PRODUCT" -Description "DESCRIPTION OF APP"

## Installation:

Create folder: C:\Printer Monitor\

Place these files in the folder:

- MPC Monitor.exe
- printers_config.json (auto-created on first run if missing)

Configuration:
Edit C:\Printer Monitor\printers_config.json:

json
[
    {
        "IP": "10.10.5.200",
        "Community": "public"
    },
    {
        "IP": "10.10.5.205",
        "Community": "private"
    }
]

SMTP Setup (Edit EXE with PowerShell if needed):

Open the PS1 source

Modify the $EmailConfig block

Recompile

Schedule the Task (Admin PowerShell):

```powershell
$Action = New-ScheduledTaskAction -Execute "C:\HPZ Printer Monitor\MPC Monitor.exe"
$Trigger = New-ScheduledTaskTrigger -Daily -At "8:00AM"
Register-ScheduledTask -TaskName "HPZ Ricoh Monitor" -Action $Action -Trigger $Trigger -RunLevel Highest
