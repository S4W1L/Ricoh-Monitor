# Ricoh Printer SNMP Monitoring Tool

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue.svg)  
![Deployment](https://img.shields.io/badge/Schedule-Monthly%201st-brightgreen)
![](https://img.shields.io/badge/Self_Made-Windows-blue)

A compiled PowerShell executable that monitors Ricoh MPC printers via SNMP and sends scheduled email reports with toner levels and page counts.

## Features

- Monthly automated execution (1st day of month)
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

### Configuration:

Edit C:\Printer Monitor\printers_config.json:


```
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
```


### SMTP Setup (Edit EXE with PowerShell if needed):

Open the PS1 source

Modify the $EmailConfig block

```powershell
$EmailConfig = @{
    SmtpServer  = "your.smtp.server"
    SmtpPort    = 587
    Username    = "your@email.com"
    Password    = "your-app-password"  
    FromAddress = "from@email.com"
    ToAddress   = "recipient1@email.com", "recipient2@email.com"
```

Recompile


### Schedule the Task (Admin PowerShell):

```powershell
$Action = New-ScheduledTaskAction -Execute "C:\Printer Monitor\APP_NAME.exe"
$Trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "9:00AM"
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd

Register-ScheduledTask `
    -TaskName "Monthly Ricoh Report" `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -RunLevel Highest `
    -Description "Executes monthly printer monitoring on the 1st at 9:00 AM"
```

### Data Collected

Metric	OID Reference
- Model Name	.1.3.6.1.4.1.367.3.2.1.1.1.1.0
- Serial Number	.1.3.6.1.4.1.367.3.2.1.2.1.4.0
- Toner Levels (CMYK)	.1.3.6.1.4.1.367.3.2.1.2.24.1.1.5
- Page Counts	.1.3.6.1.4.1.367.3.2.1.2.19.5.1.9
- Error Status	.1.3.6.1.4.1.367.3.2.1.2.2.13.0

### Troubleshooting

##### Printers not responding:

Verify SNMP is enabled on printers

Check community strings match

Test connectivity (Test-NetConnection <IP> -Port 161)

##### Email failures:

Verify SMTP credentials

Check TLS requirements for your email provider

Review Windows Event Logs for errors

##### First-run issues:

Run EXE manually to generate config file

Ensure folder has write permissions


