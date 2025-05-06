# Ricoh Printer SNMP Monitoring Tool

![PowerShell](https://img.shields.io/badge/PowerShell-v5.1+-blue.svg)

A PowerShell script that monitors Ricoh printers via SNMP and sends email reports with toner levels, page counts, and status information.

## Features

- Collects printer metrics via SNMP (v2c)
- Tracks toner levels (CMYK)
- Monitors page counts (color, black & white, copies)
- Detects printer errors
- Generates HTML email reports
- Supports multiple printers via JSON configuration
- Automatic retry mechanism for unreliable connections

## Prerequisites

- PowerShell 5.1 or later
- SNMP access to printers (community string)
- SMTP server credentials for email notifications
- Windows system with SNMP COM object available (`OlePrn.OleSNMP`)

## Installation

1. Clone or download the script to your preferred location
2. Ensure PowerShell execution is allowed:
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
