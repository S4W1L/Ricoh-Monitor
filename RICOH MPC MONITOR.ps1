<#
.NOTES
    Version: 1.5
    Author: Samuel Jesus
#>

# SMTP Configuration (kept in script)
$EmailConfig = @{
    SmtpServer  = "SMTP Settings"
    SmtpPort    = PORT
    Username    = "EMAIL"
    Password    = "PASSWORD" # App Password  
    FromAddress = "ADDRESS FROM"
    ToAddress   = "DESTINATION", "DESTINATION2"
}

# OIDs
$OIDs = @{
    "Model Name"          = ".1.3.6.1.4.1.367.3.2.1.1.1.1.0"
    "Serial Number"       = ".1.3.6.1.4.1.367.3.2.1.2.1.4.0"
    "Firmware"            = ".1.3.6.1.4.1.367.3.2.1.1.1.2.0"
    "Contador"            = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.1"
    "Total Impressoes"    = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.8" # ANTIGO:1.3.6.1.4.1.367.3.2.1.2.19.2.0                       NOVO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.8
    "Total Impressao Cores mau"    = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.25" #  ANTIGO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.66       NOVO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.25    MPC300:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.60
    "Total Impressao Preto mau"    = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.26" #  ANTIGO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.67           NOVO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.26
    "Total Copias"        = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.2" # ANTIGO:1.3.6.1.4.1.367.3.2.1.2.19.4.0                       NOVO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.2
    "Total Copia Cores"   = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.17" #  ANTIGO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.62                NOVO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.17
    "Total Copia Preto"   = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.18" #  ANTIGO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.63                NOVO:1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.18
    "Black Toner Level %" = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.1"
    "Cyan Toner Level %"  = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.2"
    "Magenta Toner Level %" = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.3"
    "Yellow Toner Level %" = ".1.3.6.1.4.1.367.3.2.1.2.24.1.1.5.4"
    "Error Status"        = ".1.3.6.1.4.1.367.3.2.1.2.2.13.0"
    "Total Impressoes preto mau" = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.26"
    #"DEBUG" = ".1.3.6.1.4.1.367.3.2.1.1.1.1.0"
    #"DEBUG IMPRESSAO CORES3" = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.25" Total impressão cores a precisar de soma de 2cores para todos os modelos 
    #"DEBUG IMPRESSAO PRETO2" = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.26"  A funcionar em todos os modelos com soma de 2CORES (tem que se subtrair)
    "DEBUG 2CORES" = ".1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.10"  #2CORES tem que somar a impressão a cores e subtrair no impressão preto
    
}

function Get-PrintersConfig {
    param(
        [string]$ConfigPath = "printers_config.json"
    )
    
    # If config file doesn't exist, create a default one
    if (-not (Test-Path $ConfigPath)) {
        $defaultConfig = @(
            @{ IP = "10.10.5.200"; Community = "public" }
            @{ IP = "10.10.5.205"; Community = "public" }
        ) | ConvertTo-Json
        
        $defaultConfig | Out-File -FilePath $ConfigPath -Encoding utf8
        Write-Host "Created default configuration file at $ConfigPath" -ForegroundColor Yellow
    }
    
    try {
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json -ErrorAction Stop
        return @($config) # Ensure it's always an array
    }
    catch {
        Write-Host "Error reading configuration file: $_" -ForegroundColor Red
        exit 1
    }
}

function Get-SnmpData {
    param(
        [string]$IP,
        [string]$Community,
        [int]$MaxRetries = 3
    )
    
    $result = @{"IP Address" = $IP}
    $retryCount = 0
    $success = $false
    
    while ($retryCount -lt $MaxRetries -and -not $success) {
        try {
            $snmp = New-Object -ComObject "OlePrn.OleSNMP"
            $snmp.Open($IP, $Community, 2, 3000)
            
            foreach ($oid in $OIDs.GetEnumerator()) {
                try {
                    $value = $snmp.Get($oid.Value)
                    $result[$oid.Name] = $value
                }
                catch {
                    $result[$oid.Name] = "Error: $_"
                }
            }
            
            $snmp.Close()
            $success = $true
        }
        catch {
            $retryCount++
            if ($retryCount -eq $MaxRetries) {
                $result["Status"] = "Failed after $MaxRetries attempts"
                foreach ($oid in $OIDs.GetEnumerator()) {
                    $result[$oid.Name] = "Unavailable"
                }
            }
            Start-Sleep -Seconds 2
        }
    }


    $result['Total Impressao Cores'] = $result['Total Impressao Cores mau'] + $result['DEBUG 2CORES']
    $result['Total Cores']        = $result['Total Impressao Cores'] + $result['Total Copia Cores']
    $result['Total Impressao Preto'] = $result['Total Impressao Preto mau'] - $result['DEBUG 2CORES']
    $result['Total Preto Branco'] = $result['Total Impressao Preto'] + $result['Total Copia Preto']
    



    $printerName = if ($result["Model Name"] -and $result["Model Name"] -ne "Unavailable") { 
        $result["Model Name"] 
    } else { 
        "Unreachable Printer ($IP)" 
    }
    
    $result["Printer Name"] = $printerName
    return $result
}

function Send-EmailReport {
    param(
        [array]$PrintersData
    )
    
    $date = Get-Date -Format "dd-MM-yyyy HH:mm"
    $subject = "Ricoh Contadores - $date"
    
    $fieldOrder = @(
        'Model Name',
        'Serial Number',
        'Firmware',
        'Contador',
        #'Total Impressoes',
        #'Total Impressao Cores',
        #'Total Impressao Preto',
        #'Total Copias',
        #'Total Copia Cores',
        #'Total Copia Preto',
        'Total Cores',
        'Total Preto Branco',
        'Black Toner Level %',
        'Cyan Toner Level %',
        'Magenta Toner Level %',
        'Yellow Toner Level %',
        'Error Status'
        #'DEBUG',
        #'DEBUG IMPRESSAO CORES3',
        #'DEBUG IMPRESSAO PRETO2',
        #'DEBUG 2CORES'
    )
    
    $html = @"
<html>
<head>
<style>
    body { font-family: Arial, sans-serif; font-size: 12px; line-height: 1.2; }
    h2 { color: #ff5733; margin: 0 0 5px 0; }
    .printer { margin-bottom: 15px; }
    .unreachable { color: #888; }
    .error { color: red; }
    .bold-field { font-weight: bold; }
    p { margin:2px 0; }
</style>
</head>
<body>
<h2>HPZ Ricoh - $date</h2>
"@

    foreach ($printer in $PrintersData) {
        $isUnreachable = $printer["Status"] -eq "Failed after 3 attempts"
        $html += if ($isUnreachable) {
            "<div class='printer unreachable'>"
        } else {
            "<div class='printer'>"
        }
        
        $html += @"
<h3>$($printer['Printer Name'])</h3>
<p><strong>IP:</strong> $($printer['IP Address'])</p>
"@
        
        if ($isUnreachable) {
            $html += "<p><strong>Status:</strong> Printer unreachable after 3 attempts</p>"
        } else {
            foreach ($field in $fieldOrder) {
                if ($printer.ContainsKey($field)) {
                    $value = $printer[$field]
                    $class = if ($value -like "*Error*") { "class='error'" } else { "" }
                    $html += "<p><strong>$field</strong>: <span $class>$value</span></p>"
                }
            }
        }
        
        $html += "</div>"
    }

    $html += @"
</body>
</html>
"@

    $credential = New-Object System.Management.Automation.PSCredential (
        $EmailConfig.Username, 
        (ConvertTo-SecureString $EmailConfig.Password -AsPlainText -Force)
    )

    try {
        Send-MailMessage -From $EmailConfig.FromAddress `
                        -To $EmailConfig.ToAddress `
                        -Subject $subject `
                        -Body $html `
                        -BodyAsHtml `
                        -SmtpServer $EmailConfig.SmtpServer `
                        -Port $EmailConfig.SmtpPort `
                        -UseSsl `
                        -Credential $credential
        Write-Host "Email sent successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to send email: $_" -ForegroundColor Red
    }
}

# Main Execution
try {
    Write-Host "Starting printer monitoring..." -ForegroundColor Cyan
    
    # Get printers from config file
    $Printers = Get-PrintersConfig
    Write-Host "Loaded configuration for $($Printers.Count) printers"
    
    $allPrintersData = @()
    
    foreach ($printer in $Printers) {
        Write-Host "Checking printer at $($printer.IP)..."
        $printerData = Get-SnmpData -IP $printer.IP -Community $printer.Community
        
        if ($printerData["Status"] -eq "Failed after 3 attempts") {
            Write-Host "  Printer unreachable after 3 attempts" -ForegroundColor Yellow
        } else {
            Write-Host "  $($printerData['Printer Name']) status collected" -ForegroundColor Green
        }
        
        $allPrintersData += $printerData
    }
    
    Send-EmailReport -PrintersData $allPrintersData
    Write-Host "All printer reports completed!" -ForegroundColor Green
}
catch {
    Write-Host "Error in main execution: $_" -ForegroundColor Red
}
