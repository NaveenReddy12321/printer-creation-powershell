# ==========================================================
# BULK PRINTER CREATION FROM CSV (LOCAL SERVER)
# ==========================================================

# -------------------------------
# ADMIN CHECK
# -------------------------------
$principal = New-Object Security.Principal.WindowsPrincipal(
    [Security.Principal.WindowsIdentity]::GetCurrent()
)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
    exit
}

# -------------------------------
# CSV PATH
# -------------------------------
$CsvPath = "C:\Temp\printers.csv" #Change accordingly

if (-not (Test-Path $CsvPath)) {
    Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
    exit
}

$Printers = Import-Csv $CsvPath

if ($Printers.Count -eq 0) {
    Write-Host "CSV file is empty. Nothing to process." -ForegroundColor Yellow
    exit
}

# -------------------------------
# LOGGING SETUP
# -------------------------------
$LogFolder = "C:\PrinterLogs"

if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder | Out-Null
}

$LogFile = "$LogFolder\PrinterCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$LogData = @()

function Write-Log {
    param (
        [string]$PrinterName,
        [string]$PrinterIP,
        [string]$PortName,
        [string]$Status,
        [string]$Message
    )

    $script:LogData += [PSCustomObject]@{
        Time        = Get-Date
        PrinterName = $PrinterName
        PrinterIP   = $PrinterIP
        PortName    = $PortName
        Status      = $Status
        Message     = $Message
    }
}

# -------------------------------
# PROCESS PRINTERS
# -------------------------------
foreach ($Printer in $Printers) {

    Write-Host "`nProcessing printer: $($Printer.PrinterName)" -ForegroundColor Cyan

    try {
        # -------------------------------
        # CHECK / CREATE PORT
        # -------------------------------
        if (-not (Get-PrinterPort -Name $Printer.PortName -ErrorAction SilentlyContinue)) {

            Write-Host "Port not found. Creating port: $($Printer.PortName)" -ForegroundColor Yellow

            Add-PrinterPort `
                -Name $Printer.PortName `
                -PrinterHostAddress $Printer.PrinterIP `
                -ErrorAction Stop
        }

        # -------------------------------
        # CHECK PRINTER EXISTS
        # -------------------------------
        if (Get-Printer -Name $Printer.PrinterName -ErrorAction SilentlyContinue) {
            Write-Host "Printer already exists. Skipping." -ForegroundColor Yellow

            Write-Log `
                -PrinterName $Printer.PrinterName `
                -PrinterIP $Printer.PrinterIP `
                -PortName $Printer.PortName `
                -Status "Skipped" `
                -Message "Printer already exists"

            continue
        }

        # -------------------------------
        # CREATE PRINTER
        # -------------------------------
        Add-Printer `
            -Name $Printer.PrinterName `
            -PortName $Printer.PortName `
            -DriverName $Printer.DriverName `
            -Comment $Printer.Comment `
            -Location $Printer.Description `
            -ErrorAction Stop

        # -------------------------------
        # SHARE PRINTER (IF REQUIRED)
        # -------------------------------
        if ($Printer.Shared -eq "Yes") {
            Set-Printer -Name $Printer.PrinterName -Shared $true
        }

        Write-Host "Printer created successfully." -ForegroundColor Green

        Write-Log `
            -PrinterName $Printer.PrinterName `
            -PrinterIP $Printer.PrinterIP `
            -PortName $Printer.PortName `
            -Status "Success" `
            -Message "Printer created successfully"
    }
    catch {
        Write-Host "Failed to create printer." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed

        Write-Log `
            -PrinterName $Printer.PrinterName `
            -PrinterIP $Printer.PrinterIP `
            -PortName $Printer.PortName `
            -Status "Failed" `
            -Message $_.Exception.Message
    }
}

# -------------------------------
# EXPORT LOG FILE
# -------------------------------
if ($LogData.Count -gt 0) {
    $LogData | Export-Csv -Path $LogFile -NoTypeInformation
    Write-Host "`nLog file created at: $LogFile" -ForegroundColor Cyan
}

Write-Host "`nPrinter creation process completed." -ForegroundColor Green
