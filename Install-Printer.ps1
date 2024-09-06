<#
.Synopsis
Created on:   31/12/2021
Created by:   Ben Whitmore
Edited By:    xPreo
Filename:     Install-Printer.ps1

Simple script to install a network printer from an INF file. The INF and required CAB files hould be in the same directory as the script if creating a Win32app

#### Win32 app Commands ####

Install:
powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf"

Uninstall:
powershell.exe -executionpolicy bypass -file .\Remove-Printer.ps1 -PrinterName "Canon Printer Upstairs"

Detection:
HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers\Canon Printer Upstairs
Name = "Canon Printer Upstairs"

.Example
.\Install-Printer.ps1 -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf"
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True)]
    [String]$PortName,
    [Parameter(Mandatory = $True)]
    [String]$PrinterIP,
    [Parameter(Mandatory = $True)]
    [String]$PrinterName,
    [Parameter(Mandatory = $True)]
    [String]$DriverName,
    [Parameter(Mandatory = $True)]
    [String]$INFFile,
    [Parameter(Mandatory = $False)]
    [String]$configfile  # Neuer optionaler Parameter
)

#Reset Error catching variable
$Throwbad = $Null

#Run script in 64bit PowerShell to enumerate correct path for pnputil
If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH -PortName $PortName -PrinterIP $PrinterIP -DriverName $DriverName -PrinterName $PrinterName -INFFile $INFFile
    }
    Catch {
        Write-Error "Failed to start $PSCOMMANDPATH"
        Write-Warning "$($_.Exception.Message)"
        $Throwbad = $True
    }
}
###################################################Variablen Setzen##############################################################################                                                                                                                                                                                                 
# Variablen
$LogfilePath = "C:\temp\Intune\logs"
$Logfile = "C:\temp\Intune\logs\$PrinterName.html"
$start_time = Get-Date

# Log Funktion
function WriteLog {
    Param (
        [string]$LogString,
        [string]$ForegroundColor = "black",
        [bool]$IsBold = $false
    )
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp <span style='color: $ForegroundColor;'>$LogString</span><br>"
    Add-Content $Logfile -Value $LogMessage
}

# Initialisierung des Logs
Write-Host "Script start"
WriteLog "Script start" -ForegroundColor Black

if (!(Test-Path $LogfilePath)) { New-Item -Path $LogfilePath -ItemType Directory -Force }
if (!(Test-Path $Logfile)) {
    New-Item -Path $Logfile -ItemType File -Force
    WriteLog "Logfile wird erstellt" -ForegroundColor Green
} else {
    Remove-Item $Logfile -Force
    WriteLog "Logfile wird gelöscht" -ForegroundColor Green
    New-Item -Path $Logfile -ItemType File -Force
    WriteLog "Logfile wird erstellt" -ForegroundColor Green
}


WriteLog "##################################"
WriteLog "Installation started"
WriteLog "##################################"
WriteLog "Install Printer using the following values..."
WriteLog "Port Name: $PortName"
WriteLog "Printer IP: $PrinterIP"
WriteLog "Printer Name: $PrinterName"
WriteLog "Driver Name: $DriverName"
WriteLog "INF File: $INFFile"

$INFARGS = @(
    "/add-driver"
    "$INFFile"
)

If (-not $ThrowBad) {

    Try {

        #Stage driver to driver store
        WriteLog "Staging Driver to Windows Driver Store using INF ""$($INFFile)"""
        WriteLog "Running command: Start-Process pnputil.exe -ArgumentList $($INFARGS) -wait -passthru"
        Start-Process pnputil.exe -ArgumentList "$INFARGS /a" -Wait -PassThru
    }
    Catch {
        WriteLog "Error staging driver to Driver Store"
        WriteLog "$($_.Exception.Message)"
        WriteLog "Error staging driver to Driver Store"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {
    
        #Install driver
        $DriverExist = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
        if (-not $DriverExist) {
            WriteLog -Stamp -Value "Adding Printer Driver ""$($DriverName)"""
            Add-PrinterDriver -Name $DriverName -Confirm:$false
        }
        else {
           WriteLog "Print Driver ""$($DriverName)"" already exists. Skipping driver installation."
        }
    }
    Catch {
        WriteLog "Error installing Printer Driver"
        WriteLog "$($_.Exception.Message)"
        WriteLog "Error installing Printer Driver"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {

        #Create Printer Port
        $PortExist = Get-Printerport -Name $PortName -ErrorAction SilentlyContinue
        if (-not $PortExist) {
            WriteLog "Adding Port ""$($PortName)"""
            Add-PrinterPort -name $PortName -PrinterHostAddress $PrinterIP -Confirm:$false
        }
        else {
           WriteLog "Port ""$($PortName)"" already exists. Skipping Printer Port installation."
        }
    }
    Catch {
        WriteLog "Error creating Printer Port"
        WriteLog "$($_.Exception.Message)"
        WriteLog "Error creating Printer Port"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {

        #Add Printer
        $PrinterExist = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if (-not $PrinterExist) {
            WriteLog "Adding Printer ""$($PrinterName)"""
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false
        }
        else {
            WriteLog "Printer ""$($PrinterName)"" already exists. Removing old printer..."
            Remove-Printer -Name $PrinterName -Confirm:$false
            WriteLog "Adding Printer ""$($PrinterName)"""
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false
        }

        $PrinterExist2 = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if ($PrinterExist2) {
           WriteLog "Printer ""$($PrinterName)"" added successfully"
            
            # Falls $configfile nicht leer ist, führe den Code aus
            if ($configfile) {
                WriteLog "Importing Printer Configuration from file: $configfile"
                Start-Process "RUNDLL32.EXE" -ArgumentList "PRINTUI.DLL,PrintUIEntry /Sr /n `"$PrinterName`" /a `"$env:USERPROFILE\Desktop\Printer\$configfile.dat`" f c d g u" -NoNewWindow -Wait
            }
        }
        else {
            WriteLog "Error creating Printer"
            WriteLog "Printer ""$($PrinterName)"" error creating printer"
            $ThrowBad = $True
        }
    }
    Catch {
        WriteLog "Error creating Printer"
        WriteLog "$($_.Exception.Message)"
        WriteLog "Error creating Printer"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If ($ThrowBad) {
    WriteLog "An error was thrown during installation. Installation failed. Refer to the log file in %temp% for details"
    WriteLog "Installation Failed"
}
