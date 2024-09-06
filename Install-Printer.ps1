<#
.Synopsis
Created on:   31/12/2021
Created by:   Ben Whitmore
Edited By:    premium3722
Version:      1.0 (06.09.2024)
Filename:     Install-Printer.ps1

Simple script to install a network printer from an INF file. The INF and required CAB files hould be in the same directory as the script if creating a Win32app

#### Win32 app Commands ####

Install:
powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "PRI-DEMO-01" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf" -configfile "PRI-DEMO-01-dat"

Uninstall:
powershell.exe -executionpolicy bypass -file .\Remove-Printer.ps1 -PrinterName "PRI-DEMO-01"

Detection:
HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers\PRI-DEMO-01
Name = "PRI-DEMO-01"

.Beispiel
.\Install-Printer.ps1 -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName ""PRI-DEMO-01" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf" -configfile "PRI-DEMO-01-dat"
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
    [String]$configfile  
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
        [string]$ForegroundColor = "black"
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
        WriteLog "Staging Driver zu Windows Driver Store mit INF ""$($INFFile)"""
        WriteLog "Starte Command: Start-Process pnputil.exe -ArgumentList "$INFARGS /a" -Wait -PassThru"
        Start-Process pnputil.exe -ArgumentList "$INFARGS /a" -Wait -PassThru
    }
    Catch {
        WriteLog "Fehler Treiber konnte nicht zum Driver Store hinzugefügt werden"
        WriteLog "$($_.Exception.Message)"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {
    
        #Install driver
        $DriverExist = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
        if (-not $DriverExist) {
            WriteLog "Hinzufügen von Drucker Treiber ""$($DriverName)"""
            Add-PrinterDriver -Name $DriverName -Confirm:$false
        }
        else {
           WriteLog "Drucker Treiber ""$($DriverName)"" existiert bereits. Überspringen Treiber installation." -ForegroundColor Orange
        }
    }
    Catch {
        WriteLog "Fehler bei Installation von Drucker Treiber" -ForegroundColor Red
        WriteLog "$($_.Exception.Message)"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {

        #Create Printer Port
        WriteLog "Prüfe ob Drucker Port existiert"
        $PortExist = Get-Printerport -Name $PortName -ErrorAction SilentlyContinue
        if (-not $PortExist) {
            WriteLog "Füge Port ""$($PortName)"" hinzu"
            Add-PrinterPort -name $PortName -PrinterHostAddress $PrinterIP -Confirm:$false
        }
        else {
           WriteLog "Port ""$($PortName)"" existiert bereits. Überspringe Drucker Port installation." -ForegroundColor Orange
        }
    }
    Catch {
        WriteLog "Fehler beim erstellen von Drucker Port" -ForegroundColor Red
        WriteLog "$($_.Exception.Message)"
        WriteLog "$($_.Exception)"
        $ThrowBad = $True
    }
}

If (-not $ThrowBad) {
    Try {
        # Abschnitt 1: Drucker hinzufügen
        # Prüfen, ob der Drucker bereits existiert
        $PrinterExist = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if (-not $PrinterExist) {
            WriteLog "Füge Drucker hinzu da nicht vorhanden ""$($PrinterName)"""
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false
        }
        else {
            WriteLog "Drucker ""$($PrinterName)"" existiert bereits. Alter Drucker wird entfernt..." 
            Remove-Printer -Name $PrinterName -Confirm:$false
            WriteLog "Füge Drucker hinzu ""$($PrinterName)"""
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false
        }

        # Erneute Prüfung, ob der Drucker jetzt vorhanden ist
        $PrinterExist2 = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if ($PrinterExist2) {
            WriteLog "Drucker ""$($PrinterName)"" wurde erfolgreich hinzugefügt" -ForegroundColor Green
        }
        else {
            WriteLog "Drucker installation Fehler bei Drucker" -ForegroundColor Red
            $ThrowBad = $True
            Throw "Drucker installation Fehler bei Drucker: $PrinterName" -ForegroundColor Red
        }
    }
    Catch {
        WriteLog "Fehler Drucker konnte nicht ersteltl werden" -ForegroundColor Red
        WriteLog "$($_.Exception.Message)"
        $ThrowBad = $True
        Throw $_
    }

    # Abschnitt 2: Konfigurationsdatei hinzufügen
    if (-not $ThrowBad -and $configfile) {
        Try {
            WriteLog "Versuche Drucker Konfig zu importieren von: $configfile"
            
            # Erster Versuch, die Konfigurationsdatei hinzuzufügen
            Start-Process "RUNDLL32.EXE" -ArgumentList "PRINTUI.DLL,PrintUIEntry /Sr /n `"$PrinterName`" /a `"$configfile`" f c d g u" -Wait -NoNewWindow
            sleep 5
            WriteLog "Drucker Konfiguration erfolgreich impoortiert von: $configfile" -ForegroundColor Green
        }
        Catch {
            WriteLog "Drucker Konfiguration konnte nicht hinzugefügt werden versuche zweite Variante..." -ForegroundColor Orange
            WriteLog "$($_.Exception.Message)"
            
            # Alternativer Versuch
            Try {
                Start-Process "RUNDLL32.EXE" -ArgumentList "PRINTUI.DLL,PrintUIEntry /Sr /n `"$PrinterName`" /a `"$configfile`" 2 7 c d g u" -Wait -NoNewWindow
                sleep 5
                WriteLog "Zweite Methode: Drucker Konfiguration wurde erfolgreich importiert von: $configfile" -ForegroundColor Green
            }
            Catch {
                WriteLog "Die Drucker konfiguration konnte auf beide Möglichkeiten nicht hinzugefügt werden." -ForegroundColor Red
                WriteLog "$($_.Exception.Message)"
            }
        }
    }
}


If ($ThrowBad) 
{
    WriteLog "Ein Fehler während der Installation. Hier befindet sich das Log: $Logfile"
    WriteLog "Installation Fehlgeschlagen"
}
else
{
    WriteLog "Script erfolgreich beendet"
}
