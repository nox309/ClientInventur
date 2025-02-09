<#
.NAME
    CL Inventur
.SYNOPSIS
    Liest div. Hardware Informantionen aus und speichert diese in einer CSV

.NOTES
    Author: nox309 | Torben Inselmann
    Email: support@inselmann.it
    Git: https://github.com/nox309
    Version: 1.0
    DateCreated: 2023/06/01
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
$WID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$Prp = New-Object System.Security.Principal.WindowsPrincipal($WID)
$Adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
$IsAdmin = $Prp.IsInRole($Adm)
if( !$IsAdmin ){
    Write-Host -ForegroundColor Red "The script does not have enough rights to run. Please start with admin rights!"
    break
    }

if (!(Get-Command write-log -ErrorAction SilentlyContinue)) {
    # Install the PsIni module from the PSGallery
    if (!(Get-Module -ListAvailable myPosh_write-log)) {
        Write-Host "Installing myPosh_write-log module from PSGallery..."
        Install-PackageProvider -Name NuGet -Force -Confirm:$false
        Install-Module -Name myPosh_write-log -Repository PSGallery -Force -Confirm:$false -Scope AllUsers
    }
    #Make sure that the function is now available.
    Import-Module myPosh_write-log
}
function Pause ($Message="Warte auf bestätigung ...")
{
  Write-Host -NoNewLine $Message
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  Write-Host ""
}

# Funktion zum Aktualisieren des Fortschrittsbalkens
function Update-Progress($i, $Tasks) {
    $percentage = ($i / $Tasks) * 100
    $percentage = [Math]::Round($percentage, 2)
    #$progress = "[{0}{1}] {2}%" -f ('#' * ($percentage / 10)), (' ' * ((100 - $percentage) / 10)), $percentage
    $progressText = "Informationssammlung für die Client-Inventur"
    Write-Progress -Activity $progressText -Status "In Progress" -PercentComplete (($i / $Tasks) * 100)
}
Set-ExecutionPolicy -ExecutionPolicy Bypass
#---------------------------------------------------------[Logic]--------------------------------------------------------
Write-Log "Starte vorbereitung für die Client Inventur" -Severity Information -console $true
Write-Log "Sammle Informationen zum Client" -Severity Information -console $true
$Tasks = 11  # Anzahl der Aufgaben


# Fortschritt initialisieren
$i = 0
Update-Progress $i $Tasks

#1. CL Namen auslesen
$CLname = $ENV:COMPUTERNAME
Write-Log "Client Name: $Clname" -Severity Information -console $true
$i++
Update-Progress $i $Tasks


#2. OSinfos auslesen auslesen
$computerInfo = Get-ComputerInfo
$CLos = ($computerInfo).OSName
$CLosbuild = ($computerInfo).OsBuildNumber
$CLInstallDate = ($computerInfo).OsInstallDate
Write-Log "Betriebsystem Name: $CLos" -Severity Information -console $true
Write-Log "Betriebsystem Build: $CLosbuild" -Severity Information -console $true
Write-Log "Betriebsystem Installations Datum: $CLInstallDate" -Severity Information -console $true
$i++
Update-Progress $i $Tasks


#3. Mac Addresse auslesen auslesen auslesen
$mac = (Get-NetAdapter -Physical | Where-Object Status -eq 'Up').MacAddress
$CLmac = $mac -replace '-', ':'
if (($CLmac).count -eq 2) {
    Write-Log "2 Aktive Mac Adressen gefunden bitte WLAN abschalten und bestätigen" -Severity Warning -console $true
    Pause
    $mac = (Get-NetAdapter -Physical | Where-Object Status -eq 'Up').MacAddress
    $CLmac = $mac -replace '-', ':'
    if (($CLmac).count -eq 2) {
        Write-Log "2 Aktive Mac Adressen gefunden bitte WLAN abschalten und bestätigen" -Severity Warning -console $true
        Write-Log "2 Aktive Mac Adressen gefunden die Erste Adresse wird nun genommen als Primäre MAC Adresse" -Severity Warning -console $true
        $CLmac = $CLmac[0]
    }
}
Write-Log "Mac Adresse: $CLmac" -Severity Information -console $true
# Fortschritt aktualisieren
$i++
Update-Progress $i $Tasks


#4. Serien Nummer auslesen 
$CLseriealNumer = ($computerInfo).BiosSeralNumber
Write-Log "Serien Nummer: $CLseriealNumer" -Severity Information -console $true
$i++
Update-Progress $i $Tasks


#5. Hersteller auslesen 
$CLHersteller = ($computerInfo).BiosManufacturer
Write-Log "Hersteller: $CLHersteller" -Severity Information -console $true
$i++
Update-Progress $i $Tasks

#6. Modell auslesen 
$CLmodel = ($computerInfo).CsModel
$CLSystemFamily = ($computerInfo).CsSystemFamily
Write-Log "Model: $CLmodel" -Severity Information -console $true
Write-Log "Family: $CLSystemFamily" -Severity Information -console $true
$i++
Update-Progress $i $Tasks


#7. Windows Key auslesen
# BIOS-Schlüssel abrufen
$BiosKey = (Get-WmiObject -Class SoftwareLicensingService).OA3xOriginalProductKey

# Registrierungs-Schlüssel abrufen
$RegistryKey = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name DigitalProductId).DigitalProductId
$WindowsKey = ([System.Text.Encoding]::ASCII.GetString($RegistryKey[52..66])).Replace("`0", "")

# Überprüfen und Priorisieren der Schlüssel
if ([string]::IsNullOrEmpty($BiosKey)) {
    $WindowsKey = $WindowsKey.Trim()
    if ([string]::IsNullOrEmpty($WindowsKey)) {
        $WindowsKey = "Kein Produktschlüssel gefunden."
    }
} else {
    $WindowsKey = $BiosKey
}

# Ergebnis in Variable speichern
$CLProductKey = $WindowsKey
Write-Log "Windows Key: $CLProductKey" -Severity Information -console $true
$i++
Update-Progress $i $Tasks


#8. Treiber exportieren
Write-Log "Treiber werden ermittelt und exportiert" -Severity Information -console $true
Export-WindowsDriver -Online -Destination .\Treiber\$CLos\$CLmodel\$CLname\ | Out-Null
Write-Log "Treiber für das Model wurden Exportiert nach .\Treiber\$CLos\$CLmodel\$CLname\" -Severity Information -console $true
$i++
Update-Progress $i $Tasks

#9. Hardware Infos einsammeln
$physicalDisk = Get-PhysicalDisk -DeviceNumber 0
$CLprocessorname = (Get-CimInstance Win32_Processor).Name
Write-Log "CPU Name/Model: $CLprocessorname" -Severity Information -console $true
$CLCPUCore = (Get-WmiObject -Class Win32_Processor -ComputerName.).NumberOfCores
Write-Log "Anzahl Core: $CLCPUCore" -Severity Information -console $true
$CLCPULogigCore = (Get-WmiObject -Class Win32_Processor -ComputerName.).ThreadCount
Write-Log "Anzahl Logische Cores: $CLCPULogigCore" -Severity Information -console $true
$Clmemory = [Math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
Write-Log "Größe Ram: $Clmemory" -Severity Information -console $true
$CLOSDiskSize = "{0:N2}" -f ($physicalDisk.size / 1GB)
Write-Log "Größe der System Platte: $CLOSDiskSize" -Severity Information -console $true
$CLOSDiskfreeSize = "{0:N2}" -f ((Get-PSDrive -Name C).Free / 1GB)
Write-Log "Freier Speicher der System Platte: $CLOSDiskfreeSize" -Severity Information -console $true
$CLOSDiskType = $physicalDisk.MediaType
Write-Log "OS Disk ist eine $CLOSDiskType" -Severity Information -console $true
$CLOSDiskBUSType = $physicalDisk.BusType
Write-Log "Der BusType der OS Disk lautet $CLOSDiskBUSType" -Severity Information -console $true
$i++
Update-Progress $i $Tasks


#10. Windows 11 Kompalibilitätsprüfung
$json = .\HardwareReadiness.ps1
$object = $json | ConvertFrom-Json -ErrorAction SilentlyContinue
# Variablen zuweisen
$returnResult = $object.returnResult
$logging = $object.logging
$returnCode = $object.returnCode
$CLreturnReason = $object.returnReason
if ($returnCode -eq "1") {$CLWin11Upgrade = "inkompatibel"}
if ($returnCode -eq "0") {$CLWin11Upgrade = "kompatibel"}
if ($returnCode -eq "-1") {$CLWin11Upgrade = "unbestimmt"}
if ($returnCode -eq "-2") {$CLWin11Upgrade = "Fehler bei der Prüfung"}

Write-Log "Pürfung Windows 11 Kompatibilität hat ergeben: $CLWin11Upgrade" -Severity Information -console $true
Write-Log "ReturnCode der Prüfung: $returnCode" -Severity Information -console $true
Write-Log "Begründung bei fehlschlag der Prüfung: $CLreturnReason" -Severity Information -console $true
Write-Log "Ausgabe der Prüfung" -Severity Debug -console $false
Write-Log "$json" -Severity Debug -console $false

$i++
Update-Progress $i $Tasks


#11. Informationen in CSV schreiben
$CSVFile = ".\Client_Infos.csv"  # Pfad zur vorhandenen CSV-Datei

# Überprüfen, ob die CSV-Datei vorhanden ist
if (Test-Path $CSVFile) {
    # Erstellen eines benutzerdefinierten Objekts mit den Variablenwerten
   <#
    $CustomObject = [PSCustomObject]@{}
    $clVariables = Get-Variable -Name "CL*"

    foreach ($variable in $clVariables) {
        $variableName = $variable.Name
        $variableValue = $variable.Value
        $CustomObject | Add-Member -NotePropertyName $variableName -NotePropertyValue $variableValue
    }
    $CSVExport = $CustomObject | Sort-Object -Property Name
   #>
    $CustomObject = [PSCustomObject]@{
        CLname         = $CLname
        CLos           = $CLos
        CLKey          = $CLProductKey
        CLosbuild      = $CLosbuild
        CLInstallDate  = $CLInstallDate
        CLmac          = $CLmac
        CLseriealNumer = $CLseriealNumer
        CLHersteller   = $CLHersteller
        CLmodel        = $CLmodel
        CLSystemFamily = $CLSystemFamily
        CLCPUName      = $CLprocessorname
        CLCPUCore      = $CLCPUCore
        CLLogiCPUCore  = $CLCPULogigCore
        CLMemory       = $Clmemory
        CLOSDiskSize   = $CLOSDiskSize
        CLOSDiskfreeSize = $CLOSDiskfreeSize
        CLOSDiskType   = $CLOSDiskType
        CLOSDiskBUSType = $CLOSDiskBUSType
        CLWin11        = $CLWin11Upgrade
        CLWin11Reason  = $returnReason
    }
 
    # Anhängen des benutzerdefinierten Objekts an die CSV-Datei
    $CustomObject | Export-Csv -Path $CSVFile -Append -NoTypeInformation -Delimiter ';'
    Write-Log "Informationen in CSV exportiert" -Severity Information -console $true

} else {
    Write-Log "Informationen konnten nicht in CSV exportiert die Datei ist nicht vorhanden" -Severity Warning -console $true
}
$i++
Update-Progress $i $Tasks

# Fortschrittsbalken abschließen
Write-Progress -Activity "Fortschritt" -Completed
Write-Log "Ermittlung der Daten abgeschlossen" -Severity Information -console $true
$date = (get-date -Format yyyyMMdd)
$filename = $date + "_logfile.txt"
$logpath = "$env:SystemDrive\tmp\myPosh_Log\"
$log = $logpath + $filename
Copy-Item -Path $log -Destination .\logs\$ENV:COMPUTERNAME.txt