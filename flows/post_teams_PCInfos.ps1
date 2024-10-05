# Dieses Script sammelt PC infos und sendet sie an Teams per Webhook
# Vorbereitungen:
#1 Powerautomate erstellen mit Warten auf Webhook von Teams
#2 Karte in einem Chat oder Kanalveröffentlichen
#3 Karten Vorlagen: https://adaptivecards.io/designer

$webhookUrl = "https://prod-201.westeurope.logic.azure.com:443/workflows/XXXXXXXXXXXXXXXXXXXX"

function Get-Administrators {
    $admins = net localgroup administratoren | Where-Object {$_ -AND $_ -notmatch "Der Befehl *"} | Select-Object -Skip 4
    return $admins -join ", "  # Administratoren mit Komma und Leerzeichen zusammenfügen
}

function Get-BitLockerStatus {
    $bitlockerStatus = (Get-BitLockerVolume -MountPoint "C:").ProtectionStatus
    if ($bitlockerStatus -eq 1) {
        return "Aktiviert"
    } else {
        return "Deaktiviert"
    }
}

function Get-RDPStatus {
    $rdpStatus = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections').fDenyTSConnections
    if ($rdpStatus -eq 0) {
        return "Aktiviert"
    } else {
        return "Deaktiviert"
    }
}

function Get-ComputerInfoDetails {
    Write-Host "Start Get-ComputerInfoDetails" -ForegroundColor Green
    $computerInfo = Get-ComputerInfo
    return @{
        csModel = $computerInfo.CsModel
        csDomain = $computerInfo.CsDomain
        biosReleaseDate = $computerInfo.BiosReleaseDate
        biosSerialNr = $computerInfo.BiosSeralNumber
        windowsProductName = $computerInfo.OsName
        windowsVersion = $computerInfo.WindowsBuildLabEx
    }
    Write-Host "Nach Get-ComputerInfoDetails $ramInfoString" -ForegroundColor Green
}

function Get-RAM {
    # Abrufen der gesamten RAM-Kapazität in Bytes und Umwandlung in GB
    $RAM1 = Get-WmiObject -Class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | Select-Object -ExpandProperty Sum
    $RAM = [math]::Round($RAM1 / 1GB, 2)
    Write-Host "$RAM"
    # Rückgabe des Ergebnisses als String
    return "$RAM"
}


function Get-AntivirusInfo {
  $antivirusInfo = Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct | Select-Object -ExpandProperty displayName
  $antivirusInfo = $antivirusInfo | ForEach-Object { $_.Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{IsCombiningDiacriticalMarks}+' }
  $antivirusInfo = $antivirusInfo | Select-Object -Unique
  $antivirusString = $antivirusInfo -join ", "
  return $antivirusString
}

function Get-Officeversion 
{
    # Pfad für 32-Bit Office auf 64-Bit Windows
    $OfficeRegPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    
    # Pfad für 64-Bit Office auf 64-Bit Windows oder 32-Bit Office auf 32-Bit Windows
    if (-not (Test-Path $OfficeRegPath)) 
    {
        $OfficeRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    }

    # Funktion zur Ermittlung der Office-Version in den installierten Programmen
    function Get-OfficeVersionFromInstalledPrograms {
        $installedPrograms = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
                            , "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $officeProduct = $installedPrograms | Where-Object {
            $_.DisplayName -match "Microsoft Office" -or $_.DisplayName -match "Office"
        }

        if ($officeProduct) {
            foreach ($product in $officeProduct) {
                if ($product.DisplayName -match "2013") { return "Office 2013" }
                elseif ($product.DisplayName -match "2016") { return "Office 2016" }
                elseif ($product.DisplayName -match "2019") { return "Office 2019" }
                elseif ($product.DisplayName -match "2021") { return "Office 2021" }
                elseif ($product.DisplayName -match "365") { return "Office 365" }
            }
        }
        return $null
    }

    # Abfrage der Office-Version
    $OfficeVersion = Get-ItemProperty -Path $OfficeRegPath -Name VersionToReport, ProductReleaseIds -ErrorAction SilentlyContinue

    if ($OfficeVersion) 
    {
        # Produktname bestimmen
        $ProductName = $OfficeVersion.ProductReleaseIds

        # Office Version bestimmen
        switch -Wildcard ($ProductName) 
        {
            "*O365*" { $OfficeVersionDetected = "Office 365" }
            "*O365ProPlusRetail*" { $OfficeVersionDetected = "Office 365 ProPlus" }
            "*Professional2021*" { $OfficeVersionDetected = "Office 2021" }
            "*HomeBusiness2021Retail*" { $OfficeVersionDetected = "Office 2021" }
            "*Professional2019*" { $OfficeVersionDetected = "Office 2019" }
            "*Professional2016*" { $OfficeVersionDetected = "Office 2016" }
            default 
            { 
                if ($OfficeVersion.VersionToReport -like "16.0.*") 
                {
                    Write-Host "Suche in installierten Programmen..."
                    $OfficeVersionDetected = Get-OfficeVersionFromInstalledPrograms
                    if ($OfficeVersionDetected) 
                    {
                        Write-Host "Installierte Office-Version (aus Programmen ermittelt): $OfficeVersionDetected"
                        #Ninja-Property-Set officeversion "$OfficeVersionDetected"
                    }
                    else 
                    {
                        Write-Host "Keine Office-Version gefunden."
                    }
                }
            }
        }
        Write-Host "Installierte Office-Version: $OfficeVersionDetected ($($OfficeVersion.VersionToReport))"
        return "$OfficeVersionDetected ($($OfficeVersion.VersionToReport))"
    } 
    else 
    {
        Write-Host "Office-Version konnte nicht ermittelt werden."
        return "Unbekannt"
    }
}


function Get-WindowsDisplayVersion {
    return Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Select-Object -ExpandProperty DisplayVersion
}

function Send-WebhookNotification {
    param (
        [string]$webhookUrl,
        [string]$pcName,
        [string]$adminsString,
        [string]$bitlockerStatus,
        [string]$rdpStatus,
        [hashtable]$computerDetails,
        [string]$RAM,
        [string]$antivirusString,
        [string]$displayVersion,
        [string]$officeversionstring
    )

    $dataJson = @{
        title = "Administrator List"
        description = "Folgender PC hat noch andere Administratoren als der Localadmin."
        pcName = $pcName
        adminsString = $adminsString
        bitlockerStatus = $bitlockerStatus
        rdpStatus = $rdpStatus  
        csModel = $computerDetails.csModel
        csDomain = $computerDetails.csDomain
        biosReleaseDate = $computerDetails.biosReleaseDate
        biosSerialNr = $computerDetails.biosSerialNr
        windowsProductName = $computerDetails.windowsProductName
        windowsVersion = $computerDetails.windowsVersion
        RAM = $RAM  # Hier der gleiche Name
        antivirus = $antivirusString
        displayversion = $displayVersion
        officeversionstring = $officeversionstring
    }

    $AdaptiveCard = @{
        type = "message"
        attachments = @(
            @{
                contentType = "application/vnd.microsoft.card.adaptive"
                contentUrl = $null
                content = @{
                    "$schema" = "http://adaptivecards.io/schemas/adaptive-card.json"
                    type = "AdaptiveCard"
                    version = "1.0"
                    body = @(
                        # Hier fügen wir die Elemente der Adaptive Card hinzu
                        @{
                            type = "TextBlock"
                            size = "Large"
                            weight = "Bolder"
                            text = $dataJson.title
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.description
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.pcName
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.adminsString
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.bitlockerStatus
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.biosSerialNr
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.rdpStatus
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.csModel
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.csDomain
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.biosReleaseDate
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.windowsProductName
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.RAM  # Korrekte Verwendung
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.antivirus
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.displayversion
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.officeversionstring
                            wrap = $true
                        }
                    )
                }
            }
        )
    }

    $AdaptiveCardJson = $AdaptiveCard | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method POST -Uri $webhookUrl -Body $AdaptiveCardJson -ContentType 'application/json'
}

function Write-HostSystemInfo {
    param (
        [hashtable]$dataJson
    )

    Write-Host "Windows Version: " + $dataJson.windowsVersion
    Write-Host "Windows Product Name: " + $dataJson.windowsProductName
    Write-Host "BIOS Release Date: " + $dataJson.biosReleaseDate
    Write-Host "Domain: " + $dataJson.csDomain
    Write-Host "Model: " + $dataJson.csModel
    Write-Host "BitLocker Status: " + $dataJson.bitlockerStatus 
    Write-Host "Administrators: " + $dataJson.adminsString
    Write-Host "PC Name: " + $dataJson.pcName
    Write-Host "BIOS Serial Number: " + $dataJson.biosSerialNr
    Write-Host "RDP Status: " + $dataJson.rdpStatus
    Write-Host "Total RAM: " + $dataJson.RAM
    Write-Host "Antivirus: " + $dataJson.antivirus
    Write-Host "DisplayVersion: " + $dataJson.displayversion
}



$pcName = $env:COMPUTERNAME
$adminsString = Get-Administrators
$bitlockerStatusText = Get-BitLockerStatus
$rdpStatusText = Get-RDPStatus
$computerDetails = Get-ComputerInfoDetails
$antivirusString = Get-AntivirusInfo
$displayVersion = Get-WindowsDisplayVersion
$officeversionstring = Get-Officeversion
$RAM = Get-RAM
Write-Host "$RAM"


Send-WebhookNotification -webhookUrl $webhookUrl -pcName $pcName -adminsString $adminsString -bitlockerStatus $bitlockerStatusText -rdpStatus $rdpStatusText -computerDetails $computerDetails -antivirusString $antivirusString -RAM $RAM -displayVersion $displayVersion -officeversionstring $officeversionstring

$dataJson = @{
    windowsVersion = $computerDetails.windowsVersion
    windowsProductName = $computerDetails.windowsProductName
    biosReleaseDate = $computerDetails.biosReleaseDate
    csDomain = $computerDetails.csDomain
    csModel = $computerDetails.csModel
    bitlockerStatus = $bitlockerStatusText
    adminsString = $adminsString
    pcName = $pcName
    biosSerialNr = $computerDetails.biosSerialNr
    rdpStatus = $rdpStatusText
    RAM = $RAM
    antivirus = $antivirusString
    displayversion = $displayVersion
    officeversionstring = $displayofficeversionstring
}

Write-HostSystemInfo -dataJson $dataJson
