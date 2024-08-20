$webhookUrl = "https://prod-188.westeurope.logic.azure.com:443/workflows/"

# Überprüfen und ggf. Installation der Module
$modules = @("MSOnline", "ExchangeOnlineManagement", "AzureAD")
foreach ($module in $modules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        Install-Module -Name $module -Force -AllowClobber
    }
}

# Importieren der Module
Import-Module MSOnline
Import-Module ExchangeOnlineManagement
Import-Module AzureAD

# Verbindung zu Office 365 herstellen
try {
    Connect-MsolService -ErrorAction Stop
    Write-Host "Erfolgreich mit MSOnline verbunden."
} catch {
    Write-Host "Fehler bei der Verbindung zu MSOnline: $_"
    exit
}

# Verbindung zu Exchange Online herstellen
try {
    Connect-ExchangeOnline -ErrorAction Stop
    Write-Host "Erfolgreich mit Exchange Online verbunden."
} catch {
    Write-Host "Fehler bei der Verbindung zu Exchange Online: $_"
    exit
}

# Verbindung zu Azure AD herstellen
try {
    Connect-AzureAD -ErrorAction Stop
    Write-Host "Erfolgreich mit Azure AD verbunden."
} catch {
    Write-Host "Fehler bei der Verbindung zu Azure AD: $_"
    exit
}

# Den Tenant-Namen abrufen
$tenantName = (Get-MsolCompanyInformation).DisplayName -replace "[^\w\s]", ""
Write-Host "$tenantName"

# Alle lizenzierten Benutzerkonten abrufen
$licensedUsers = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true }

function Get-O365Users {
    $csvData = @() # Initialisiere das Array für die CSV-Daten
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    foreach ($mailbox in $mailboxes) {
        $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName
        $totalItemSize = $stats.TotalItemSize.ToString()  # Größe als String
        $totalItemSizeGB = $totalItemSize.Split(" ")[0]  # Extrahiere nur den GB-Wert
        $lastLogonTime = $stats.LastLogonTime

        try {
            $user = Get-AzureADUser -ObjectId $mailbox.UserPrincipalName -ErrorAction Stop
            $mfaStatus = Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName | Select-Object -ExpandProperty StrongAuthenticationMethods
            if ($mfaStatus.Count -eq 0) {
                $mfaStatusText = "Kein MFA"
            } else {
                $mfaStatusText = "MFA OK"
            }

            $assignedLicenses = (Get-MsolUser -UserPrincipalName $mailbox.UserPrincipalName).Licenses | ForEach-Object { $_.AccountSkuId }

            $assignedLicensesText = $assignedLicenses -join ", "

            $csvData += [PSCustomObject]@{
                Benutzername          = $mailbox.UserPrincipalName
                Anzeigename           = $mailbox.DisplayName
                GesamterSpeicherplatz = "$totalItemSizeGB GB"
                LetzteAnmeldung       = $lastLogonTime
                MFAStatus             = $mfaStatusText
                ZugeordneteLizenzen   = $assignedLicensesText
            }
        } catch {
            Write-Host "Benutzer nicht gefunden oder ein Fehler ist aufgetreten: $mailbox.UserPrincipalName"
        }
    }
    return $csvData
}

function Send-WebhookNotification {
    param (
        [string]$webhookUrl,
        [PSCustomObject]$user
    )

    $dataJson = @{
        title = "Benutzer List"
        description = "O365 User."
        Benutzername = $user.Benutzername
        Anzeigename = $user.Anzeigename
        GesamterSpeicherplatz = $user.GesamterSpeicherplatz
        LetzteAnmeldung = $user.LetzteAnmeldung
        MFAStatus = $user.MFAStatus
        ZugeordneteLizenzen = $user.ZugeordneteLizenzen  # Lizenzen hinzufügen
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
                            text = $dataJson.Benutzername
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.Anzeigename
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.GesamterSpeicherplatz
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.LetzteAnmeldung
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.MFAStatus
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = $dataJson.ZugeordneteLizenzen 
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
        [PSCustomObject]$user
    )
    Write-Host "Benutzername: " + $user.Benutzername
    Write-Host "Anzeigename: " + $user.Anzeigename
    Write-Host "GesamterSpeicherplatz: " + $user.GesamterSpeicherplatz
    Write-Host "LetzteAnmeldung: " + $user.LetzteAnmeldung
    Write-Host "MFA: " + $user.MFAStatus
    Write-Host "Lizenz: " + $user.ZugeordneteLizenzen  # Lizenzen in der Konsole ausgeben
}

$users = Get-O365Users
Write-Host "Starte Webhook"
foreach ($user in $users) {
    # Schreibe die Benutzerinfo in die Konsole
    Write-HostSystemInfo -user $user
    # Sende eine Webhook-Benachrichtigung
    Start-Sleep -Seconds 10
    Send-WebhookNotification -webhookUrl $webhookUrl -user $user
}
