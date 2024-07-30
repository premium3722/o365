# Array mit den Gruppennamen
$kuerzel = "Firmenname"
$groupNames = @("SP_Data-HUB_${kuerzel}_Ordner01-RW", "SP_Data-HUB_${kuerzel}_Ordner01-R")

# Funktion zum Überprüfen und Installieren des AzureAD-Moduls
function Install-AzureADModule {
    if (-not (Get-Module -ListAvailable -Name AzureAD)) {
        Write-Host "AzureAD Modul nicht gefunden. Installiere AzureAD Modul..." -ForegroundColor Yellow
        Install-Module -Name AzureAD -Force -Scope CurrentUser
        Write-Host " "
    } else {
        Write-Host "AzureAD Modul ist bereits installiert." -ForegroundColor Green
        Write-Host " "
    }
}

# Verbinden mit Azure AD
Connect-AzureAD

# Schleife durch die Array-Elemente und erstellen der Gruppen
foreach ($groupName in $groupNames) {
    $existingGroup = Get-AzureADGroup | Where-Object { $_.DisplayName -eq $groupName }
    if ($existingGroup) {
        # Ausgabe in Orange für bereits vorhandene Gruppe
        Write-Host "Fehler: Gruppe '$groupName' existiert bereits." -ForegroundColor Yellow
    } else {
        try {
            New-AzureADGroup -DisplayName $groupName -MailEnabled $false -SecurityEnabled $true -MailNickname $groupName
            Write-Host "Gruppe '$groupName' erfolgreich erstellt." -ForegroundColor Green
        } catch {
            Write-Host "Fehler beim Erstellen der Gruppe '$groupName': $_" -ForegroundColor Red
        }
    }
}

# Überprüfen der Erstellung
foreach ($groupName in $groupNames) {
    $group = Get-AzureADGroup | Where-Object { $_.DisplayName -eq $groupName }
    if ($group) {
        Write-Host "Überprüfung: Gruppe '$groupName' OK." -ForegroundColor Green
    } else {
        Write-Host "Überprüfung: Fehler bei Gruppe '$groupName'." -ForegroundColor Red
    }
}
