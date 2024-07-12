#1 Powerautomate erstellen mit Warten auf Webhook von Teams
#2 Karte in einem Chat oder Kanalveröffentlichen
#3 Karten Vorlagen: https://adaptivecards.io/designer

$webhookUrl = "https://prod-201.westeurope.logic.azure.com:443/workflows/XXXXXXXXXXXXXXXXXXXX"


$admins = net localgroup administratoren | Where-Object {$_ -AND $_ -notmatch "Der Befehl wurde erfolgreich ausgeführt"} | Select-Object -Skip 4
$pcName = $env:COMPUTERNAME
$list = $admins -join "`n"

if ($admins.Count -eq 1 -and $admins -contains 'localadmin') {

    Write-Host "'localadmin' ist der einzige Administrator. Kein Webhook wird gesendet."
} else {
    # Definieren  Daten
    $dataJson = @{
        title = "Administrator List"
        description = "Folgender PC hat noch andere Administratoren als der Localadmin."
        pcName = $pcName
        admins = $admins
        ticketUrl = "https://XXXX.XXX"
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
                            text = "PC Name: " + $dataJson.pcName
                            wrap = $true
                        }
                        @{
                            type = "TextBlock"
                            text = "Administrators:"
                            wrap = $true
                        }
                        @{
                            type = "FactSet"
                            facts = @(
                                foreach ($admin in $dataJson.admins) {
                                    @{
                                        title = ""
                                        value = $admin -replace '[^ -~]', ''  # Entfernt nicht druckbare Zeichen
                                    }
                                }
                            )
                        }
                    )
                    actions = @(
                        @{
                            type = "Action.OpenUrl"
                            title = "Ticket erstellen"
                            url = $dataJson.ticketUrl
                        }
                    )
                }
            }
        )
    }

    $AdaptiveCardJson = $AdaptiveCard | ConvertTo-Json -Depth 10
    $encodedAdaptiveCardJson = [System.Text.Encoding]::UTF8.GetBytes($AdaptiveCardJson)
    Invoke-RestMethod -Method POST -Uri $webhookUrl -Body $encodedAdaptiveCardJson -ContentType 'application/json'
}

