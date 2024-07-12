#1 Powerautomate erstellen mit Warten auf Webhook von Teams
#2 Karte in einem Chat oder Kanalveröffentlichen
#3 Karten Vorlagen: https://adaptivecards.io/designer

# Definiere die URL des Webhooks
$webhookUrl = "https://prod-201.westeurope.logic.azure.com:443/workflows/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
# Definieren Sie die Daten
$dataJson = @{
    title = "Publish Adaptive Card Schema"
    description = "Now that we have defined the main rules and features of the format, we need to produce a schema and publish it to GitHub. The schema will be the starting point of our reference documentation."
    creator = @{
        name = "Matt Hidinger"
        profileImage = "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg"
    }
    createdUtc = "2017-02-14T06:08:39Z"
    viewUrl = "https://adaptivecards.io"
    properties = @(
        @{ key = "Board"; value = "Adaptive Cards" }
        @{ key = "List"; value = "Backlog" }
        @{ key = "Assigned to"; value = "Matt Hidinger" }
        @{ key = "Due date"; value = "Not set" }
    )
}

# Erstellen Sie die Adaptive Card JSON
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
                        type = "ColumnSet"
                        columns = @(
                            @{
                                type = "Column"
                                width = "auto"
                                items = @(
                                    @{
                                        type = "Image"
                                        url = $dataJson.creator.profileImage
                                        size = "Small"
                                        style = "Person"
                                    }
                                )
                            }
                            @{
                                type = "Column"
                                width = "stretch"
                                items = @(
                                    @{
                                        type = "TextBlock"
                                        weight = "Bolder"
                                        text = $dataJson.creator.name
                                        wrap = $true
                                    }
                                    @{
                                        type = "TextBlock"
                                        spacing = "None"
                                        text = "Created: " + $dataJson.createdUtc
                                        isSubtle = $true
                                        wrap = $true
                                    }
                                )
                            }
                        )
                    }
                    @{
                        type = "FactSet"
                        facts = @(
                            foreach ($property in $dataJson.properties) {
                                @{
                                    title = $property.key
                                    value = $property.value
                                }
                            }
                        )
                    }
                )
                actions = @(
                    @{
                        type = "Action.OpenUrl"
                        title = "View"
                        url = $dataJson.viewUrl
                    }
                )
            }
        }
    )
}

# Konvertieren Sie die Adaptive Card in JSON-Format
$AdaptiveCardJson = $AdaptiveCard | ConvertTo-Json -Depth 10

# Senden Sie die Adaptive Card über den Webhook
Invoke-RestMethod -Method POST -Uri $webhookUrl -Body $AdaptiveCardJson -ContentType 'application/json'
