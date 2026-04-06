#Requires -Module Microsoft.Graph.Authentication,Microsoft.Graph.Beta.DeviceManagement.Actions,Microsoft.Graph.Beta.Devices.CorporateManagement,Microsoft.Graph.Groups,Microsoft.Graph.Beta.DeviceManagement

function Invoke-TeamsWebhook {
    <#
    .SYNOPSIS
    Sends a Microsoft Teams adaptive card for a newly deployed EAM app version.

    .DESCRIPTION
    Builds an adaptive card that summarizes the published application, the old and new
    version numbers, and optionally the migrated assignments and app icon.

    .PARAMETER TeamsWebhookUri
    The Teams or Power Automate webhook endpoint that receives the adaptive card payload.

    .PARAMETER DeployedAppDisplayName
    The display name of the application that was published.

    .PARAMETER DeployedAppNewVersion
    The version that was just deployed.

    .PARAMETER DeployedAppPreviousVersion
    The version that was previously deployed.

    .PARAMETER ResizedBase64String
    Optional Base64-encoded PNG used in the card header.

    .PARAMETER AssignmentInfo
    Optional assignment summary objects with AssignmentMode, Group, FilterMode, and FilterName.

    .PARAMETER EspUpdateInfo
    Optional ESP update summary objects with EnrollmentStatusPage and Action.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TeamsWebhookUri,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DeployedAppDisplayName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DeployedAppNewVersion,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DeployedAppPreviousVersion,

        [Parameter(Mandatory = $false)]
        [string]
        $ResizedBase64String,

        [Parameter(Mandatory = $false)]
        [PSObject[]]
        $AssignmentInfo,

        [Parameter(Mandatory = $false)]
        [PSObject[]]
        $EspUpdateInfo
    )

    process {
        $columns = @()

        if ($ResizedBase64String) {
            $columns += @{
                type  = 'Column'
                width = 'auto'
                items = @(
                    @{
                        type = 'Image'
                        url  = "data:image/png;base64,$ResizedBase64String"
                        size = 'Medium'
                    }
                )
            }
        }

        $columns += @{
            type  = 'Column'
            width = 'stretch'
            items = @(
                @{
                    type   = 'TextBlock'
                    weight = 'Bolder'
                    text   = "New version of $DeployedAppDisplayName released"
                    wrap   = $true
                },
                @{
                    type     = 'TextBlock'
                    spacing  = 'None'
                    text     = "A new version of application $DeployedAppDisplayName has been deployed using the Enterprise app catalog. Devices will start to receive version $DeployedAppNewVersion."
                    isSubtle = $true
                    wrap     = $true
                }
            )
        }

        $cardBody = @(
            @{
                type    = 'ColumnSet'
                columns = $columns
            },
            @{
                type  = 'FactSet'
                facts = @(
                    @{
                        title = 'Previous deployed version'
                        value = $DeployedAppPreviousVersion
                    },
                    @{
                        title = 'Latest version starting to deploy'
                        value = $DeployedAppNewVersion
                    }
                )
            }
        )

        if ($AssignmentInfo) {
            $groupedAssignments = $AssignmentInfo | Group-Object -Property AssignmentMode

            foreach ($group in $groupedAssignments) {
                $cardBody += @{
                    type   = 'TextBlock'
                    weight = 'Bolder'
                    text   = "$($group.Name) assignments"
                    wrap   = $true
                }

                $assignmentRows = @(
                    @{
                        type  = 'TableRow'
                        cells = @(
                            @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = 'Group' }) },
                            @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = 'Filter mode' }) },
                            @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = 'Filter name' }) }
                        )
                    }
                )

                foreach ($assignment in $group.Group) {
                    $assignmentRows += @{
                        type  = 'TableRow'
                        cells = @(
                            @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = "$($assignment.Group)"; wrap = $true }) },
                            @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = "$($assignment.FilterMode)"; wrap = $true }) },
                            @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = "$($assignment.FilterName)"; wrap = $true }) }
                        )
                    }
                }

                $cardBody += @{
                    type    = 'Table'
                    columns = @(
                        @{ width = 1 },
                        @{ width = 1 },
                        @{ width = 1 }
                    )
                    rows    = $assignmentRows
                }
            }
        }

        if ($EspUpdateInfo) {
            $espRows = @(
                @{
                    type  = 'TableRow'
                    cells = @(
                        @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = 'ESP' }) },
                        @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = 'Action' }) }
                    )
                }
            )

            foreach ($espEntry in $EspUpdateInfo) {
                $espRows += @{
                    type  = 'TableRow'
                    cells = @(
                        @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = "$($espEntry.EnrollmentStatusPage)"; wrap = $true }) },
                        @{ type = 'TableCell'; items = @(@{ type = 'TextBlock'; text = "$($espEntry.Action)"; wrap = $true }) }
                    )
                }
            }

            $cardBody += @(
                @{
                    type   = 'TextBlock'
                    weight = 'Bolder'
                    text   = 'Enrollment Status Page updates'
                    wrap   = $true
                },
                @{
                    type    = 'Table'
                    columns = @(
                        @{ width = 2 },
                        @{ width = 3 }
                    )
                    rows    = $espRows
                }
            )
        }

        $card = @{
            type        = 'message'
            attachments = @(
                @{
                    contentType = 'application/vnd.microsoft.card.adaptive'
                    contentUrl  = $null
                    content     = @{
                        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
                        type      = 'AdaptiveCard'
                        body      = $cardBody
                    }
                }
            )
        } | ConvertTo-Json -Depth 20

        Write-Output "Sending Teams notification for $DeployedAppDisplayName..."

        try {
            Invoke-MgGraphRequest -Uri $TeamsWebhookUri -Method Post -Body $card -ContentType 'application/json'
        }
        catch {
            throw "Could not send the Teams webhook. Error: $_"
        }
    }
}

function Get-EAMAppsWithAvailableUpdates {
    <#
    .SYNOPSIS
    Retrieves Enterprise Application Management catalog apps with updates available.

    .DESCRIPTION
    Calls the Intune reporting endpoint used by the Enterprise App Management portal,
    stores the report output in a temporary file, and converts the returned row data
    into readable PowerShell objects.
    #>
    [CmdletBinding()]
    param ()

    function Get-ReportColumnIndex {
        param (
            [Parameter(Mandatory = $true)]
            [object[]]
            $Schema,

            [Parameter(Mandatory = $true)]
            [string]
            $ColumnName
        )

        for ($index = 0; $index -lt $Schema.Count; $index++) {
            $schemaEntry = $Schema[$index]

            if ($schemaEntry -is [string]) {
                if ($schemaEntry -eq $ColumnName) {
                    return $index
                }

                continue
            }

            $possibleNames = @(
                $schemaEntry.column,
                $schemaEntry.columnName,
                $schemaEntry.name,
                $schemaEntry.propertyName,
                $schemaEntry.displayName,
                $schemaEntry.localizedName
            ) | Where-Object { $_ }

            if ($possibleNames -contains $ColumnName) {
                return $index
            }
        }

        return -1
    }

    $requestBody = @{
        select  = @(
            'CurrentRevisionId',
            'LatestRevisionId',
            'CurrentAppVersion',
            'LatestAvailableVersion',
            'Publisher',
            'ApplicationName',
            'ApplicationId'
        )
        skip    = 0
        top     = 50
        orderBy = @()
        filter  = "UpdateAvailable eq 'true' and IsSuperseded eq 'false'"
    }

    $jsonBody = $requestBody | ConvertTo-Json -Depth 3
    $reportPath = Join-Path -Path $env:TEMP -ChildPath 'output.json'

    Write-Host 'Retrieving EAM catalog update report...'

    try {
        Invoke-MgGraphRequest -Uri '/beta/deviceManagement/reports/retrieveWin32CatalogAppsUpdateReport' -ContentType 'application/json' -Method Post -Body $jsonBody -OutputFilePath $reportPath | Out-Null
    }
    catch {
        throw "Failed to retrieve EAM apps with updates. Error: $_"
    }

    $jsonData = Get-Content -Path $reportPath -Raw | ConvertFrom-Json
    $report = [System.Collections.Generic.List[object]]::new()

    $schema = @($jsonData.schema)
    $columnIndexes = @{
        ApplicationId          = 0
        ApplicationName        = 1
        CurrentAppVersion      = 2
        CurrentRevisionId      = 3
        LatestAvailableVersion = 4
        LatestRevisionId       = 5
        Publisher              = 6
    }

    if ($schema.Count -gt 0) {
        foreach ($columnName in @($columnIndexes.Keys)) {
            $resolvedIndex = Get-ReportColumnIndex -Schema $schema -ColumnName $columnName

            if ($resolvedIndex -ge 0) {
                $columnIndexes[$columnName] = $resolvedIndex
            }
        }
    }

    foreach ($value in $jsonData.values) {
        $row = [PSCustomObject][ordered]@{
            ApplicationId          = $value[$columnIndexes.ApplicationId]
            ApplicationName        = $value[$columnIndexes.ApplicationName]
            CurrentAppVersion      = $value[$columnIndexes.CurrentAppVersion]
            CurrentRevisionId      = $value[$columnIndexes.CurrentRevisionId]
            LatestAvailableVersion = $value[$columnIndexes.LatestAvailableVersion]
            LatestRevisionId       = $value[$columnIndexes.LatestRevisionId]
            Publisher              = $value[$columnIndexes.Publisher]
        }

        if ([string]::IsNullOrWhiteSpace([string]$row.ApplicationId) -or
            [string]::IsNullOrWhiteSpace([string]$row.ApplicationName) -or
            [string]::IsNullOrWhiteSpace([string]$row.LatestRevisionId)) {
            Write-Warning ("Skipping malformed EAM report row: " + ($row | ConvertTo-Json -Compress))
            continue
        }

        $report.Add($row)
    }

    if ($report.Count -eq 0) {
        Write-Host 'No apps with available updates found.'
        return @()
    }

    return $report
}

function Get-ResizedBase64Icon {
    <#
    .SYNOPSIS
    Resizes a Base64-encoded image so it can be embedded in a Teams adaptive card.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Base64String,

        [Parameter(Mandatory = $false)]
        [int]
        $TargetWidth = 50
    )

    Add-Type -AssemblyName System.Drawing

    $imageBytes = [Convert]::FromBase64String($Base64String)
    $memoryStream = [System.IO.MemoryStream]::new([byte[]]$imageBytes)
    $image = [System.Drawing.Image]::FromStream($memoryStream)

    try {
        $aspectRatio = $image.Width / $image.Height
        $newWidth = $TargetWidth
        $newHeight = [math]::Round($TargetWidth / $aspectRatio)

        $resizedImage = [System.Drawing.Bitmap]::new($newWidth, $newHeight)
        $graphics = [System.Drawing.Graphics]::FromImage($resizedImage)

        try {
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage($image, 0, 0, $newWidth, $newHeight)

            $resizedStream = [System.IO.MemoryStream]::new()
            $resizedImage.Save($resizedStream, [System.Drawing.Imaging.ImageFormat]::Png)

            return [Convert]::ToBase64String($resizedStream.ToArray())
        }
        finally {
            $graphics.Dispose()
            $resizedImage.Dispose()
        }
    }
    finally {
        $image.Dispose()
        $memoryStream.Dispose()
    }
}

function Remove-OldSupersedenceChain {
    <#
    .SYNOPSIS
    Removes any supersedence chain behind the current app so only the new and previous versions remain.

    .DESCRIPTION
    Starting from the currently deployed app version, this function traverses any older
    superseded apps, removes the supersedence relationships from each source app, and
    deletes all older target apps. After it finishes, only the new app and the immediate
    previous app remain in Intune.

    .PARAMETER CurrentAppId
    The Intune app ID of the current version that the new deployment supersedes.

    .PARAMETER CurrentAppName
    The display name of the current app, used for logging.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CurrentAppId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CurrentAppName
    )

    Write-Output "Checking for older superseded versions of $CurrentAppName..."

    $pendingAppIds = [System.Collections.Generic.Queue[string]]::new()
    $visitedAppIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $appsToDelete = [System.Collections.Generic.List[PSCustomObject]]::new()
    $pendingAppIds.Enqueue($CurrentAppId)

    while ($pendingAppIds.Count -gt 0) {
        $sourceAppId = $pendingAppIds.Dequeue()

        if (-not $visitedAppIds.Add($sourceAppId)) {
            continue
        }

        try {
            $relationships = @(Get-MgBetaDeviceAppManagementMobileAppRelationship -MobileAppId $sourceAppId -All)
        }
        catch {
            Write-Warning "Could not retrieve relationships for app ID $sourceAppId. Error: $_"
            continue
        }

        # Filter to only supersedence relationships where this app is the superseding app (child direction).
        # targetType 'child' means this app supersedes the target (older app).
        # targetType 'parent' means the target supersedes this app (newer app) — skip those.
        $supersedenceRelationships = @($relationships | Where-Object {
            $odataType = if ($_.AdditionalProperties -and $_.AdditionalProperties.'@odata.type') {
                $_.AdditionalProperties.'@odata.type'
            } else {
                $null
            }

            $targetType = if ($_.TargetType) {
                [string]$_.TargetType
            } elseif ($_.AdditionalProperties -and $_.AdditionalProperties.targetType) {
                [string]$_.AdditionalProperties.targetType
            } else {
                $null
            }

            $odataType -eq '#microsoft.graph.mobileAppSupersedence' -and $targetType -eq 'child'
        })

        if ($supersedenceRelationships.Count -eq 0) {
            continue
        }

        foreach ($relationship in $supersedenceRelationships) {
            $targetId = if ($relationship.TargetId) { $relationship.TargetId } elseif ($relationship.AdditionalProperties.targetId) { $relationship.AdditionalProperties.targetId } else { $null }
            $targetName = if ($relationship.TargetDisplayName) { $relationship.TargetDisplayName } elseif ($relationship.AdditionalProperties.targetDisplayName) { $relationship.AdditionalProperties.targetDisplayName } else { 'Unknown' }
            $targetVersion = if ($relationship.TargetDisplayVersion) { $relationship.TargetDisplayVersion } elseif ($relationship.AdditionalProperties.targetDisplayVersion) { $relationship.AdditionalProperties.targetDisplayVersion } else { 'Unknown' }

            if (-not $targetId -or $visitedAppIds.Contains($targetId)) {
                continue
            }

            $pendingAppIds.Enqueue([string]$targetId)
            $appsToDelete.Add([PSCustomObject]@{
                SourceAppId   = $sourceAppId
                TargetId      = [string]$targetId
                TargetName    = $targetName
                TargetVersion = $targetVersion
            })
        }

        # Clear all supersedence relationships from this source app.
        Write-Output "Clearing supersedence links from app $sourceAppId..."

        try {
            $body = @{ relationships = @() } | ConvertTo-Json -Depth 5
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$sourceAppId/updateRelationships" -Method Post -Body $body -ContentType 'application/json' | Out-Null
        }
        catch {
            throw "Could not remove supersedence relationships from app ID $sourceAppId. Error: $_"
        }
    }

    if ($appsToDelete.Count -eq 0) {
        Write-Output "No older superseded versions of $CurrentAppName found."
        return
    }

    $uniqueAppsToDelete = $appsToDelete | Group-Object -Property TargetId | ForEach-Object { $_.Group | Select-Object -First 1 }

    foreach ($appToDelete in $uniqueAppsToDelete) {
        Write-Output "Deleting older app $($appToDelete.TargetName) version $($appToDelete.TargetVersion)..."

        try {
            Remove-MgBetaDeviceAppManagementMobileApp -MobileAppId $appToDelete.TargetId
        }
        catch {
            throw "Could not delete old app $($appToDelete.TargetName) version $($appToDelete.TargetVersion) with app ID $($appToDelete.TargetId). Error: $_"
        }
    }
}

function Update-EspTrackedApps {
    <#
    .SYNOPSIS
    Replaces the current app with the new app in any Enrollment Status Pages that track it.

    .DESCRIPTION
    Retrieves all Windows Enrollment Status Page configurations, checks whether the
    current app is present in selectedMobileAppIds, and when found replaces it with the
    newly deployed app. The function returns a summary that can be included in the Teams card.

    .PARAMETER CurrentAppId
    The Intune app ID of the currently deployed app.

    .PARAMETER DeployedAppId
    The Intune app ID of the newly deployed app.

    .PARAMETER CurrentAppName
    The app display name used for logging.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CurrentAppId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DeployedAppId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CurrentAppName
    )

    Write-Host "Checking ESP configurations for $CurrentAppName..."

    try {
        $espConfigurations = [System.Collections.Generic.List[object]]::new()
        $requestUri = 'https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations'

        do {
            $response = Invoke-MgGraphRequest -Uri $requestUri -Method Get -OutputType PSObject

            foreach ($configuration in @($response.value)) {
                if ($configuration.'@odata.type' -eq '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration') {
                    $espConfigurations.Add($configuration)
                }
            }

            $requestUri = $response.'@odata.nextLink'
        }
        while ($requestUri)
    }
    catch {
        throw "Could not retrieve Enrollment Status Page configurations. Error: $_"
    }

    $espUpdateInfo = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($espConfiguration in $espConfigurations) {
        $selectedMobileAppIds = @($espConfiguration.selectedMobileAppIds)

        if ($selectedMobileAppIds -notcontains $CurrentAppId) {
            continue
        }

        $updatedMobileAppIds = @(
            $selectedMobileAppIds |
                Where-Object { $_ -and $_ -ne $CurrentAppId }
        )

        if ($updatedMobileAppIds -notcontains $DeployedAppId) {
            $updatedMobileAppIds += $DeployedAppId
        }

        $updateBody = @{
            '@odata.type'        = '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration'
            selectedMobileAppIds = @($updatedMobileAppIds)
        } | ConvertTo-Json -Depth 5

        Write-Host "Updating ESP '$($espConfiguration.displayName)' with new app version..."

        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations/$($espConfiguration.id)" -Method Patch -Body $updateBody -ContentType 'application/json' | Out-Null
        }
        catch {
            throw "Could not update Enrollment Status Page $($espConfiguration.displayName). Error: $_"
        }

        $espUpdateInfo.Add([PSCustomObject]@{
                EnrollmentStatusPage = $espConfiguration.displayName
                Action               = 'Replaced previous app with the new version'
            })
    }

    if ($espUpdateInfo.Count -eq 0) {
        Write-Host "No ESP references found for $CurrentAppName."
    }

    return $espUpdateInfo
}

function Get-MobileAppCategories {
    <#
    .SYNOPSIS
    Retrieves the categories assigned to an Intune mobile app.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $MobileAppId
    )

    try {
        $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$MobileAppId/categories" -Method Get -OutputType PSObject
        return @($response.value)
    }
    catch {
        throw "Could not retrieve categories for app ID $MobileAppId. Error: $_"
    }
}

function Copy-MobileAppMetadata {
    <#
    .SYNOPSIS
    Copies supported metadata from the old app version to the newly deployed app.

    .DESCRIPTION
    Copies scope tags, company portal featured state, owner, and notes. It also recreates
    the old app's category associations on the new app.

    .PARAMETER SourceApp
    The existing Intune mobile app object whose metadata should be copied.

    .PARAMETER TargetAppId
    The Intune app ID of the newly deployed app.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [psobject]
        $SourceApp,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $TargetAppId
    )

    $sourceCategories = @(Get-MobileAppCategories -MobileAppId $SourceApp.Id)

    $roleScopeTagIds = @($SourceApp.RoleScopeTagIds) |
        Where-Object { $_ -ne $null -and $_ -ne '' } |
        ForEach-Object { "$_" } |
        Select-Object -Unique

    $metadataParams = [ordered]@{
        '@odata.type'   = '#microsoft.graph.win32CatalogApp'
        roleScopeTagIds = @($roleScopeTagIds)
        isFeatured      = [bool]$SourceApp.IsFeatured
        owner           = [string]$SourceApp.Owner
        notes           = [string]$SourceApp.Notes
    }

    Write-Output 'Migrating metadata (scope tags, featured state, owner, notes) to the new app...'

    try {
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$TargetAppId" -Method Patch -Body ($metadataParams | ConvertTo-Json -Depth 10) -ContentType 'application/json' | Out-Null
    }
    catch {
        throw "Failed to update scope tags and metadata for app ID $TargetAppId. Error: $_"
    }

    if ($sourceCategories.Count -eq 0) {
        Write-Output 'No categories to migrate.'
        return
    }

    foreach ($category in $sourceCategories) {
        $categoryId = $category.id

        if ([string]::IsNullOrWhiteSpace([string]$categoryId)) {
            continue
        }

        $refBody = @{
            '@odata.id' = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppCategories/$categoryId"
        } | ConvertTo-Json

        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$TargetAppId/categories/`$ref" -Method Post -Body $refBody -ContentType 'application/json' | Out-Null
        }
        catch {
            throw "Failed to assign category $categoryId to app ID $TargetAppId. Error: $_"
        }
    }

    Write-Output "Migrated $($sourceCategories.Count) category link(s) to the new app."
}

function Invoke-EAMAutoupdate {
    <#
    .SYNOPSIS
    Publishes newer EAM catalog app versions in Intune and migrates the old configuration.

    .DESCRIPTION
    For each Enterprise Application Management app with an available update, the script:
    - creates a new Intune app from the latest catalog revision
    - links the new app to the current app through supersedence
    - removes any supersedence chain older than the current app so only N and N-1 remain
    - migrates existing assignments from the current app to the new app
    - optionally updates Enrollment Status Pages that track the current app
    - copies scope tags and the app icon
    - optionally sends a Teams notification for the deployment

    .PARAMETER TeamsWebhookUri
    Optional Teams or Power Automate webhook URL used for deployment notifications.

    .PARAMETER UpdateESP
    When specified, replaces the current app with the newly deployed app in any
    Enrollment Status Page that tracks the current version.

    .PARAMETER ExcludeApps
    An array of application display names to skip during processing.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $TeamsWebhookUri,

        [Parameter(Mandatory = $false)]
        [switch]
        $UpdateESP,

        [Parameter(Mandatory = $false)]
        [string[]]
        $ExcludeApps = @()
    )

    begin {
        # Intune uses fixed synthetic group identifiers for the built-in All Users and All Devices targets.
        $intuneAllUsersId = 'acacacac-9df4-4c7d-9d50-4ef0226f57a9'
        $intuneAllDevicesId = 'adadadad-808e-44e2-905a-0b7873a8a531'
    }

    process {
        Write-Output 'Retrieving Intune assignment filters...'

        try {
            $allIntuneFilters = Get-MgBetaDeviceManagementAssignmentFilter -All -Property Id, DisplayName | Group-Object -Property Id -AsHashTable -AsString
        }
        catch {
            throw 'Could not retrieve all Intune filters.'
        }

        Write-Output "Found $($allIntuneFilters.Count) assignment filters."

        $appsWithUpdate = Get-EAMAppsWithAvailableUpdates
        Write-Output "Found $($appsWithUpdate.Count) app(s) with available updates."

        foreach ($app in $appsWithUpdate) {
            if ($ExcludeApps -and $ExcludeApps -contains $app.ApplicationName) {
                Write-Output "Skipping excluded app $($app.ApplicationName)."
                continue
            }

            [System.Collections.Generic.List[PSCustomObject]]$appAssignmentsObject = [System.Collections.Generic.List[PSCustomObject]]::new()
            [System.Collections.Generic.List[PSCustomObject]]$espUpdateInfo = [System.Collections.Generic.List[PSCustomObject]]::new()

            $currentAppVersion = $app.CurrentAppVersion
            $latestAvailableVersion = $app.LatestAvailableVersion
            $latestPackageId = $app.LatestRevisionId

            Write-Output "Processing $($app.ApplicationName) ($currentAppVersion -> $latestAvailableVersion)..."

            if ([string]::IsNullOrWhiteSpace([string]$latestPackageId)) {
                Write-Warning "Skipping app $($app.ApplicationName) because LatestRevisionId is empty."
                continue
            }

            # Convert the latest catalog package into a new mobile app payload.
            $catalogConversionUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/convertFromMobileAppCatalogPackage(mobileAppCatalogPackageId='$latestPackageId')"
            $mobileAppPayload = (Invoke-MgGraphRequest -Uri $catalogConversionUri -Method Get -OutputType PSObject) |
                Select-Object * -ExcludeProperty '@odata.context', id, largeIcon, createdDateTime, lastModifiedDateTime, owner, notes, size, minimumSupportedOperatingSystem, minimumFreeDiskSpaceInMB, minimumMemoryInMB, minimumNumberOfProcessors, minimumCpuSpeedInMHz

            if ($null -eq $mobileAppPayload) {
                throw "The catalog conversion returned an empty response for app $($app.ApplicationName) with LatestRevisionId $latestPackageId."
            }

            $appPayloadJson = $mobileAppPayload | ConvertTo-Json -Depth 20

            if ([string]::IsNullOrWhiteSpace($appPayloadJson)) {
                throw "The catalog conversion produced an empty JSON payload for app $($app.ApplicationName) with LatestRevisionId $latestPackageId."
            }

            Write-Output "Deploying $($app.ApplicationName) version $latestAvailableVersion from the catalog..."

            try {
                $deployedApp = Invoke-MgGraphRequest -Method Post -Uri 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps' -Body $appPayloadJson -ContentType 'application/json'
            }
            catch {
                throw "Could not deploy the latest version $latestAvailableVersion of the app $($app.ApplicationName). Error: $_"
            }

            Write-Output "Deployed $($app.ApplicationName) version $latestAvailableVersion successfully."

            # Link the newly created app to the currently deployed version so Intune treats it as an update.
            $relationships = @{
                relationships = @(
                    @{
                        '@odata.type'    = '#microsoft.graph.mobileAppSupersedence'
                        targetId         = "$($app.ApplicationId)"
                        supersedenceType = 'update'
                    }
                )
            }

            Write-Output "Configuring supersedence: $latestAvailableVersion supersedes $currentAppVersion..."

            try {
                Invoke-MgGraphRequest -Method Post -Uri "/beta/deviceAppManagement/mobileApps/$($deployedApp.id)/updateRelationships" -Body ($relationships | ConvertTo-Json -Depth 10) -ContentType 'application/json' | Out-Null
            }
            catch {
                throw "Could not initiate supersedence between version $currentAppVersion and $latestAvailableVersion of the app $($app.ApplicationName). Error: $_"
            }

            # Read the current app assignments so they can be re-created on the new app.
            $previousVersionAppAssignments = Get-MgBetaDeviceAppManagementMobileAppAssignment -MobileAppId $app.ApplicationId | Sort-Object Settings
            $previousVersionAppAssignments = $previousVersionAppAssignments | Sort-Object -Descending { $_.Settings.AdditionalProperties.Count }

            if ($null -eq $previousVersionAppAssignments) {
                Write-Output "No assignments to migrate for $($app.ApplicationName)."
            }
            else {
                Write-Output "Migrating $($previousVersionAppAssignments.Count) assignment(s) for $($app.ApplicationName)..."

                foreach ($previousVersionAppAssignment in $previousVersionAppAssignments) {
                    # Assignment identifiers are emitted in the format '<groupId>_0_0'.
                    $assignmentGroupId = ($previousVersionAppAssignment.Id -split '_')[0]
                    $isExcludeAssignment = $previousVersionAppAssignment.Target.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.exclusionGroupAssignmentTarget'

                    if (-not $isExcludeAssignment) {
                        if ($previousVersionAppAssignment.Target.DeviceAndAppManagementAssignmentFilterType -eq 'none') {
                            $previousFilterType = 'none'
                            $previousFilterId = $null
                        }
                        else {
                            $previousFilterType = $previousVersionAppAssignment.Target.DeviceAndAppManagementAssignmentFilterType
                            $previousFilterId = $previousVersionAppAssignment.Target.DeviceAndAppManagementAssignmentFilterId
                        }

                        if ($assignmentGroupId -eq $intuneAllUsersId) {
                            $assignmentGroup = @{ DisplayName = 'All Users'; Id = $assignmentGroupId }
                            $target = @{
                                '@odata.type'                                = '#microsoft.graph.allLicensedUsersAssignmentTarget'
                                deviceAndAppManagementAssignmentFilterId   = "$previousFilterId"
                                deviceAndAppManagementAssignmentFilterType = "$previousFilterType"
                            }
                        }
                        elseif ($assignmentGroupId -eq $intuneAllDevicesId) {
                            $assignmentGroup = @{ DisplayName = 'All Devices'; Id = $assignmentGroupId }
                            $target = @{
                                '@odata.type'                                = '#microsoft.graph.allDevicesAssignmentTarget'
                                deviceAndAppManagementAssignmentFilterId   = "$previousFilterId"
                                deviceAndAppManagementAssignmentFilterType = "$previousFilterType"
                            }
                        }
                        else {
                            $assignmentGroup = Get-MgGroup -GroupId $assignmentGroupId
                            $target = @{
                                '@odata.type'                                = '#microsoft.graph.groupAssignmentTarget'
                                groupId                                    = $assignmentGroup.Id
                                deviceAndAppManagementAssignmentFilterId   = "$previousFilterId"
                                deviceAndAppManagementAssignmentFilterType = "$previousFilterType"
                            }
                        }

                        if ($target.deviceAndAppManagementAssignmentFilterType -eq 'none') {
                            $target.Remove('deviceAndAppManagementAssignmentFilterId')
                            $target.Remove('deviceAndAppManagementAssignmentFilterType')
                        }

                        $assignmentIntent = [string]$previousVersionAppAssignment.Intent
                        $assignmentMode = switch ($assignmentIntent) {
                            'available' { 'Available' }
                            'required' { 'Required' }
                            default { 'Uninstall' }
                        }

                        $filterName = if ($previousFilterType -eq 'none') {
                            'Null'
                        }
                        elseif ($allIntuneFilters.ContainsKey($previousFilterId)) {
                            $allIntuneFilters[$previousFilterId].DisplayName
                        }
                        else {
                            $previousFilterId
                        }

                        $appAssignmentsObject.Add([PSCustomObject]@{
                                AssignmentMode = $assignmentMode
                                Group          = $assignmentGroup.DisplayName
                                FilterMode     = if ($previousFilterType) { $previousFilterType } else { 'none' }
                                FilterName     = $filterName
                            })

                        $settings = @{
                            '@odata.type'                 = '#microsoft.graph.win32CatalogAppAssignmentSettings'
                            installTimeSettings          = $null
                            deliveryOptimizationPriority = "$($previousVersionAppAssignment.Settings.AdditionalProperties.deliveryOptimizationPriority)"
                            notifications                = "$($previousVersionAppAssignment.Settings.AdditionalProperties.notifications)"
                            restartSettings              = $null
                        }

                        if ($assignmentIntent -eq 'available') {
                            $settings.autoUpdateSettings = @{
                                autoUpdateSupersededAppsState = 'enabled'
                                '@odata.type'                 = '#microsoft.graph.win32LobAppAutoUpdateSettings'
                            }
                        }

                        $params = @{
                            '@odata.type' = '#microsoft.graph.mobileAppAssignment'
                            intent        = $assignmentIntent
                            target        = $target
                            settings      = $settings
                        }

                        try {
                            New-MgBetaDeviceAppManagementMobileAppAssignment -MobileAppId $deployedApp.Id -BodyParameter $params | Out-Null
                        }
                        catch {
                            throw "Failed to create the app assignment. Error: $_"
                        }

                        Write-Output "  Migrated assignment: $assignmentMode -> $($assignmentGroup.DisplayName)"
                    }
                    else {
                        $assignmentGroup = Get-MgGroup -GroupId $assignmentGroupId
                        $target = @{
                            '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'
                            groupId       = $assignmentGroup.Id
                        }
                        $params = @{
                            '@odata.type' = '#microsoft.graph.mobileAppAssignment'
                            intent        = "$($previousVersionAppAssignment.Intent)"
                            target        = $target
                        }

                        try {
                            New-MgBetaDeviceAppManagementMobileAppAssignment -MobileAppId $deployedApp.Id -BodyParameter $params | Out-Null
                        }
                        catch {
                            throw "Failed to create the exclusion assignment. Error: $_"
                        }

                        Write-Output "  Migrated exclusion: $($assignmentGroup.DisplayName)"

                        $appAssignmentsObject.Add([PSCustomObject]@{
                                AssignmentMode = "Exclude: $($previousVersionAppAssignment.Intent)"
                                Group          = $assignmentGroup.DisplayName
                                FilterMode     = 'N/A'
                                FilterName     = 'N/A'
                            })
                    }
                }
            }

            # Keep the newly deployed app and the current app, but remove anything older than the current app.
            Remove-OldSupersedenceChain -CurrentAppId $app.ApplicationId -CurrentAppName $app.ApplicationName

            if ($UpdateESP) {
                foreach ($espResult in @(Update-EspTrackedApps -CurrentAppId $app.ApplicationId -DeployedAppId $deployedApp.Id -CurrentAppName $app.ApplicationName)) {
                    $espUpdateInfo.Add($espResult)
                }
            }

            try {
                $oldAppInfo = Get-MgBetaDeviceAppManagementMobileApp -MobileAppId $app.ApplicationId
            }
            catch {
                throw "Could not retrieve the old app's metadata. Error: $_"
            }

            Copy-MobileAppMetadata -SourceApp $oldAppInfo -TargetAppId $deployedApp.Id

            $hasPreviousIcon = $oldAppInfo.LargeIcon -and $oldAppInfo.LargeIcon.Value -and $oldAppInfo.LargeIcon.Value.Length -gt 0

            if (-not $hasPreviousIcon) {
                Write-Output 'No icon found on the previous version, skipping icon migration.'

                if ($TeamsWebhookUri) {
                    $notificationParams = @{
                        TeamsWebhookUri            = $TeamsWebhookUri
                        DeployedAppDisplayName     = $app.ApplicationName
                        DeployedAppNewVersion      = $deployedApp.DisplayVersion
                        DeployedAppPreviousVersion = $app.CurrentAppVersion
                    }

                    if ($appAssignmentsObject.Count -gt 0) {
                        $notificationParams.AssignmentInfo = $appAssignmentsObject
                    }

                    if ($espUpdateInfo.Count -gt 0) {
                        $notificationParams.EspUpdateInfo = $espUpdateInfo
                    }

                    Invoke-TeamsWebhook @notificationParams
                }
            }
            else {
                Write-Output 'Migrating app icon...'

                $byteArray = [byte[]]@($oldAppInfo.LargeIcon.Value)
                $base64String = [Convert]::ToBase64String($byteArray)

                $imageBodyUpdate = @{
                    '@odata.type' = '#microsoft.graph.win32CatalogApp'
                    largeIcon     = @{
                        '@odata.type' = '#microsoft.graph.mimeContent'
                        type          = 'String'
                        value         = $base64String
                    }
                } | ConvertTo-Json -Depth 10

                try {
                    Update-MgBetaDeviceAppManagementMobileApp -MobileAppId $deployedApp.Id -BodyParameter $imageBodyUpdate
                }
                catch {
                    throw "Failed to update the icon of the app $($deployedApp.DisplayName). Error: $_"
                }

                if ($TeamsWebhookUri) {
                    $notificationParams = @{
                        TeamsWebhookUri            = $TeamsWebhookUri
                        DeployedAppDisplayName     = $app.ApplicationName
                        DeployedAppNewVersion      = $deployedApp.DisplayVersion
                        DeployedAppPreviousVersion = $app.CurrentAppVersion
                        ResizedBase64String        = Get-ResizedBase64Icon -Base64String $base64String
                    }

                    if ($appAssignmentsObject.Count -gt 0) {
                        $notificationParams.AssignmentInfo = $appAssignmentsObject
                    }

                    if ($espUpdateInfo.Count -gt 0) {
                        $notificationParams.EspUpdateInfo = $espUpdateInfo
                    }

                    Invoke-TeamsWebhook @notificationParams
                }
            }
        }
    }
}

try {
    Connect-MgGraph -Identity
}
catch {
    throw "Failed to connect to Graph. Error: $_"
}

Invoke-EAMAutoupdate -TeamsWebhookUri <TeamsWebhookUri> -UpdateESP -ExcludeApps <"draw.io Desktop">