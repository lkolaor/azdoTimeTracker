<#
.SYNOPSIS
    Azure DevOps REST API functions for the Time Tracker module.
.DESCRIPTION
    Functions to query work items, fetch details/comments,
    and update time fields via the Azure DevOps REST API.
#>

# ── Helper: Build auth header ──────────────────────────────────────
function Get-AzDoAuthHeader {
    param([string]$PAT)
    $base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
    return @{ Authorization = "Basic $base64" }
}

# ── Helper: Safely read a field that may not exist ─────────────────
function Get-SafeField {
    param($Fields, [string]$Name)
    $prop = $Fields.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

# ── Helper: Write debug log (only when -Debug is active) ──────────
function Write-TTDebugLog {
    param([string]$Message)
    if (-not $script:TTDebugEnabled) { return }
    $logPath = Join-Path (Get-TTConfigDir) "debug.log"
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] $Message"
    Add-Content -Path $logPath -Value $line -ErrorAction SilentlyContinue
}

# ── Get the current user's display name via Connection Data API ────
$script:CachedDisplayName = $null
function Get-CurrentUserDisplayName {
    param(
        [string]$Organization,
        [string]$PAT
    )

    if ($script:CachedDisplayName) { return $script:CachedDisplayName }

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/_apis/connectiondata?api-version=7.1-preview"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        if ($result.authenticatedUser -and $result.authenticatedUser.providerDisplayName) {
            $script:CachedDisplayName = $result.authenticatedUser.providerDisplayName
            Write-TTDebugLog "Get-CurrentUserDisplayName: $($script:CachedDisplayName)"
            return $script:CachedDisplayName
        }
    }
    catch {
        Write-TTDebugLog "Get-CurrentUserDisplayName error: $($_.Exception.Message)"
    }
    return $null
}

# ── Get my work items via WIQL ─────────────────────────────────────
function Get-MyWorkItems {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"

    # WIQL query: get all work items assigned to me
    $wiql = @{
        query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.AssignedTo] = @me
  AND [System.State] <> 'Closed'
  AND [System.State] <> 'Removed'
  AND [System.State] <> 'Done'
  AND [System.State] <> 'Released'
ORDER BY [System.WorkItemType], [System.Title]
"@
    } | ConvertTo-Json

    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiql -ErrorAction Stop
    }
    catch {
        Write-Host "Error querying Azure DevOps: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }

    if (-not $result.workItems -or $result.workItems.Count -eq 0) {
        return @()
    }

    # Batch-fetch work item details (API supports max 200 at a time)
    $allItems = @()
    $ids = $result.workItems | ForEach-Object { $_.id }
    $batchSize = 200

    for ($i = 0; $i -lt $ids.Count; $i += $batchSize) {
        $batchIds = $ids[$i .. [Math]::Min($i + $batchSize - 1, $ids.Count - 1)]
        $idsString = $batchIds -join ","

        $fields = @(
            "System.Id",
            "System.Title",
            "System.WorkItemType",
            "System.State",
            "System.AssignedTo",
            "System.Description",
            "System.Parent",
            "Microsoft.VSTS.Scheduling.OriginalEstimate",
            "Microsoft.VSTS.Scheduling.CompletedWork",
            "Microsoft.VSTS.Scheduling.RemainingWork"
        ) -join ","

        $detailUrl = "$baseUrl/wit/workitems?ids=$idsString&fields=$fields&api-version=7.1"

        try {
            $details = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers -ErrorAction Stop
            $allItems += $details.value
        }
        catch {
            Write-Host "Error fetching work item details: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Collect active item IDs for parent-match check below
    $myActiveIds = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($item in $allItems) {
        [void]$myActiveIds.Add([int]$item.fields.'System.Id')
    }

    # Fetch closed/done items also assigned to me; keep only those whose parent
    # is still in the active set so they remain visible under the parent.
    $wiqlClosed = @{
        query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.AssignedTo] = @me
  AND ([System.State] = 'Closed'
    OR [System.State] = 'Done'
    OR [System.State] = 'Released')
  AND [System.ChangedDate] >= @Today - 90
ORDER BY [System.ChangedDate] DESC
"@
    } | ConvertTo-Json

    try {
        $closedResult = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiqlClosed -ErrorAction Stop

        if ($closedResult.workItems -and $closedResult.workItems.Count -gt 0) {
            $closedIds = $closedResult.workItems | ForEach-Object { $_.id }
            # Exclude IDs already in $allItems
            $closedIds = $closedIds | Where-Object { -not $myActiveIds.Contains([int]$_) }

            if ($closedIds.Count -gt 0) {
                for ($i = 0; $i -lt $closedIds.Count; $i += $batchSize) {
                    $batchIds = $closedIds[$i .. [Math]::Min($i + $batchSize - 1, $closedIds.Count - 1)]
                    $idsString = $batchIds -join ","
                    $detailUrl = "$baseUrl/wit/workitems?ids=$idsString&fields=$fields&api-version=7.1"
                    try {
                        $details = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers -ErrorAction Stop
                        foreach ($wi in $details.value) {
                            $parentId = Get-SafeField -Fields $wi.fields -Name 'System.Parent'
                            if ($parentId -and $myActiveIds.Contains([int]$parentId)) {
                                $allItems += $wi
                            }
                        }
                    }
                    catch {
                        # Silently ignore; closed items are best-effort
                    }
                }
            }
        }
    }
    catch {
        # Silently ignore; closed-child fetch is best-effort
    }

    # Also fetch parent items that we don't own but need for hierarchy display
    $parentIds = @()
    foreach ($item in $allItems) {
        $parentId = Get-SafeField -Fields $item.fields -Name 'System.Parent'
        if ($parentId -and ($ids -notcontains $parentId)) {
            $parentIds += $parentId
        }
    }
    $parentIds = $parentIds | Select-Object -Unique

    $parentItems = @()
    if ($parentIds.Count -gt 0) {
        for ($i = 0; $i -lt $parentIds.Count; $i += $batchSize) {
            $batchIds = $parentIds[$i .. [Math]::Min($i + $batchSize - 1, $parentIds.Count - 1)]
            $idsString = $batchIds -join ","
            $detailUrl = "$baseUrl/wit/workitems?ids=$idsString&fields=$fields&api-version=7.1"
            try {
                $details = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers -ErrorAction Stop
                $parentItems += $details.value
            }
            catch {
                # Silently continue; parent not accessible
            }
        }
    }

    return @{
        MyItems    = $allItems
        ParentItems = $parentItems
    }
}

# ── Build hierarchical tree ────────────────────────────────────────
function Build-WorkItemTree {
    param(
        [array]$MyItems,
        [array]$ParentItems
    )

    # Combine all items into a lookup
    $allLookup = @{}
    foreach ($item in $MyItems) {
        $f = $item.fields
        $id = $f.'System.Id'
        $assignedTo = $null
        $assignedToField = Get-SafeField -Fields $f -Name 'System.AssignedTo'
        if ($assignedToField) { $assignedTo = $assignedToField.displayName }
        $allLookup[$id] = @{
            Id              = $id
            Title           = $f.'System.Title'
            Type            = $f.'System.WorkItemType'
            State           = $f.'System.State'
            AssignedTo      = $assignedTo
            ParentId        = Get-SafeField -Fields $f -Name 'System.Parent'
            OriginalEstimate = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.OriginalEstimate'
            CompletedWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
            RemainingWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
            IsMine          = $true
            IsSeparator     = $false
            IsRelated       = $false
            Depth           = 0
            Children        = @()
            Raw             = $item
        }
    }

    foreach ($item in $ParentItems) {
        $f = $item.fields
        $id = $f.'System.Id'
        if (-not $allLookup.ContainsKey($id)) {
            $assignedTo = $null
            $assignedToField = Get-SafeField -Fields $f -Name 'System.AssignedTo'
            if ($assignedToField) { $assignedTo = $assignedToField.displayName }
            $allLookup[$id] = @{
                Id              = $id
                Title           = $f.'System.Title'
                Type            = $f.'System.WorkItemType'
                State           = $f.'System.State'
                AssignedTo      = $assignedTo
                ParentId        = Get-SafeField -Fields $f -Name 'System.Parent'
                OriginalEstimate = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.OriginalEstimate'
                CompletedWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
                RemainingWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
                IsMine          = $false
                IsSeparator     = $false
                IsRelated       = $false
                Depth           = 0
                Children        = @()
                Raw             = $item
            }
        }
    }

    # Build parent-child relationships (deduplicate children)
    foreach ($key in @($allLookup.Keys)) {
        $node = $allLookup[$key]
        $parentId = $node.ParentId
        if ($parentId -and $allLookup.ContainsKey($parentId)) {
            $parent = $allLookup[$parentId]
            $existingChildren = @($parent['Children'])
            $alreadyChild = $false
            foreach ($c in $existingChildren) {
                if ([string]$c -eq [string]$key) { $alreadyChild = $true; break }
            }
            if (-not $alreadyChild) {
                $parent['Children'] = $existingChildren + @($key)
            }
        }
    }

    # Find root items (no parent in our set)
    $roots = @()
    foreach ($key in @($allLookup.Keys)) {
        $node = $allLookup[$key]
        $parentId = $node.ParentId
        if (-not $parentId -or -not $allLookup.ContainsKey($parentId)) {
            $roots += $key
        }
    }

    # Flatten tree with indentation – depth-first
    $flat = [System.Collections.ArrayList]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new()

    function Add-Subtree {
        param($NodeId, $Depth, $Lookup, [System.Collections.ArrayList]$List, [System.Collections.Generic.HashSet[string]]$Visited)
        $nodeKey = [string]$NodeId
        if ($Visited.Contains($nodeKey)) { return }
        $node = $Lookup[$NodeId]
        if ($null -eq $node) { return }
        [void]$Visited.Add($nodeKey)
        $node['Depth'] = $Depth
        [void]$List.Add($node)
        foreach ($childId in ($node['Children'] | Sort-Object)) {
            Add-Subtree -NodeId $childId -Depth ($Depth + 1) -Lookup $Lookup -List $List -Visited $Visited
        }
    }

    # Sort roots: type priority
    $typePriority = @{
        'Epic' = 0; 'Feature' = 1; 'User Story' = 2; 'Product Backlog Item' = 2;
        'Incident' = 3; 'Bug' = 4; 'Task' = 5; 'Issue' = 3
    }

    $sortedRoots = $roots | Sort-Object {
        $node = $allLookup[$_]
        $pri = $typePriority[$node.Type]
        if ($null -eq $pri) { $pri = 99 }
        $pri
    }, { $allLookup[$_].Title }

    foreach ($rootId in $sortedRoots) {
        Add-Subtree -NodeId $rootId -Depth 0 -Lookup $allLookup -List $flat -Visited $visited
    }

    return $flat
}

# ── Get work item comments ────────────────────────────────────────
function Get-WorkItemComments {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [int]$WorkItemId
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$WorkItemId/comments?api-version=7.1-preview.4"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        return $result.comments
    }
    catch {
        return @()
    }
}

# ── Convert plain text to HTML for ADO comments ──────────────────
function ConvertTo-CommentHtml {
    param([string]$Text)
    if (-not $Text) { return "" }
    # Escape HTML special characters
    $html = [System.Net.WebUtility]::HtmlEncode($Text)
    # Convert newlines to <br/>
    $html = $html -replace "`r?`n", "<br/>"
    return $html
}

# ── Add a comment to a work item ──────────────────────────────────
function Add-WorkItemComment {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        [string]$Text
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$WorkItemId/comments?format=markdown&api-version=7.1-preview.4"

    $body = @{ text = $Text } | ConvertTo-Json -Depth 3

    try {
        $result = Invoke-RestMethod -Uri $url -Method Post -Headers $headers `
            -ContentType "application/json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
        return $result
    }
    catch {
        throw "Error adding comment: $($_.Exception.Message)"
    }
}

# ── Update a comment on a work item ───────────────────────────────
function Update-WorkItemComment {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        $CommentId,
        [string]$Text
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$WorkItemId/comments/${CommentId}?format=markdown&api-version=7.1-preview.4"

    $body = @{ text = $Text } | ConvertTo-Json -Depth 3

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
        return $result
    }
    catch {
        throw "Error updating comment: $($_.Exception.Message)"
    }
}

# ── Delete a comment from a work item ─────────────────────────────
function Remove-WorkItemComment {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        $CommentId
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$WorkItemId/comments/${CommentId}?api-version=7.1-preview.4"

    try {
        Invoke-RestMethod -Uri $url -Method Delete -Headers $headers -ErrorAction Stop
    }
    catch {
        throw "Error deleting comment: $($_.Exception.Message)"
    }
}

# ── Get work item full details (description etc.) ──────────────────
function Get-WorkItemDetail {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [int]$WorkItemId
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?`$expand=all&api-version=7.1"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        return $result
    }
    catch {
        Write-Host "Error fetching details: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# ── Update work item time fields ──────────────────────────────────
function Update-WorkItemTime {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        [double]$CompletedWork,
        [double]$RemainingWork
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT

    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?api-version=7.1"

    Write-TTDebugLog "Update-WorkItemTime: WI=$WorkItemId raw params: Completed=$CompletedWork Remaining=$RemainingWork"
    Write-TTDebugLog "  URL: $url"

    $patchOps = @(
        [ordered]@{
            op    = "replace"
            path  = "/fields/Microsoft.VSTS.Scheduling.CompletedWork"
            value = $CompletedWork
        },
        [ordered]@{
            op    = "replace"
            path  = "/fields/Microsoft.VSTS.Scheduling.RemainingWork"
            value = $RemainingWork
        }
    )

    $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"

    Write-TTDebugLog "  Body: $body"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json-patch+json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop

        # Log what the API actually returned
        $respC = Get-SafeField -Fields $result.fields -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
        $respR = Get-SafeField -Fields $result.fields -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
        Write-TTDebugLog "  Response: rev=$($result.rev) Completed=$respC Remaining=$respR"

        return $result
    }
    catch {
        $errMsg = $_.Exception.Message
        # Try to capture response body for more detail
        $respBody = ''
        if ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = [System.IO.StreamReader]::new($stream)
                $respBody = $reader.ReadToEnd()
                $reader.Close()
            } catch { }
        }
        Write-TTDebugLog "  ERROR: $errMsg"
        if ($respBody) { Write-TTDebugLog "  Response body: $respBody" }

        # If 'replace' fails (field never set), retry with 'add'
        Write-TTDebugLog "  Retrying with op='add'..."
        $patchOps | ForEach-Object { $_['op'] = 'add' }
        $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"
        Write-TTDebugLog "  Retry Body: $body"

        try {
            $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
                -ContentType "application/json-patch+json" `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop

            $respC = Get-SafeField -Fields $result.fields -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
            $respR = Get-SafeField -Fields $result.fields -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
            Write-TTDebugLog "  Retry Response: rev=$($result.rev) Completed=$respC Remaining=$respR"

            return $result
        }
        catch {
            Write-TTDebugLog "  Retry also failed: $($_.Exception.Message)"
            throw "Error updating work item ${WorkItemId}: $errMsg"
        }
    }
}

# ── Update work item scheduling fields (any combination) ────────
function Update-WorkItemHours {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        [hashtable]$Fields
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?api-version=7.1"

    $patchOps = @()
    foreach ($key in $Fields.Keys) {
        $patchOps += [ordered]@{
            op    = "replace"
            path  = "/fields/$key"
            value = $Fields[$key]
        }
    }

    $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"
    Write-TTDebugLog "Update-WorkItemHours: WI=$WorkItemId Body=$body"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json-patch+json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
        return $result
    }
    catch {
        # Retry with 'add' in case fields never had a value
        $patchOps | ForEach-Object { $_['op'] = 'add' }
        $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"
        Write-TTDebugLog "Update-WorkItemHours retry (add): Body=$body"
        try {
            $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
                -ContentType "application/json-patch+json" `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
            return $result
        }
        catch {
            throw "Error updating hours on work item ${WorkItemId}: $($_.Exception.Message)"
        }
    }
}

# ── Get available states for a work item type ─────────────────────
function Get-WorkItemTypeStates {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [string]$WorkItemType
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $encodedType = [Uri]::EscapeDataString($WorkItemType)
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitemtypes/$encodedType/states?api-version=7.1"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -ErrorAction Stop
        $states = $result.value | Where-Object { $_.category -ne 'Hidden' } | ForEach-Object { $_.name }
        return $states
    }
    catch {
        # Fallback common states
        return @('New', 'Active', 'Resolved', 'Closed')
    }
}

# ── Update work item state ────────────────────────────────────────
function Update-WorkItemState {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        [string]$NewState
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT

    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?api-version=7.1"

    $patchOps = @(
        @{
            op    = "add"
            path  = "/fields/System.State"
            value = $NewState
        }
    )

    $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json-patch+json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
        return $result
    }
    catch {
        throw "Error updating state for work item ${WorkItemId}: $($_.Exception.Message)"
    }
}

# ── Set (or clear) the assignee on a work item ────────────────────
function Set-WorkItemAssignee {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        [AllowNull()][string]$AssignedTo   # empty string or $null to unassign
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?api-version=7.1"

    # ADO requires "" (empty string) to clear an identity field; $null is silently ignored
    $value = if ($AssignedTo) { $AssignedTo } else { "" }

    $patchOps = @(
        [ordered]@{
            op    = "add"
            path  = "/fields/System.AssignedTo"
            value = $value
        }
    )

    $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"
    Write-TTDebugLog "Set-WorkItemAssignee: WI=$WorkItemId AssignedTo='$AssignedTo' Body=$body"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json-patch+json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
        return $result
    }
    catch {
        throw "Error updating assignee for work item ${WorkItemId}: $($_.Exception.Message)"
    }
}

# ── Fetch all members from all teams in a project ─────────────────
function Get-ProjectTeamMembers {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $memberNames = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    $encodedProject = [Uri]::EscapeDataString($Project)
    $teamsUrl = "https://dev.azure.com/$Organization/_apis/projects/$encodedProject/teams?api-version=7.1"

    try {
        $teamsResult = Invoke-RestMethod -Uri $teamsUrl -Method Get -Headers $headers -ErrorAction Stop
        Write-TTDebugLog "Get-ProjectTeamMembers: $($teamsResult.value.Count) teams found"

        foreach ($team in $teamsResult.value) {
            $membersUrl = "https://dev.azure.com/$Organization/_apis/projects/$encodedProject/teams/$($team.id)/members?api-version=7.1"
            try {
                $membersResult = Invoke-RestMethod -Uri $membersUrl -Method Get -Headers $headers -ErrorAction Stop
                foreach ($member in $membersResult.value) {
                    if ($member.identity -and $member.identity.displayName) {
                        [void]$memberNames.Add($member.identity.displayName)
                    }
                }
            }
            catch {
                Write-TTDebugLog "Get-ProjectTeamMembers: team '$($team.name)' members error: $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-TTDebugLog "Get-ProjectTeamMembers: teams list error: $($_.Exception.Message)"
    }

    Write-TTDebugLog "Get-ProjectTeamMembers: returning $($memberNames.Count) unique members"
    return @($memberNames | Sort-Object)
}

# ── Create a child Task linked to a parent work item ─────────────
function New-ChildTask {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [int]$ParentId,
        [string]$Title,
        [string]$AssignedTo,
        [double]$OriginalEstimate = 5,
        [double]$RemainingWork = 5
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/`$Task?api-version=7.1"

    $patchOps = [System.Collections.ArrayList]::new()
    [void]$patchOps.Add([ordered]@{
        op    = "add"
        path  = "/fields/System.Title"
        value = $Title
    })
    if ($AssignedTo) {
        [void]$patchOps.Add([ordered]@{
            op    = "add"
            path  = "/fields/System.AssignedTo"
            value = $AssignedTo
        })
    }
    [void]$patchOps.Add([ordered]@{
        op    = "add"
        path  = "/fields/Microsoft.VSTS.Scheduling.OriginalEstimate"
        value = $OriginalEstimate
    })
    [void]$patchOps.Add([ordered]@{
        op    = "add"
        path  = "/fields/Microsoft.VSTS.Scheduling.RemainingWork"
        value = $RemainingWork
    })
    [void]$patchOps.Add([ordered]@{
        op    = "add"
        path  = "/relations/-"
        value = [ordered]@{
            rel = "System.LinkTypes.Hierarchy-Reverse"
            url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$ParentId"
        }
    })

    $body = ConvertTo-Json -InputObject @($patchOps) -Depth 10 -Compress
    Write-TTDebugLog "New-ChildTask: Parent=$ParentId Title=$Title Body=$body"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json-patch+json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop

        # Set state to Active in a separate call (some processes reject it during creation)
        try {
            $null = Update-WorkItemState -Organization $Organization -Project $Project `
                -PAT $PAT -WorkItemId $result.id -NewState "Active"
        }
        catch {
            Write-TTDebugLog "New-ChildTask: Failed to set Active state: $($_.Exception.Message)"
        }

        return $result
    }
    catch {
        $errDetail = $_.Exception.Message
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = [System.IO.StreamReader]::new($stream)
            $respBody = $reader.ReadToEnd()
            $reader.Close()
            Write-TTDebugLog "New-ChildTask ERROR response: $respBody"
            $errDetail = "$errDetail | $respBody"
        } catch { }
        throw "Error creating child task for work item ${ParentId}: $errDetail"
    }
}

# ── Update a work item field ──────────────────────────────────────
function Update-WorkItemField {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        $WorkItemId,
        [string]$FieldPath,
        [string]$Value
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?api-version=7.1"

    $patchOps = @(
        @{
            op    = "add"
            path  = "/fields/$FieldPath"
            value = $Value
        }
    )

    $body = "[$( ($patchOps | ForEach-Object { $_ | ConvertTo-Json -Depth 5 -Compress }) -join ',' )]"

    try {
        $result = Invoke-RestMethod -Uri $url -Method Patch -Headers $headers `
            -ContentType "application/json-patch+json" `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) -ErrorAction Stop
        return $result
    }
    catch {
        throw "Error updating field '${FieldPath}' on work item ${WorkItemId}: $($_.Exception.Message)"
    }
}
# ── Delete (soft-delete) a work item ───────────────────────────
function Remove-WorkItem {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [int]$WorkItemId
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${WorkItemId}?api-version=7.1"

    try {
        Invoke-RestMethod -Uri $url -Method Delete -Headers $headers -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Error deleting work item ${WorkItemId}: $($_.Exception.Message)"
    }
}

# ── Helper: Convert raw API item to flat work item hashtable ─────
function ConvertTo-FlatWorkItem {
    param($RawItem)
    $f = $RawItem.fields
    $assignedTo = $null
    $assignedToField = Get-SafeField -Fields $f -Name 'System.AssignedTo'
    if ($assignedToField) {
        $assignedTo = $assignedToField.displayName
    }
    return @{
        Id               = $f.'System.Id'
        Title            = $f.'System.Title'
        Type             = $f.'System.WorkItemType'
        State            = $f.'System.State'
        AssignedTo       = $assignedTo
        ParentId         = Get-SafeField -Fields $f -Name 'System.Parent'
        TeamProject      = Get-SafeField -Fields $f -Name 'System.TeamProject'
        OriginalEstimate = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.OriginalEstimate'
        CompletedWork    = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
        RemainingWork    = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
        IsMine           = $true
        IsSeparator      = $false
        IsRelated        = $false
        Depth            = 0
        Children         = @()
        Raw              = $RawItem
    }
}

# ── Helper: Batch-fetch work item details from IDs ───────────────
function Get-WorkItemDetailsFromIds {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [array]$Ids
    )

    if (-not $Ids -or $Ids.Count -eq 0) { return @() }

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $batchSize = 200

    $fields = @(
        "System.Id",
        "System.Title",
        "System.WorkItemType",
        "System.State",
        "System.AssignedTo",
        "System.Description",
        "System.Parent",
        "System.TeamProject",
        "Microsoft.VSTS.Scheduling.OriginalEstimate",
        "Microsoft.VSTS.Scheduling.CompletedWork",
        "Microsoft.VSTS.Scheduling.RemainingWork"
    ) -join ","

    $allItems = [System.Collections.ArrayList]::new()

    for ($i = 0; $i -lt $Ids.Count; $i += $batchSize) {
        $batchIds = $Ids[$i .. [Math]::Min($i + $batchSize - 1, $Ids.Count - 1)]
        $idsString = $batchIds -join ","
        $detailUrl = "$baseUrl/wit/workitems?ids=$idsString&fields=$fields&api-version=7.1"

        try {
            $details = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers -ErrorAction Stop
            foreach ($item in $details.value) {
                [void]$allItems.Add((ConvertTo-FlatWorkItem -RawItem $item))
            }
        }
        catch {
            Write-TTDebugLog "Get-WorkItemDetailsFromIds error: $($_.Exception.Message)"
        }
    }

    return $allItems
}

# ── Search Azure DevOps users by display name prefix ─────────────
function Search-AzDoUsers {
    param(
        [string]$Organization,
        [string]$PAT,
        [string]$SearchTerm,
        [string]$Project = ""   # optional; enables the WIQL fallback
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $encoded = [Uri]::EscapeDataString($SearchTerm)

    # ── Attempt 1: dev.azure.com identities (standard PAT scope) ──
    $url1 = "https://dev.azure.com/$Organization/_apis/identities?searchFilter=General&filterValue=$encoded&queryMembership=None&api-version=7.1"
    try {
        $r1 = Invoke-RestMethod -Uri $url1 -Method Get -Headers $headers -ErrorAction Stop
        Write-TTDebugLog "Search-AzDoUsers identities count: $(if ($r1.value) { $r1.value.Count } else { 0 })"
        if ($r1.value -and $r1.value.Count -gt 0) {
            $names = @($r1.value |
                Where-Object { $_.isActive -ne $false -and $_.providerDisplayName } |
                ForEach-Object { $_.providerDisplayName } |
                Select-Object -Unique | Sort-Object)
            if ($names.Count -gt 0) {
                Write-TTDebugLog "Search-AzDoUsers identities found $($names.Count) names"
                return $names
            }
        }
    }
    catch {
        Write-TTDebugLog "Search-AzDoUsers identities error: $($_.Exception.Message)"
    }

    # ── Attempt 2: IdentityPicker (POST) ──────────────────────────
    $url2 = "https://dev.azure.com/$Organization/_apis/IdentityPicker/Identities?api-version=6.0"
    $body2 = @{
        query           = $SearchTerm
        identityTypes   = @("user")
        operationScopes = @("ims")
        properties      = @("DisplayName", "SubjectDescriptor")
    } | ConvertTo-Json -Depth 5
    try {
        $r2 = Invoke-RestMethod -Uri $url2 -Method Post -Headers $headers `
            -ContentType "application/json" -Body $body2 -ErrorAction Stop
        if ($r2.results -and $r2.results.Count -gt 0) {
            $identities = $r2.results[0].identities
            if ($identities -and $identities.Count -gt 0) {
                $names = @($identities | Where-Object { $_.displayName } |
                    ForEach-Object { $_.displayName } | Select-Object -Unique | Sort-Object)
                if ($names.Count -gt 0) {
                    Write-TTDebugLog "Search-AzDoUsers IdentityPicker found $($names.Count) names"
                    return $names
                }
            }
        }
    }
    catch {
        Write-TTDebugLog "Search-AzDoUsers IdentityPicker error: $($_.Exception.Message)"
    }

    # ── Attempt 3: WIQL – find assignees via work item query ───────
    # Works with any PAT that has work item read access.
    if ($Project) {
        $safe = $SearchTerm -replace "'", "''"
        $wiqlBody = @{
            query = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = @project AND [System.AssignedTo] Contains '$safe' AND [System.State] <> 'Removed' ORDER BY [System.AssignedTo]"
        } | ConvertTo-Json
        $wiqlUrl = "https://dev.azure.com/$Organization/$Project/_apis/wit/wiql?`$top=50&api-version=7.1"
        try {
            $wResult = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
                -ContentType "application/json" -Body $wiqlBody -ErrorAction Stop
            Write-TTDebugLog "Search-AzDoUsers WIQL items: $(if ($wResult.workItems) { $wResult.workItems.Count } else { 0 })"
            if ($wResult.workItems -and $wResult.workItems.Count -gt 0) {
                $ids = @($wResult.workItems | ForEach-Object { $_.id } | Select-Object -First 50)
                $idsStr = $ids -join ","
                $detailUrl = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems?ids=$idsStr&fields=System.AssignedTo&api-version=7.1"
                $details = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers -ErrorAction Stop
                $names = @($details.value | ForEach-Object {
                    $f = Get-SafeField -Fields $_.fields -Name 'System.AssignedTo'
                    if ($f -and $f.displayName) { $f.displayName }
                } | Select-Object -Unique | Sort-Object)
                if ($names.Count -gt 0) {
                    Write-TTDebugLog "Search-AzDoUsers WIQL found $($names.Count) names"
                    return $names
                }
            }
        }
        catch {
            Write-TTDebugLog "Search-AzDoUsers WIQL error: $($_.Exception.Message)"
        }
    }

    Write-TTDebugLog "Search-AzDoUsers: all attempts exhausted, returning empty"
    return @()
}

# ── Get work items I've been mentioned in ─────────────────────────
function Get-MentionedWorkItems {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [switch]$IncludeClosed
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"

    $stateFilter = ""
    if (-not $IncludeClosed) {
        $stateFilter = @"

  AND [System.State] <> 'Closed'
  AND [System.State] <> 'Removed'
  AND [System.State] <> 'Done'
  AND [System.State] <> 'Released'
"@
    }

    $queryText = @"
SELECT [System.Id], [System.Title]
FROM WorkItems
WHERE [System.Id] IN (@RecentMentions)$stateFilter
ORDER BY [System.ChangedDate] DESC
"@

    $wiql = @{ query = $queryText } | ConvertTo-Json
    Write-TTDebugLog "Get-MentionedWorkItems WIQL: $queryText"

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiql -ErrorAction Stop
    }
    catch {
        Write-TTDebugLog "Get-MentionedWorkItems error: $($_.Exception.Message)"
        return @()
    }

    if (-not $result.workItems -or $result.workItems.Count -eq 0) {
        return @()
    }

    $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 200)
    return @(Get-WorkItemDetailsFromIds -Organization $Organization -Project $Project -PAT $PAT -Ids $ids)
}

# ── Get work items I'm following ──────────────────────────────────
function Get-FollowingWorkItems {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [switch]$IncludeClosed
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"

    $stateFilter = ""
    if (-not $IncludeClosed) {
        $stateFilter = @"

  AND [System.State] <> 'Closed'
  AND [System.State] <> 'Removed'
  AND [System.State] <> 'Done'
  AND [System.State] <> 'Released'
"@
    }

    $queryText = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.Id] IN (@Follows)$stateFilter
ORDER BY [System.ChangedDate] DESC
"@

    $wiql = @{ query = $queryText } | ConvertTo-Json
    Write-TTDebugLog "Get-FollowingWorkItems WIQL: $queryText"

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiql -ErrorAction Stop
    }
    catch {
        Write-TTDebugLog "Get-FollowingWorkItems error: $($_.Exception.Message)"
        return @()
    }

    if (-not $result.workItems -or $result.workItems.Count -eq 0) {
        return @()
    }

    $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 200)
    return @(Get-WorkItemDetailsFromIds -Organization $Organization -Project $Project -PAT $PAT -Ids $ids)
}

# ── Get work items created by me ─────────────────────────────────
function Get-CreatedByMeWorkItems {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [switch]$IncludeClosed
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"

    $stateFilter = ""
    if (-not $IncludeClosed) {
        $stateFilter = @"

  AND [System.State] <> 'Closed'
  AND [System.State] <> 'Removed'
  AND [System.State] <> 'Done'
  AND [System.State] <> 'Released'
"@
    }

    $queryText = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.CreatedBy] = @me$stateFilter
ORDER BY [System.CreatedDate] DESC
"@

    $wiql = @{ query = $queryText } | ConvertTo-Json
    Write-TTDebugLog "Get-CreatedByMeWorkItems WIQL: $queryText"

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiql -ErrorAction Stop
    }
    catch {
        Write-TTDebugLog "Get-CreatedByMeWorkItems error: $($_.Exception.Message)"
        return @()
    }

    if (-not $result.workItems -or $result.workItems.Count -eq 0) {
        return @()
    }

    $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 200)
    return @(Get-WorkItemDetailsFromIds -Organization $Organization -Project $Project -PAT $PAT -Ids $ids)
}

# ── Search work items with filters ───────────────────────────────
function Search-WorkItemsByFilters {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [string]$TitleContains = "",
        [string]$StateFilter = "",
        [string]$TypeFilter = "",
        [string]$AssignedTo = "",
        [int]$WorkItemId = 0
    )

    # If searching by ID, just fetch that specific item
    if ($WorkItemId -gt 0) {
        $item = Get-WorkItemDetail -Organization $Organization -Project $Project `
            -PAT $PAT -WorkItemId $WorkItemId
        if ($item) {
            return @(ConvertTo-FlatWorkItem -RawItem $item)
        }
        return @()
    }

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"

    $conditions = [System.Collections.ArrayList]::new()

    if ($TitleContains) {
        [void]$conditions.Add("[System.Title] Contains '$($TitleContains -replace "'", "''")'")
    }

    if ($StateFilter -and $StateFilter -ne 'Any') {
        [void]$conditions.Add("[System.State] = '$($StateFilter -replace "'", "''")'")
    }

    if ($TypeFilter -and $TypeFilter -ne 'Any') {
        [void]$conditions.Add("[System.WorkItemType] = '$($TypeFilter -replace "'", "''")'")
    }

    if ($AssignedTo) {
        if ($AssignedTo -eq '@me') {
            [void]$conditions.Add("[System.AssignedTo] = @me")
        }
        else {
            [void]$conditions.Add("[System.AssignedTo] Contains '$($AssignedTo -replace "'", "''")'")
        }
    }

    if ($conditions.Count -eq 0) {
        return @()
    }

    $whereClause = $conditions -join " AND "
    $queryText = "SELECT [System.Id] FROM WorkItems WHERE $whereClause ORDER BY [System.ChangedDate] DESC"

    $wiql = @{ query = $queryText } | ConvertTo-Json
    Write-TTDebugLog "Search-WorkItemsByFilters WIQL: $queryText"

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiql -ErrorAction Stop
    }
    catch {
        Write-TTDebugLog "Search-WorkItemsByFilters error: $($_.Exception.Message)"
        return @()
    }

    if (-not $result.workItems -or $result.workItems.Count -eq 0) {
        return @()
    }

    $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 200)
    return @(Get-WorkItemDetailsFromIds -Organization $Organization -Project $Project -PAT $PAT -Ids $ids)
}

# ── Get children and related items for a parent work item (Pri tab) ─
function Get-PriWorkItems {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT,
        [int]$ParentId,
        [switch]$IncludeClosed
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"

    # ── 1. Fetch parent with $expand=all to get its relations ─────
    $parentUrl = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/${ParentId}?`$expand=all&api-version=7.1"
    $parentRaw = $null
    try {
        $parentRaw = Invoke-RestMethod -Uri $parentUrl -Method Get -Headers $headers -ErrorAction Stop
    }
    catch {
        Write-TTDebugLog "Get-PriWorkItems: error fetching parent $ParentId : $($_.Exception.Message)"
        return @()
    }

    # ── 2. WIQL recursive query for all descendants ───────────────
    $wiqlDescendants = @{
        query = @"
SELECT [System.Id]
FROM WorkItemLinks
WHERE [Source].[System.Id] = $ParentId
  AND [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward'
  AND [Target].[System.TeamProject] = @project
  AND [Target].[System.State] <> 'Removed'
MODE (Recursive)
"@
    } | ConvertTo-Json

    # linkParentMap: child-id -> direct-parent-id
    $linkParentMap = @{}
    $descendantIds = [System.Collections.ArrayList]::new()

    try {
        $wiqlResult = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $wiqlDescendants -ErrorAction Stop

        foreach ($link in $wiqlResult.workItemRelations) {
            if ($null -eq $link.target) { continue }
            $targetId = $link.target.id
            if ($targetId -eq $ParentId) { continue }
            $sourceId = if ($null -ne $link.source) { $link.source.id } else { $ParentId }
            $linkParentMap[$targetId] = $sourceId
            if (-not $descendantIds.Contains($targetId)) {
                [void]$descendantIds.Add($targetId)
            }
        }
    }
    catch {
        Write-TTDebugLog "Get-PriWorkItems: WIQL descendants error: $($_.Exception.Message)"
    }

    # ── 3. Collect related/dependency IDs from parent's relations ──
    $relatedIds = [System.Collections.ArrayList]::new()
    if ($parentRaw.relations) {
        foreach ($rel in $parentRaw.relations) {
            $relType = $rel.rel
            if ($rel.url -match '/workitems/(\d+)') {
                $wiId = [int]$Matches[1]
                if ($wiId -eq $ParentId) { continue }
                if ($relType -in @(
                        'System.LinkTypes.Related',
                        'System.LinkTypes.Dependency-Forward',
                        'System.LinkTypes.Dependency-Reverse',
                        'Microsoft.VSTS.Common.Affects-Forward',
                        'Microsoft.VSTS.Common.Affects-Reverse'
                    ) -and -not $descendantIds.Contains($wiId)) {
                    [void]$relatedIds.Add($wiId)
                }
            }
        }
    }

    # ── 4. Build depth map iteratively ────────────────────────────
    $depthMap = @{ [int]$ParentId = 0 }
    $iterations = 0
    $maxIter = 20
    $depthChanged = $true
    while ($depthChanged -and $iterations -lt $maxIter) {
        $depthChanged = $false
        $iterations++
        foreach ($childId in @($linkParentMap.Keys)) {
            if ($depthMap.ContainsKey([int]$childId)) { continue }
            $parentOfChild = $linkParentMap[$childId]
            if ($depthMap.ContainsKey([int]$parentOfChild)) {
                $depthMap[[int]$childId] = $depthMap[[int]$parentOfChild] + 1
                $depthChanged = $true
            }
        }
    }

    # ── 5. Build result list ──────────────────────────────────────
    $resultList = [System.Collections.ArrayList]::new()

    # Parent at depth 0
    $parentItem = ConvertTo-FlatWorkItem -RawItem $parentRaw
    $parentItem['Depth'] = 0
    [void]$resultList.Add($parentItem)

    # Closed states to exclude when $IncludeClosed is $false
    $closedStates = @('Closed', 'Done', 'Released', 'Removed')

    # Descendants in WIQL traversal order (preserves tree structure)
    if ($descendantIds.Count -gt 0) {
        $descItems = @(Get-WorkItemDetailsFromIds -Organization $Organization `
            -Project $Project -PAT $PAT -Ids @($descendantIds))

        # Filter out closed items unless IncludeClosed is set
        if (-not $IncludeClosed) {
            $descItems = @($descItems | Where-Object { $closedStates -notcontains $_.State })
        }

        # Build id->item map for sorted insertion
        $itemById = @{}
        foreach ($di in $descItems) { $itemById[[int]$di.Id] = $di }

        foreach ($dId in $descendantIds) {
            $di = $itemById[[int]$dId]
            if ($null -eq $di) { continue }
            $depth = if ($depthMap.ContainsKey([int]$dId)) { $depthMap[[int]$dId] } else { 1 }
            $di['Depth'] = $depth
            [void]$resultList.Add($di)
        }
    }

    # Related items section
    if ($relatedIds.Count -gt 0) {
        $uniqueRelated = @($relatedIds | Select-Object -Unique)
        $relItems = @(Get-WorkItemDetailsFromIds -Organization $Organization `
            -Project $Project -PAT $PAT -Ids $uniqueRelated)

        # Filter out closed related items unless IncludeClosed is set
        if (-not $IncludeClosed) {
            $relItems = @($relItems | Where-Object { $closedStates -notcontains $_.State })
        }

        if ($relItems.Count -gt 0) {
            [void]$resultList.Add(@{
                Id               = $null
                Title            = "--- Related ---"
                Type             = "__separator__"
                State            = ""
                AssignedTo       = ""
                ParentId         = $null
                OriginalEstimate = $null
                CompletedWork    = $null
                RemainingWork    = $null
                Depth            = 0
                IsMine           = $false
                IsSeparator      = $true
                IsRelated        = $false
                Children         = @()
                Raw              = $null
            })

            foreach ($ri in $relItems) {
                $ri['Depth'] = 0
                $ri['IsRelated'] = $true
                [void]$resultList.Add($ri)
            }
        }
    }

    return $resultList
}

# ── Get SCRUM report data (yesterday's activity & today's plan) ────
function Get-ScrumReportData {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $baseUrl = "https://dev.azure.com/$Organization/$Project/_apis"
    $wiqlUrl = "$baseUrl/wit/wiql?api-version=7.1"
    $myDisplayName = Get-CurrentUserDisplayName -Organization $Organization -PAT $PAT
    $yesterday = (Get-Date).AddDays(-1).Date

    Write-TTDebugLog "Get-ScrumReportData: myDisplayName=$myDisplayName yesterday=$($yesterday.ToString('yyyy-MM-dd'))"

    # ── 1. Items I closed/resolved yesterday ─────────────────────────
    $closedQuery = @{
        query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.ChangedBy] = @me
  AND [System.ChangedDate] >= @Today - 1
  AND [System.ChangedDate] < @Today
  AND ([System.State] = 'Closed'
    OR [System.State] = 'Done'
    OR [System.State] = 'Resolved'
    OR [System.State] = 'Released')
ORDER BY [System.ChangedDate] DESC
"@
    } | ConvertTo-Json

    $closedItems = @()
    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $closedQuery -ErrorAction Stop
        if ($result.workItems -and $result.workItems.Count -gt 0) {
            $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 100)
            $closedItems = @(Get-WorkItemDetailsFromIds -Organization $Organization `
                -Project $Project -PAT $PAT -Ids $ids)
        }
        Write-TTDebugLog "Get-ScrumReportData: closedItems=$($closedItems.Count)"
    }
    catch {
        Write-TTDebugLog "Get-ScrumReportData closed error: $($_.Exception.Message)"
    }

    $closedIds = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($ci in $closedItems) { [void]$closedIds.Add([int]$ci.Id) }

    # ── 2. Items I edited yesterday (still open) ─────────────────────
    $changedQuery = @{
        query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.ChangedBy] = @me
  AND [System.ChangedDate] >= @Today - 1
  AND [System.ChangedDate] < @Today
  AND [System.State] <> 'Closed'
  AND [System.State] <> 'Done'
  AND [System.State] <> 'Resolved'
  AND [System.State] <> 'Released'
  AND [System.State] <> 'Removed'
ORDER BY [System.ChangedDate] DESC
"@
    } | ConvertTo-Json

    $changedItems = @()
    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $changedQuery -ErrorAction Stop
        if ($result.workItems -and $result.workItems.Count -gt 0) {
            $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 100)
            $ids = @($ids | Where-Object { -not $closedIds.Contains([int]$_) })
            if ($ids.Count -gt 0) {
                $changedItems = @(Get-WorkItemDetailsFromIds -Organization $Organization `
                    -Project $Project -PAT $PAT -Ids $ids)
            }
        }
        Write-TTDebugLog "Get-ScrumReportData: changedItems=$($changedItems.Count)"
    }
    catch {
        Write-TTDebugLog "Get-ScrumReportData changed error: $($_.Exception.Message)"
    }

    # ── 3. Check comments to split commented vs edited ───────────────
    $commentedItems = [System.Collections.ArrayList]::new()
    $editedItems = [System.Collections.ArrayList]::new()
    $allYesterdayIds = [System.Collections.Generic.HashSet[int]]::new($closedIds)

    foreach ($item in $changedItems) {
        [void]$allYesterdayIds.Add([int]$item.Id)
        $hasMyComment = $false
        try {
            $comments = @(Get-WorkItemComments -Organization $Organization `
                -Project $Project -PAT $PAT -WorkItemId $item.Id)
            foreach ($comment in $comments) {
                if ($comment.createdBy -and $comment.createdBy.displayName -eq $myDisplayName) {
                    $commentDate = ([datetime]$comment.createdDate).Date
                    if ($commentDate -eq $yesterday) {
                        $hasMyComment = $true
                        break
                    }
                }
            }
        }
        catch {
            Write-TTDebugLog "Get-ScrumReportData comment check error #$($item.Id): $($_.Exception.Message)"
        }

        if ($hasMyComment) {
            [void]$commentedItems.Add($item)
        }
        else {
            [void]$editedItems.Add($item)
        }
    }

    Write-TTDebugLog "Get-ScrumReportData: commentedItems=$($commentedItems.Count) editedItems=$($editedItems.Count)"

    # ── Helper: check if a work item has unanswered comments from others ─
    # Returns $true ONLY when there is at least one real comment from
    # someone other than $MyName AND we have NOT commented after it.
    function Test-HasUnansweredComment {
        param($Organization, $Project, $PAT, [int]$WorkItemId, [string]$MyName)

        $rawComments = $null
        try {
            $rawComments = Get-WorkItemComments -Organization $Organization `
                -Project $Project -PAT $PAT -WorkItemId $WorkItemId
        }
        catch {
            Write-TTDebugLog "Test-HasUnansweredComment #$WorkItemId error: $($_.Exception.Message)"
            return $false
        }

        if ($null -eq $rawComments) { return $false }

        # Build a clean array of only real comments
        $comments = [System.Collections.ArrayList]::new()
        foreach ($c in $rawComments) {
            if ($null -eq $c) { continue }
            if (-not $c.createdBy) { continue }
            if (-not $c.createdBy.displayName) { continue }
            [void]$comments.Add($c)
        }
        Write-TTDebugLog "Test-HasUnansweredComment #$WorkItemId : $($comments.Count) real comments"

        if ($comments.Count -eq 0) { return $false }

        # Find the last comment by someone other than me
        $lastOtherComment = $null
        for ($i = $comments.Count - 1; $i -ge 0; $i--) {
            if ($comments[$i].createdBy.displayName -ne $MyName) {
                $lastOtherComment = $comments[$i]
                break
            }
        }
        if ($null -eq $lastOtherComment) { return $false }

        # Check if I commented after that
        $otherDate = [datetime]$lastOtherComment.createdDate
        for ($i = $comments.Count - 1; $i -ge 0; $i--) {
            $c = $comments[$i]
            if ($c.createdBy.displayName -eq $MyName -and [datetime]$c.createdDate -gt $otherDate) {
                return $false
            }
        }

        return $true
    }

    # ── 4. Mentioned yesterday but not responded ─────────────────────
    $mentionQuery = @{
        query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.Id] IN (@RecentMentions)
  AND [System.ChangedDate] >= @Today - 1
  AND [System.ChangedDate] < @Today
ORDER BY [System.ChangedDate] DESC
"@
    } | ConvertTo-Json

    $followUpItems = [System.Collections.ArrayList]::new()
    $followUpIds = [System.Collections.Generic.HashSet[int]]::new()

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $mentionQuery -ErrorAction Stop
        if ($result.workItems -and $result.workItems.Count -gt 0) {
            $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 100)
            $ids = @($ids | Where-Object { -not $allYesterdayIds.Contains([int]$_) })
            if ($ids.Count -gt 0) {
                $mentionedItems = @(Get-WorkItemDetailsFromIds -Organization $Organization `
                    -Project $Project -PAT $PAT -Ids $ids)
                foreach ($item in $mentionedItems) {
                    if (Test-HasUnansweredComment -Organization $Organization `
                            -Project $Project -PAT $PAT `
                            -WorkItemId $item.Id -MyName $myDisplayName) {
                        [void]$followUpItems.Add($item)
                        [void]$followUpIds.Add([int]$item.Id)
                    }
                }
            }
        }
        Write-TTDebugLog "Get-ScrumReportData: followUpItems (mentions)=$($followUpItems.Count)"
    }
    catch {
        Write-TTDebugLog "Get-ScrumReportData mentions error: $($_.Exception.Message)"
    }

    # ── 4b. Items assigned to me with unanswered comments from others ─
    $assignedOpenQuery = @{
        query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.AssignedTo] = @me
  AND [System.State] <> 'Closed'
  AND [System.State] <> 'Done'
  AND [System.State] <> 'Released'
  AND [System.State] <> 'Removed'
ORDER BY [System.ChangedDate] DESC
"@
    } | ConvertTo-Json

    try {
        $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
            -ContentType "application/json" -Body $assignedOpenQuery -ErrorAction Stop
        if ($result.workItems -and $result.workItems.Count -gt 0) {
            $ids = @($result.workItems | ForEach-Object { $_.id } | Select-Object -First 200)
            $ids = @($ids | Where-Object {
                -not $allYesterdayIds.Contains([int]$_) -and
                -not $followUpIds.Contains([int]$_)
            })
            if ($ids.Count -gt 0) {
                $assignedItems = @(Get-WorkItemDetailsFromIds -Organization $Organization `
                    -Project $Project -PAT $PAT -Ids $ids)
                foreach ($item in $assignedItems) {
                    if (Test-HasUnansweredComment -Organization $Organization `
                            -Project $Project -PAT $PAT `
                            -WorkItemId $item.Id -MyName $myDisplayName) {
                        [void]$followUpItems.Add($item)
                        [void]$followUpIds.Add([int]$item.Id)
                    }
                }
            }
        }
        Write-TTDebugLog "Get-ScrumReportData: followUpItems (total)=$($followUpItems.Count)"
    }
    catch {
        Write-TTDebugLog "Get-ScrumReportData assigned-followup error: $($_.Exception.Message)"
    }

    # ── 5. Items in the "Active" board column on my Teams' boards ────
    $activeItems = @()
    try {
        $activeItems = @(Get-MyTeamBoardActiveItems -Organization $Organization `
            -Project $Project -PAT $PAT)
        Write-TTDebugLog "Get-ScrumReportData: activeItems=$($activeItems.Count)"
    }
    catch {
        Write-TTDebugLog "Get-ScrumReportData active error: $($_.Exception.Message)"
    }

    return @{
        ClosedItems    = @($closedItems)
        CommentedItems = @($commentedItems)
        EditedItems    = @($editedItems)
        FollowUpItems  = @($followUpItems)
        ActiveItems    = @($activeItems)
    }
}

# ── Get items in the "Active" board column on teams I belong to ────
function Get-MyTeamBoardActiveItems {
    param(
        [string]$Organization,
        [string]$Project,
        [string]$PAT
    )

    $headers = Get-AzDoAuthHeader -PAT $PAT
    $encodedProject = [Uri]::EscapeDataString($Project)

    # Get teams I am a member of ($mine=true filters server-side)
    $teamsUrl = "https://dev.azure.com/$Organization/_apis/projects/$encodedProject/teams?`$mine=true&api-version=7.1"
    $myTeams = @()
    try {
        $teamsResult = Invoke-RestMethod -Uri $teamsUrl -Method Get -Headers $headers -ErrorAction Stop
        $myTeams = @($teamsResult.value)
        Write-TTDebugLog "Get-MyTeamBoardActiveItems: $($myTeams.Count) teams (mine)"
    }
    catch {
        Write-TTDebugLog "Get-MyTeamBoardActiveItems: teams error: $($_.Exception.Message)"
        return @()
    }

    if ($myTeams.Count -eq 0) { return @() }

    # For each team, run a WIQL query scoped to that team so
    # System.BoardColumn resolves against that team's board.
    $allActiveIds = [System.Collections.Generic.HashSet[int]]::new()

    foreach ($team in $myTeams) {
        $encodedTeam = [Uri]::EscapeDataString($team.name)
        $wiqlUrl = "https://dev.azure.com/$Organization/$encodedProject/$encodedTeam/_apis/wit/wiql?api-version=7.1"

        $wiql = @{
            query = @"
SELECT [System.Id]
FROM WorkItems
WHERE [System.AssignedTo] = @me
  AND [System.BoardColumn] = 'Active'
ORDER BY [System.ChangedDate] DESC
"@
        } | ConvertTo-Json

        try {
            $result = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers `
                -ContentType "application/json" -Body $wiql -ErrorAction Stop
            if ($result.workItems) {
                foreach ($wi in $result.workItems) {
                    [void]$allActiveIds.Add([int]$wi.id)
                }
            }
            Write-TTDebugLog "Get-MyTeamBoardActiveItems: team '$($team.name)' -> $($result.workItems.Count) items"
        }
        catch {
            Write-TTDebugLog "Get-MyTeamBoardActiveItems: team '$($team.name)' WIQL error: $($_.Exception.Message)"
        }
    }

    if ($allActiveIds.Count -eq 0) { return @() }

    $ids = @($allActiveIds)
    return @(Get-WorkItemDetailsFromIds -Organization $Organization -Project $Project -PAT $PAT -Ids $ids)
}