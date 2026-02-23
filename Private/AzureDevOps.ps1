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
        $allLookup[$id] = @{
            Id              = $id
            Title           = $f.'System.Title'
            Type            = $f.'System.WorkItemType'
            State           = $f.'System.State'
            ParentId        = Get-SafeField -Fields $f -Name 'System.Parent'
            OriginalEstimate = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.OriginalEstimate'
            CompletedWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
            RemainingWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
            IsMine          = $true
            Depth           = 0
            Children        = @()
            Raw             = $item
        }
    }

    foreach ($item in $ParentItems) {
        $f = $item.fields
        $id = $f.'System.Id'
        if (-not $allLookup.ContainsKey($id)) {
            $allLookup[$id] = @{
                Id              = $id
                Title           = $f.'System.Title'
                Type            = $f.'System.WorkItemType'
                State           = $f.'System.State'
                ParentId        = Get-SafeField -Fields $f -Name 'System.Parent'
                OriginalEstimate = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.OriginalEstimate'
                CompletedWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
                RemainingWork   = Get-SafeField -Fields $f -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
                IsMine          = $false
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
