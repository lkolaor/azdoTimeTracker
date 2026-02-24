function Start-TimeTracker {
    <#
    .SYNOPSIS
        Launches the Azure DevOps Time Tracker interactive TUI.

    .DESCRIPTION
        An interactive terminal application for tracking time against Azure DevOps
        work items (Epics, Features, User Stories, Tasks, Bugs, Incidents).

        On first run you will be prompted to configure your Azure DevOps organization,
        project, and Personal Access Token (PAT). The configuration is stored in your
        user config directory (~/.config/AzDoTimeTracker/config.json on Linux/macOS,
        %APPDATA%\AzDoTimeTracker\config.json on Windows).

    .PARAMETER Reconfigure
        Force the interactive configuration setup, even if a valid config exists.

    .EXAMPLE
        Start-TimeTracker

        Launches the time tracker with the saved configuration.

    .EXAMPLE
        Start-TimeTracker -Reconfigure

        Re-prompts for organization, project, and PAT before launching.

    .LINK
        https://dev.azure.com
    #>
    [CmdletBinding()]
    param(
        [switch]$Reconfigure
    )

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # Enable debug logging when -Debug is passed
    $script:TTDebugEnabled = $DebugPreference -ne 'SilentlyContinue'
    if ($script:TTDebugEnabled) {
        $logPath = Join-Path (Get-TTConfigDir) "debug.log"
        Add-Content -Path $logPath -Value "" -ErrorAction SilentlyContinue
        Add-Content -Path $logPath -Value "=== Debug session started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" -ErrorAction SilentlyContinue
    }

    # ── Initialize configuration ─────────────────────────────────────
    if ($Reconfigure) {
        $config = Request-TTConfig
    }
    else {
        $config = Initialize-TTConfig
    }

    # ── Fetch work items ─────────────────────────────────────────────
    function Refresh-TabItems {
        param($Config, [int]$TabIndex = 0, [bool]$ShowAll = $false, $PriData = $null)

        switch ($TabIndex) {
            0 {
                $result = Get-MyWorkItems -Organization $Config.Organization `
                                          -Project $Config.Project `
                                          -PAT $Config.PAT
                if (-not $result -or ($result.MyItems.Count -eq 0)) {
                    return [System.Collections.ArrayList]::new()
                }
                return Build-WorkItemTree -MyItems $result.MyItems -ParentItems $result.ParentItems
            }
            1 {
                return [System.Collections.ArrayList]@(
                    Get-MentionedWorkItems -Organization $Config.Organization `
                        -Project $Config.Project -PAT $Config.PAT `
                        -IncludeClosed:$ShowAll
                )
            }
            2 {
                return [System.Collections.ArrayList]@(
                    Get-FollowingWorkItems -Organization $Config.Organization `
                        -Project $Config.Project -PAT $Config.PAT `
                        -IncludeClosed:$ShowAll
                )
            }
            3 {
                return [System.Collections.ArrayList]@(
                    Get-CreatedByMeWorkItems -Organization $Config.Organization `
                        -Project $Config.Project -PAT $Config.PAT `
                        -IncludeClosed:$ShowAll
                )
            }
            4 {
                return [System.Collections.ArrayList]::new()
            }
            5 {
                if ($PriData -and $PriData.HasParent) {
                    return [System.Collections.ArrayList]@(
                        Get-PriWorkItems -Organization $Config.Organization `
                            -Project $Config.Project -PAT $Config.PAT `
                            -ParentId $PriData.ParentId `
                            -IncludeClosed:$ShowAll
                    )
                }
                return [System.Collections.ArrayList]::new()
            }
        }
    }

    $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex 0)

    # ── State ────────────────────────────────────────────────────────
    $selectedIndex = 0
    $scrollOffset = 0
    $statusMessage = ""
    $mode = "list"  # list | detail | statuspicker | assigneepicker | commentpicker | fieldpicker | toolsmenu | queryform | priform
    $detailScrollOffset = 0
    $detailData = $null
    $activeTimers = @{}          # hashtable: workItemId -> Stopwatch
    $statusPickerData = $null    # @{ Item; Statuses; SelectedIndex }
    $commentPickerData = $null   # @{ Item; Comments; SelectedIndex; Action }
    $fieldPickerData = $null     # @{ Item; Fields; SelectedIndex }
    $assigneePickerData = $null   # @{ Item; SearchText; Results; SelectedIndex; LastSearchText }
    $toolsMenuData = $null       # @{ SelectedIndex; MenuItems }

    # ── Tab State ────────────────────────────────────────────────────
    $tabNames = @("Mine", "Mentions", "Following", "Created by me", "Query", "Pri")
    $activeTab = 0
    $tabState = @(
        @{ Items = $items; SelectedIndex = 0; ScrollOffset = 0; ShowAll = $false; Loaded = $true;  EmptyMessage = "No work items found assigned to you." }
        @{ Items = $null;  SelectedIndex = 0; ScrollOffset = 0; ShowAll = $false; Loaded = $false; EmptyMessage = "No work items mentioning you." }
        @{ Items = $null;  SelectedIndex = 0; ScrollOffset = 0; ShowAll = $false; Loaded = $false; EmptyMessage = "No followed work items." }
        @{ Items = $null;  SelectedIndex = 0; ScrollOffset = 0; ShowAll = $false; Loaded = $false; EmptyMessage = "No work items created by you." }
        @{ Items = $null;  SelectedIndex = 0; ScrollOffset = 0; ShowAll = $false; Loaded = $true;  EmptyMessage = "Press / to search for work items." }
        @{ Items = $null;  SelectedIndex = 0; ScrollOffset = 0; ShowAll = $false; Loaded = $false; EmptyMessage = "No parent selected. Press / to set a parent work item." }
    )
    $queryData = @{
        TitleContains     = ""
        State             = ""
        Type              = ""
        AssignedTo        = ""
        WorkItemId        = ""
        FormSelectedIndex = 0
        HasSearched       = $false
    }
    $priData = @{
        ParentId          = 0
        ParentTitle       = ""
        HasParent         = $false
        SearchInput       = ""
        SearchResults     = @()
        SearchResultIndex = 0
        FormState         = "input"   # "input" or "results"
    }

    # Load persisted Pri parent from config (use PSObject.Properties to avoid StrictMode errors on old configs)
    $savedParentId = 0
    $savedParentTitle = ""
    $priParentIdProp = $config.PSObject.Properties['PriParentId']
    if ($null -ne $priParentIdProp -and $null -ne $priParentIdProp.Value) {
        $parsedId = 0
        if ([int]::TryParse([string]$priParentIdProp.Value, [ref]$parsedId) -and $parsedId -gt 0) {
            $savedParentId = $parsedId
        }
    }
    $priParentTitleProp = $config.PSObject.Properties['PriParentTitle']
    if ($null -ne $priParentTitleProp -and $priParentTitleProp.Value) {
        $savedParentTitle = [string]$priParentTitleProp.Value
    }
    if ($savedParentId -gt 0) {
        $priData.ParentId    = $savedParentId
        $priData.ParentTitle = if ($savedParentTitle) { $savedParentTitle } else { "#$savedParentId" }
        $priData.HasParent   = $true
        $tabState[5].Loaded  = $false   # will load on first visit to tab
        $tabState[5].EmptyMessage = "No children or related items found for #$savedParentId."
    }

    # ── Save a single timer ──────────────────────────────────────────
    function Save-Timer {
        param($Item, [System.Diagnostics.Stopwatch]$Stopwatch, $Config)

        $Stopwatch.Stop()
        $elapsedHours = [Math]::Round($Stopwatch.Elapsed.TotalHours, 2)

        if ($elapsedHours -lt 0.01) {
            return "Timer for #$($Item.Id) cancelled (< 1 min)"
        }

        # Fetch fresh values from Azure DevOps
        $freshItem = Get-WorkItemDetail -Organization $Config.Organization `
            -Project $Config.Project -PAT $Config.PAT `
            -WorkItemId $Item.Id

        $currentCompleted    = 0.0
        $currentRemaining    = 0.0
        $originalEstimate    = $null
        if ($null -ne $freshItem) {
            $fields = $freshItem.fields
            if ($null -ne $fields) {
                $cwProp = $fields.PSObject.Properties['Microsoft.VSTS.Scheduling.CompletedWork']
                if ($null -ne $cwProp -and $null -ne $cwProp.Value) { $currentCompleted = [double]$cwProp.Value }
                $rwProp = $fields.PSObject.Properties['Microsoft.VSTS.Scheduling.RemainingWork']
                if ($null -ne $rwProp -and $null -ne $rwProp.Value) { $currentRemaining = [double]$rwProp.Value }
                $oeProp = $fields.PSObject.Properties['Microsoft.VSTS.Scheduling.OriginalEstimate']
                if ($null -ne $oeProp -and $null -ne $oeProp.Value) { $originalEstimate = [double]$oeProp.Value }
            }
        }

        Write-TTDebugLog "Save-Timer: WI=$($Item.Id) elapsed=${elapsedHours}h fresh_C=$currentCompleted fresh_R=$currentRemaining OE=$originalEstimate"

        [double]$newCompleted = $currentCompleted + $elapsedHours
        # Remaining = OriginalEstimate - newCompleted; fall back to currentRemaining - elapsed if no estimate
        [double]$newRemaining = if ($null -ne $originalEstimate) {
            $originalEstimate - $newCompleted
        } else {
            $currentRemaining - $elapsedHours
        }
        if ($newRemaining -lt 0) { $newRemaining = 0.0 }

        Write-TTDebugLog "Save-Timer: newCompleted=$newCompleted newRemaining=$newRemaining (OE-based=$($null -ne $originalEstimate))"

        try {
            $apiResult = Update-WorkItemTime -Organization $Config.Organization `
                -Project $Config.Project -PAT $Config.PAT `
                -WorkItemId $Item.Id `
                -CompletedWork $newCompleted `
                -RemainingWork $newRemaining

            # Verify what the API actually set
            $actualC = $newCompleted
            $actualR = $newRemaining
            if ($null -ne $apiResult) {
                $apiFields = $apiResult.fields
                if ($null -ne $apiFields) {
                    $acProp = $apiFields.PSObject.Properties['Microsoft.VSTS.Scheduling.CompletedWork']
                    if ($null -ne $acProp -and $null -ne $acProp.Value) { $actualC = [double]$acProp.Value }
                    $arProp = $apiFields.PSObject.Properties['Microsoft.VSTS.Scheduling.RemainingWork']
                    if ($null -ne $arProp -and $null -ne $arProp.Value) { $actualR = [double]$arProp.Value }
                }
            }

            $Item['CompletedWork'] = $actualC
            $Item['RemainingWork'] = $actualR

            $oeStr = if ($null -ne $originalEstimate) { "OE:${originalEstimate}h" } else { "OE:none" }
            return "Saved ${elapsedHours}h to #$($Item.Id) C:$currentCompleted->$actualC R:$currentRemaining->$actualR $oeStr"
        }
        catch {
            return "Error saving #$($Item.Id): $($_.Exception.Message)"
        }
    }

    # ── Main loop ────────────────────────────────────────────────────
    [Console]::Clear()

    try {
        while ($true) {
            # Always work with a filtered list of valid items (include separator rows)
            $validItems = @($items | Where-Object { ($_.Id -and $_.Title) -or $_.IsSeparator })

            # Ensure selectedIndex is always valid for the filtered list
            if ($selectedIndex -lt 0) { $selectedIndex = 0 }
            if ($selectedIndex -ge $validItems.Count) { $selectedIndex = [Math]::Max(0, $validItems.Count - 1) }

            switch ($mode) {
                "list" {
                    $hasTimers = $activeTimers.Count -gt 0

                    $scrollOffset = Render-WorkItemList -Items $validItems `
                        -SelectedIndex $selectedIndex -ScrollOffset $scrollOffset `
                        -StatusMessage $statusMessage -ActiveTimers $activeTimers `
                        -TabNames $tabNames -ActiveTabIndex $activeTab `
                        -ShowAllEnabled $tabState[$activeTab].ShowAll `
                        -ShowAllAvailable ($activeTab -in 1, 2, 3, 5) `
                        -NoItemsMessage $tabState[$activeTab].EmptyMessage
                    $statusMessage = ""

                    # If timers are running, use non-blocking input with refresh
                    if ($hasTimers) {
                        if ([Console]::KeyAvailable) {
                            $key = [Console]::ReadKey($true)
                        }
                        else {
                            Start-Sleep -Milliseconds 250
                            continue
                        }
                    }
                    else {
                        $key = [Console]::ReadKey($true)
                    }

                    switch ($key.Key) {
                        'UpArrow' {
                            if ($selectedIndex -gt 0) {
                                $selectedIndex--
                                # Skip separator rows
                                while ($selectedIndex -gt 0 -and $validItems[$selectedIndex].IsSeparator) { $selectedIndex-- }
                            }
                        }
                        'DownArrow' {
                            if ($selectedIndex -lt ($validItems.Count - 1)) {
                                $selectedIndex++
                                # Skip separator rows
                                while ($selectedIndex -lt ($validItems.Count - 1) -and $validItems[$selectedIndex].IsSeparator) { $selectedIndex++ }
                            }
                        }
                        'PageUp' {
                            $pageSize = [Console]::WindowHeight - 4
                            $selectedIndex = [Math]::Max(0, $selectedIndex - $pageSize)
                        }
                        'PageDown' {
                            $pageSize = [Console]::WindowHeight - 4
                            $selectedIndex = [Math]::Min($validItems.Count - 1, $selectedIndex + $pageSize)
                        }
                        'Home' {
                            $selectedIndex = 0
                        }
                        'End' {
                            $selectedIndex = $validItems.Count - 1
                        }
                        'R' {
                            if ($activeTab -eq 5 -and $priData.HasParent) {
                                # Refresh Pri tab - reload children/related for saved parent
                                [Console]::Clear()
                                Write-Host "`n  Refreshing items for parent #$($priData.ParentId)..." -ForegroundColor Cyan
                                $items = [System.Collections.ArrayList]@(
                                    Get-PriWorkItems -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -ParentId $priData.ParentId `
                                        -IncludeClosed:$tabState[5].ShowAll
                                )
                                $tabState[5].Items = $items
                                $tabState[5].Loaded = $true
                                $selectedIndex = 0
                                $scrollOffset = 0
                                $statusMessage = "Refreshed - $($items.Count) items"
                                [Console]::Clear()
                            }
                            elseif ($activeTab -eq 5) {
                                # No parent set, open priform
                                $mode = "priform"
                                [Console]::Clear()
                            }
                            elseif ($activeTab -eq 4 -and $queryData.HasSearched) {
                                # Re-run last search on Query tab
                                [Console]::Clear()
                                Write-Host "`n  Re-running search..." -ForegroundColor Cyan
                                $wiId = 0
                                if ($queryData.WorkItemId -match '^\d+$') { $wiId = [int]$queryData.WorkItemId }
                                $results = @(Search-WorkItemsByFilters `
                                    -Organization $config.Organization `
                                    -Project $config.Project -PAT $config.PAT `
                                    -TitleContains $queryData.TitleContains `
                                    -StateFilter $queryData.State `
                                    -TypeFilter $queryData.Type `
                                    -AssignedTo $queryData.AssignedTo `
                                    -WorkItemId $wiId)
                                $items = [System.Collections.ArrayList]@($results)
                                $tabState[4].Items = $items
                                $selectedIndex = 0
                                $scrollOffset = 0
                                $statusMessage = "Found $($items.Count) item(s)"
                                [Console]::Clear()
                            }
                            elseif ($activeTab -eq 4) {
                                # No previous search, open query form
                                $mode = "queryform"
                                [Console]::Clear()
                            }
                            else {
                                $statusMessage = "Refreshing..."
                                $prevSelectedId = if ($validItems.Count -gt 0) { $validItems[$selectedIndex].Id } else { $null }
                                $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                $tabState[$activeTab].Items = $items
                                $tabState[$activeTab].Loaded = $true
                                $validItems = @($items | Where-Object { ($_.Id -and $_.Title) -or $_.IsSeparator })
                                $selectedIndex = 0
                                if ($prevSelectedId) {
                                    for ($i = 0; $i -lt $validItems.Count; $i++) {
                                        if ($validItems[$i].Id -eq $prevSelectedId) {
                                            $selectedIndex = $i
                                            break
                                        }
                                    }
                                }
                                $scrollOffset = 0
                                $statusMessage = "Refreshed - $($validItems.Count) items loaded"
                                [Console]::Clear()
                            }
                        }
                        'Enter' {
                            # Show detail view
                            if ($validItems.Count -gt 0) {
                                $item = $validItems[$selectedIndex]
                                if ($item.IsSeparator) { break }  # Don't open separator items
                                [Console]::Clear()

                                $detail = Get-WorkItemDetail -Organization $config.Organization `
                                    -Project $config.Project -PAT $config.PAT `
                                    -WorkItemId $item.Id

                                $comments = @(Get-WorkItemComments -Organization $config.Organization `
                                    -Project $config.Project -PAT $config.PAT `
                                    -WorkItemId $item.Id)

                                $description = ""
                                $reproSteps = ""
                                $systemInfo = ""
                                if ($detail -and $detail.fields) {
                                    $descProp = $detail.fields.PSObject.Properties['System.Description']
                                    if ($descProp -and $descProp.Value) {
                                        $description = $descProp.Value
                                    }
                                    $reproProp = $detail.fields.PSObject.Properties['Microsoft.VSTS.TCM.ReproSteps']
                                    if ($reproProp -and $reproProp.Value) {
                                        $reproSteps = $reproProp.Value
                                    }
                                    $sysInfoProp = $detail.fields.PSObject.Properties['Microsoft.VSTS.TCM.SystemInfo']
                                    if ($sysInfoProp -and $sysInfoProp.Value) {
                                        $systemInfo = $sysInfoProp.Value
                                    }
                                }

                                $detailData = @{
                                    Item        = $item
                                    Description = $description
                                    ReproSteps  = $reproSteps
                                    SystemInfo  = $systemInfo
                                    Comments    = $comments
                                }
                                $detailScrollOffset = 0
                                $mode = "detail"
                                [Console]::Clear()
                            }
                        }
                        'T' {
                            # Toggle time tracking on selected item
                            if ($validItems.Count -gt 0) {
                                $item = $validItems[$selectedIndex]
                                if ($item.IsSeparator) { break }  # Can't track time on separator
                                $itemId = $item.Id

                                if ($activeTimers.ContainsKey($itemId)) {
                                    # Stop timer and save
                                    $msg = Save-Timer -Item $item -Stopwatch $activeTimers[$itemId] -Config $config
                                    $activeTimers.Remove($itemId)
                                    $statusMessage = $msg
                                }
                                else {
                                    # Check if item supports time tracking
                                    $supportsTime = ($null -ne $item.OriginalEstimate) -or
                                                    ($null -ne $item.CompletedWork) -or
                                                    ($null -ne $item.RemainingWork)

                                    if (-not $supportsTime -and $item.Type -eq 'User Story') {
                                        # Create a child Task for the User Story
                                        $taskTitle = "$($item.Id) $($item.Title)"
                                        $assignedTo = Get-SafeField -Fields $item.Raw.fields -Name 'System.AssignedTo'
                                        $assignedToValue = if ($assignedTo -and $assignedTo.uniqueName) { $assignedTo.uniqueName } else { "" }
                                        $statusMessage = "Creating child task for #$itemId..."
                                        try {
                                            $newWI = New-ChildTask -Organization $config.Organization `
                                                -Project $config.Project -PAT $config.PAT `
                                                -ParentId $itemId -Title $taskTitle `
                                                -AssignedTo $assignedToValue `
                                                -OriginalEstimate 5 -RemainingWork 5

                                            $newId = $newWI.id
                                            # Add the new task to the item list right after the parent
                                            $newTaskNode = @{
                                                Id               = $newId
                                                Title            = $taskTitle
                                                Type             = 'Task'
                                                State            = $newWI.fields.'System.State'
                                                AssignedTo       = $item.AssignedTo
                                                ParentId         = $itemId
                                                OriginalEstimate = 5.0
                                                CompletedWork    = $null
                                                RemainingWork    = 5.0
                                                IsMine           = $true
                                                IsSeparator      = $false
                                                IsRelated        = $false
                                                Depth            = $item.Depth + 1
                                                Children         = @()
                                                Raw              = $newWI
                                            }
                                            # Insert after current item
                                            $insertIdx = $items.IndexOf($item)
                                            if ($insertIdx -ge 0) {
                                                $items.Insert($insertIdx + 1, $newTaskNode)
                                                $selectedIndex = $insertIdx + 1
                                            } else {
                                                [void]$items.Add($newTaskNode)
                                                $selectedIndex = $items.Count - 1
                                            }
                                            # Start timer on the new task
                                            $activeTimers[$newId] = [System.Diagnostics.Stopwatch]::StartNew()
                                            $statusMessage = "Created task #$newId and started timer"
                                        }
                                        catch {
                                            $statusMessage = "Error creating task: $($_.Exception.Message)"
                                        }
                                    }
                                    elseif (-not $supportsTime) {
                                        $statusMessage = "Item #$itemId has no time tracking fields"
                                    }
                                    else {
                                        # If the item is closed/done, reactivate it and reset remaining work
                                        $closedStates = @('Closed', 'Done', 'Released', 'Removed')
                                        if ($closedStates -contains $item.State) {
                                            $reactivateMsg = ""
                                            try {
                                                # Pick the best active state for this work item type
                                                $states = @(Get-WorkItemTypeStates -Organization $config.Organization `
                                                    -Project $config.Project -PAT $config.PAT `
                                                    -WorkItemType $item.Type)

                                                $preferredActive = @('Active', 'In Progress', 'Committed', 'In Development')
                                                $targetState = $null
                                                foreach ($ps in $preferredActive) {
                                                    if ($states -contains $ps) { $targetState = $ps; break }
                                                }
                                                if (-not $targetState) {
                                                    # Fall back: first state that isn't in the closed set and isn't 'New'/'Proposed'
                                                    $skipStates = $closedStates + @('New', 'Proposed', 'Removed')
                                                    $targetState = $states | Where-Object { $skipStates -notcontains $_ } | Select-Object -First 1
                                                }
                                                if (-not $targetState -and $states.Count -gt 0) {
                                                    $targetState = $states[0]
                                                }

                                                if ($targetState) {
                                                    Update-WorkItemState -Organization $config.Organization `
                                                        -Project $config.Project -PAT $config.PAT `
                                                        -WorkItemId $itemId -NewState $targetState
                                                    $item['State'] = $targetState
                                                    $reactivateMsg = "Reactivated #$itemId to '$targetState'"
                                                }
                                            }
                                            catch {
                                                $reactivateMsg = "Warning: could not reactivate #$itemId ($($_.Exception.Message))"
                                            }

                                            # Reset remaining work = max(0, OriginalEstimate - CompletedWork)
                                            try {
                                                $oe = if ($null -ne $item.OriginalEstimate) { [double]$item.OriginalEstimate } else { $null }
                                                $cw = if ($null -ne $item.CompletedWork)    { [double]$item.CompletedWork    } else { 0.0 }
                                                if ($null -ne $oe) {
                                                    $newRemaining = [Math]::Max(0.0, $oe - $cw)
                                                    Update-WorkItemTime -Organization $config.Organization `
                                                        -Project $config.Project -PAT $config.PAT `
                                                        -WorkItemId $itemId `
                                                        -CompletedWork $cw `
                                                        -RemainingWork $newRemaining | Out-Null
                                                    $item['RemainingWork'] = $newRemaining
                                                    $reactivateMsg += " | Remaining set to $([Math]::Round($newRemaining,2))h (OE:${oe}h - C:${cw}h)"
                                                }
                                            }
                                            catch {
                                                $reactivateMsg += " | Warning: could not reset remaining work ($($_.Exception.Message))"
                                            }

                                            $statusMessage = "$reactivateMsg | Timer started"
                                        }
                                        else {
                                            $statusMessage = "Started timer on #$itemId"
                                        }

                                        $activeTimers[$itemId] = [System.Diagnostics.Stopwatch]::StartNew()
                                    }
                                }
                            }
                        }
                        'M' {
                            # Open Tools menu
                            $toolsMenuData = @{
                                SelectedIndex = 0
                                MenuItems     = @(
                                    @{ Label = "Reconfigure (Organization, Project, PAT)"; Action = "reconfigure" }
                                    @{ Label = "Delete selected Task";                    Action = "deletetask"  }
                                    @{ Label = "View debug log";                          Action = "viewlog"     }
                                    @{ Label = "View README";                             Action = "readme"      }
                                    @{ Label = "About";                                  Action = "about"       }
                                )
                            }
                            $mode = "toolsmenu"
                            [Console]::Clear()
                        }
                        'Q' {
                            # Save all active timers before quitting
                            if ($activeTimers.Count -gt 0) {
                                [Console]::Clear()
                                Write-Host ""
                                Write-Host "  Saving all active timers..." -ForegroundColor Yellow
                                foreach ($timerId in @($activeTimers.Keys)) {
                                    $timerItem = $validItems | Where-Object { $_.Id -eq $timerId } | Select-Object -First 1
                                    if ($timerItem) {
                                        $msg = Save-Timer -Item $timerItem -Stopwatch $activeTimers[$timerId] -Config $config
                                        Write-Host "  $msg" -ForegroundColor Gray
                                    }
                                }
                                $activeTimers.Clear()
                                Start-Sleep -Seconds 2
                            }
                            [Console]::Clear()
                            [Console]::CursorVisible = $true
                            Write-Host "Goodbye!" -ForegroundColor Cyan
                            return
                        }
                        'X' {
                            # Toggle show all (closed/removed) for applicable tabs
                            if ($activeTab -in 1, 2, 3) {
                                $tabState[$activeTab].ShowAll = -not $tabState[$activeTab].ShowAll
                                $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll)
                                $tabState[$activeTab].Items = $items
                                $tabState[$activeTab].Loaded = $true
                                $selectedIndex = 0
                                $scrollOffset = 0
                                $statusMessage = if ($tabState[$activeTab].ShowAll) { "Showing all items (including closed)" } else { "Showing active items only" }
                                [Console]::Clear()
                            }
                            elseif ($activeTab -eq 5 -and $priData.HasParent) {
                                $tabState[5].ShowAll = -not $tabState[5].ShowAll
                                $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex 5 -ShowAll $tabState[5].ShowAll -PriData $priData)
                                $tabState[5].Items = $items
                                $tabState[5].Loaded = $true
                                $selectedIndex = 0
                                $scrollOffset = 0
                                $statusMessage = if ($tabState[5].ShowAll) { "Showing all items (including closed)" } else { "Showing active items only" }
                                [Console]::Clear()
                            }
                        }
                        'Tab' {
                            # Tab / Shift+Tab to cycle through tabs
                            if ($key.Modifiers -band [ConsoleModifiers]::Shift) {
                                $newTab = if ($activeTab -eq 0) { $tabNames.Count - 1 } else { $activeTab - 1 }
                            }
                            else {
                                $newTab = if ($activeTab -eq ($tabNames.Count - 1)) { 0 } else { $activeTab + 1 }
                            }
                            # Save current tab state
                            $tabState[$activeTab].SelectedIndex = $selectedIndex
                            $tabState[$activeTab].ScrollOffset = $scrollOffset
                            $tabState[$activeTab].Items = $items

                            $activeTab = $newTab

                            if (-not $tabState[$activeTab].Loaded) {
                                $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                $tabState[$activeTab].Items = $items
                                $tabState[$activeTab].Loaded = $true
                            }
                            else {
                                $items = $tabState[$activeTab].Items
                                if (-not $items) { $items = [System.Collections.ArrayList]::new() }
                            }

                            $selectedIndex = $tabState[$activeTab].SelectedIndex
                            $scrollOffset = $tabState[$activeTab].ScrollOffset

                            if ($activeTab -eq 4 -and (-not $queryData.HasSearched)) {
                                $mode = "queryform"
                            }
                            elseif ($activeTab -eq 5 -and (-not $priData.HasParent)) {
                                $mode = "priform"
                            }

                            [Console]::Clear()
                        }
                        default {
                            # Check for tab switching (digit keys 1-6) and '/' for query/pri
                            $keyChar = $key.KeyChar
                            if ($keyChar -ge '1' -and $keyChar -le '6') {
                                $newTab = [int]::Parse($keyChar.ToString()) - 1
                                if ($newTab -ne $activeTab -and $newTab -lt $tabNames.Count) {
                                    # Save current tab state
                                    $tabState[$activeTab].SelectedIndex = $selectedIndex
                                    $tabState[$activeTab].ScrollOffset = $scrollOffset
                                    $tabState[$activeTab].Items = $items

                                    $activeTab = $newTab

                                    # Load tab if needed
                                    if (-not $tabState[$activeTab].Loaded) {
                                        $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                        $tabState[$activeTab].Items = $items
                                        $tabState[$activeTab].Loaded = $true
                                    }
                                    else {
                                        $items = $tabState[$activeTab].Items
                                        if (-not $items) { $items = [System.Collections.ArrayList]::new() }
                                    }

                                    $selectedIndex = $tabState[$activeTab].SelectedIndex
                                    $scrollOffset = $tabState[$activeTab].ScrollOffset

                                    # Auto-open form for Query/Pri tabs with no data yet
                                    if ($activeTab -eq 4 -and (-not $queryData.HasSearched)) {
                                        $mode = "queryform"
                                    }
                                    elseif ($activeTab -eq 5 -and (-not $priData.HasParent)) {
                                        $mode = "priform"
                                    }

                                    [Console]::Clear()
                                }
                            }
                            elseif ($keyChar -eq '/' -and $activeTab -eq 4) {
                                $mode = "queryform"
                                [Console]::Clear()
                            }
                            elseif ($keyChar -eq '/' -and $activeTab -eq 5) {
                                $priData.FormState = 'input'
                                $mode = "priform"
                                [Console]::Clear()
                            }
                        }
                    }
                }

                "detail" {
                    $renderResult = Render-DetailView -Item $detailData.Item `
                        -Description $detailData.Description `
                        -ReproSteps $detailData.ReproSteps `
                        -SystemInfo $detailData.SystemInfo `
                        -Comments $detailData.Comments `
                        -ScrollOffset $detailScrollOffset `
                        -Organization $config.Organization `
                        -Project $config.Project

                    $detailScrollOffset = $renderResult.ScrollOffset
                    $totalDetailLines = $renderResult.TotalLines

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $mode = "list"
                            $detailData = $null
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($detailScrollOffset -gt 0) { $detailScrollOffset-- }
                        }
                        'DownArrow' {
                            $maxScroll = [Math]::Max(0, $totalDetailLines - ([Console]::WindowHeight - 3))
                            if ($detailScrollOffset -lt $maxScroll) { $detailScrollOffset++ }
                        }
                        'PageUp' {
                            $detailScrollOffset = [Math]::Max(0, $detailScrollOffset - ([Console]::WindowHeight - 4))
                        }
                        'PageDown' {
                            $maxScroll = [Math]::Max(0, $totalDetailLines - ([Console]::WindowHeight - 3))
                            $detailScrollOffset = [Math]::Min($maxScroll, $detailScrollOffset + ([Console]::WindowHeight - 4))
                        }
                        'S' {
                            # Open status picker
                            $item = $detailData.Item
                            $states = @(Get-WorkItemTypeStates -Organization $config.Organization `
                                -Project $config.Project -PAT $config.PAT `
                                -WorkItemType $item.Type)

                            if ($states.Count -gt 0) {
                                $statusPickerData = @{
                                    Item          = $item
                                    Statuses      = $states
                                    SelectedIndex = 0
                                }
                                for ($si = 0; $si -lt $states.Count; $si++) {
                                    if ($states[$si] -eq $item.State) {
                                        $statusPickerData.SelectedIndex = $si
                                        break
                                    }
                                }
                                $mode = "statuspicker"
                                [Console]::Clear()
                            }
                        }
                        'A' {
                            # Add a new comment using full-screen editor
                            $newText = Edit-TextBlock -Title "Add Comment to #$($detailData.Item.Id)" -InitialText ""

                            [Console]::CursorVisible = $false

                            if ($null -ne $newText -and $newText.Trim().Length -gt 0) {
                                try {
                                    Add-WorkItemComment -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -WorkItemId $detailData.Item.Id -Text $newText

                                    # Refresh comments
                                    $detailData.Comments = @(Get-WorkItemComments -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -WorkItemId $detailData.Item.Id)
                                    $statusMessage = "Comment added"
                                }
                                catch {
                                    $statusMessage = "Error: $($_.Exception.Message)"
                                }
                            }
                            else {
                                $statusMessage = "Comment cancelled"
                            }
                            [Console]::Clear()
                        }
                        'E' {
                            # Edit a comment - open picker
                            $comments = @($detailData.Comments)
                            if ($comments.Count -eq 0) {
                                $statusMessage = "No comments to edit"
                            }
                            else {
                                $commentPickerData = @{
                                    Item          = $detailData.Item
                                    Comments      = $comments
                                    SelectedIndex = 0
                                    Action        = "edit"
                                }
                                $mode = "commentpicker"
                                [Console]::Clear()
                            }
                        }
                        'D' {
                            # Delete a comment - open picker
                            $comments = @($detailData.Comments)
                            if ($comments.Count -eq 0) {
                                $statusMessage = "No comments to delete"
                            }
                            else {
                                $commentPickerData = @{
                                    Item          = $detailData.Item
                                    Comments      = $comments
                                    SelectedIndex = 0
                                    Action        = "delete"
                                }
                                $mode = "commentpicker"
                                [Console]::Clear()
                            }
                        }
                        'H' {
                            # Edit hours (Original Estimate, Completed, Remaining)
                            $item = $detailData.Item
                            $supportsTime = ($null -ne $item.OriginalEstimate) -or
                                            ($null -ne $item.CompletedWork) -or
                                            ($null -ne $item.RemainingWork)

                            if (-not $supportsTime) {
                                $statusMessage = "This item has no time tracking fields"
                            }
                            else {
                                [Console]::Clear()
                                [Console]::CursorVisible = $true
                                Write-Host ""
                                Write-Host "  ── Edit Hours for #$($item.Id): $($item.Title) ──" -ForegroundColor Cyan
                                Write-Host ""

                                $curOE = if ($null -ne $item.OriginalEstimate) { $item.OriginalEstimate } else { "" }
                                $curCW = if ($null -ne $item.CompletedWork) { $item.CompletedWork } else { "" }
                                $curRW = if ($null -ne $item.RemainingWork) { $item.RemainingWork } else { "" }

                                Write-Host "  Enter new values ('c' to clear, Enter to keep default):" -ForegroundColor Gray
                                Write-Host ""
                                $inputOE = Read-Host "  Original Estimate [$curOE]"
                                if ($inputOE -eq '') { $inputOE = [string]$curOE }

                                $inputCW = Read-Host "  Completed Work    [$curCW]"
                                if ($inputCW -eq '') { $inputCW = [string]$curCW }

                                $inputRW = Read-Host "  Remaining Work    [$curRW]"
                                if ($inputRW -eq '') { $inputRW = [string]$curRW }

                                [Console]::CursorVisible = $false

                                $fieldsToUpdate = @{}
                                $changed = @()

                                # Process Original Estimate
                                if ($inputOE -eq 'c') {
                                    $fieldsToUpdate['Microsoft.VSTS.Scheduling.OriginalEstimate'] = $null
                                    $changed += 'OE cleared'
                                }
                                elseif ($inputOE -ne '' -and $inputOE -ne [string]$curOE) {
                                    $parsedOE = 0.0
                                    if ([double]::TryParse($inputOE, [ref]$parsedOE)) {
                                        $fieldsToUpdate['Microsoft.VSTS.Scheduling.OriginalEstimate'] = $parsedOE
                                        $changed += "OE=$parsedOE"
                                    }
                                    else {
                                        $statusMessage = "Invalid number for Original Estimate: $inputOE"
                                        [Console]::Clear()
                                        continue
                                    }
                                }

                                # Process Completed Work
                                if ($inputCW -eq 'c') {
                                    $fieldsToUpdate['Microsoft.VSTS.Scheduling.CompletedWork'] = $null
                                    $changed += 'CW cleared'
                                }
                                elseif ($inputCW -ne '' -and $inputCW -ne [string]$curCW) {
                                    $parsedCW = 0.0
                                    if ([double]::TryParse($inputCW, [ref]$parsedCW)) {
                                        $fieldsToUpdate['Microsoft.VSTS.Scheduling.CompletedWork'] = $parsedCW
                                        $changed += "CW=$parsedCW"
                                    }
                                    else {
                                        $statusMessage = "Invalid number for Completed Work: $inputCW"
                                        [Console]::Clear()
                                        continue
                                    }
                                }

                                # Process Remaining Work
                                if ($inputRW -eq 'c') {
                                    $fieldsToUpdate['Microsoft.VSTS.Scheduling.RemainingWork'] = $null
                                    $changed += 'RW cleared'
                                }
                                elseif ($inputRW -ne '' -and $inputRW -ne [string]$curRW) {
                                    $parsedRW = 0.0
                                    if ([double]::TryParse($inputRW, [ref]$parsedRW)) {
                                        $fieldsToUpdate['Microsoft.VSTS.Scheduling.RemainingWork'] = $parsedRW
                                        $changed += "RW=$parsedRW"
                                    }
                                    else {
                                        $statusMessage = "Invalid number for Remaining Work: $inputRW"
                                        [Console]::Clear()
                                        continue
                                    }
                                }

                                if ($fieldsToUpdate.Count -gt 0) {
                                    try {
                                        $apiResult = Update-WorkItemHours -Organization $config.Organization `
                                            -Project $config.Project -PAT $config.PAT `
                                            -WorkItemId $item.Id -Fields $fieldsToUpdate

                                        # Update local item from API response
                                        if ($null -ne $apiResult -and $null -ne $apiResult.fields) {
                                            $rf = $apiResult.fields
                                            $item['OriginalEstimate'] = Get-SafeField -Fields $rf -Name 'Microsoft.VSTS.Scheduling.OriginalEstimate'
                                            $item['CompletedWork'] = Get-SafeField -Fields $rf -Name 'Microsoft.VSTS.Scheduling.CompletedWork'
                                            $item['RemainingWork'] = Get-SafeField -Fields $rf -Name 'Microsoft.VSTS.Scheduling.RemainingWork'
                                        }

                                        $statusMessage = "Hours updated: $($changed -join ', ')"
                                    }
                                    catch {
                                        $statusMessage = "Error: $($_.Exception.Message)"
                                    }
                                }
                                else {
                                    $statusMessage = "No changes made"
                                }
                                [Console]::Clear()
                            }
                        }
                        'F' {
                            # Edit fields - build list of editable fields
                            $item = $detailData.Item
                            $editableFields = [System.Collections.ArrayList]::new()

                            # Title
                            $titlePreview = if ($item.Title) { $item.Title } else { "(empty)" }
                            [void]$editableFields.Add(@{
                                Label     = "Title"
                                FieldPath = "System.Title"
                                Value     = $(if ($item.Title) { $item.Title } else { "" })
                                Preview   = $titlePreview
                            })

                            # Description (not available on Bug/Incident)
                            if ($item.Type -ne "Bug" -and $item.Type -ne "Incident") {
                                $descPreview = if ($detailData.Description) {
                                    $dp = Remove-Html -Html $detailData.Description
                                    if ($dp.Length -gt 50) { $dp.Substring(0, 47) + "..." } else { $dp }
                                } else { "(empty)" }
                                [void]$editableFields.Add(@{
                                    Label     = "Description"
                                    FieldPath = "System.Description"
                                    Value     = $(if ($detailData.Description) { $detailData.Description } else { "" })
                                    Preview   = $descPreview
                                })
                            }

                            # Repro Steps & System Info (Bugs and Incidents)
                            if ($item.Type -eq "Bug" -or $item.Type -eq "Incident") {
                                $reproPreview = if ($detailData.ReproSteps) {
                                    $rp = Remove-Html -Html $detailData.ReproSteps
                                    if ($rp.Length -gt 50) { $rp.Substring(0, 47) + "..." } else { $rp }
                                } else { "(empty)" }
                                [void]$editableFields.Add(@{
                                    Label     = "Repro Steps"
                                    FieldPath = "Microsoft.VSTS.TCM.ReproSteps"
                                    Value     = $(if ($detailData.ReproSteps) { $detailData.ReproSteps } else { "" })
                                    Preview   = $reproPreview
                                })

                                $sysPreview = if ($detailData.SystemInfo) {
                                    $sp = Remove-Html -Html $detailData.SystemInfo
                                    if ($sp.Length -gt 50) { $sp.Substring(0, 47) + "..." } else { $sp }
                                } else { "(empty)" }
                                [void]$editableFields.Add(@{
                                    Label     = "System Info"
                                    FieldPath = "Microsoft.VSTS.TCM.SystemInfo"
                                    Value     = $(if ($detailData.SystemInfo) { $detailData.SystemInfo } else { "" })
                                    Preview   = $sysPreview
                                })
                            }

                            $fieldPickerData = @{
                                Item          = $item
                                Fields        = $editableFields
                                SelectedIndex = 0
                            }
                            $mode = "fieldpicker"
                            [Console]::Clear()
                        }
                        'N' {
                            # Open assignee picker instantly; results load as the user types
                            $item = $detailData.Item
                            $assigneePickerData = @{
                                Item           = $item
                                SearchText     = ""
                                Results        = @()
                                SelectedIndex  = 0
                                LastSearchText = ""
                            }
                            $mode = "assigneepicker"
                            [Console]::Clear()
                        }
                        default { }
                    }
                }

                "assigneepicker" {
                    Render-AssigneePicker -Item $assigneePickerData.Item `
                        -SearchText $assigneePickerData.SearchText `
                        -Results $assigneePickerData.Results `
                        -SelectedIndex $assigneePickerData.SelectedIndex

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $assigneePickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($assigneePickerData.SelectedIndex -gt 0) {
                                $assigneePickerData.SelectedIndex--
                            }
                        }
                        'DownArrow' {
                            $maxIdx = $assigneePickerData.Results.Count  # 0=Unassign, 1..Count=results
                            if ($assigneePickerData.SelectedIndex -lt $maxIdx) {
                                $assigneePickerData.SelectedIndex++
                            }
                        }
                        'Backspace' {
                            if ($assigneePickerData.SearchText.Length -gt 0) {
                                $assigneePickerData.SearchText = $assigneePickerData.SearchText.Substring(
                                    0, $assigneePickerData.SearchText.Length - 1)
                                $assigneePickerData.SelectedIndex = 0
                            }
                        }
                        'Enter' {
                            $item = $assigneePickerData.Item
                            $selIdx = $assigneePickerData.SelectedIndex

                            if ($selIdx -eq 0) {
                                # Unassign
                                try {
                                    Set-WorkItemAssignee -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -WorkItemId $item.Id -AssignedTo ""
                                    $item['AssignedTo'] = $null
                                    $statusMessage = "Assignee cleared for #$($item.Id)"
                                }
                                catch {
                                    $statusMessage = "Error: $($_.Exception.Message)"
                                }
                            }
                            else {
                                # Assign to selected user
                                $selectedUser = $assigneePickerData.Results[$selIdx - 1]
                                try {
                                    Set-WorkItemAssignee -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -WorkItemId $item.Id -AssignedTo $selectedUser
                                    $item['AssignedTo'] = $selectedUser
                                    $statusMessage = "#$($item.Id) assigned to '$selectedUser'"
                                }
                                catch {
                                    $statusMessage = "Error: $($_.Exception.Message)"
                                }
                            }

                            $assigneePickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        default {
                            if ($key.KeyChar -and -not [char]::IsControl($key.KeyChar)) {
                                $assigneePickerData.SearchText += $key.KeyChar
                                $assigneePickerData.SelectedIndex = 0
                            }
                        }
                    }

                    # Search API when text changes and meets minimum length (2+ chars)
                    if ($null -ne $assigneePickerData) {
                        $st = $assigneePickerData.SearchText
                        if ($st.Length -ge 2 -and $st -ne $assigneePickerData.LastSearchText) {
                            $assigneePickerData.Results = @(Search-AzDoUsers `
                                -Organization $config.Organization `
                                -PAT $config.PAT `
                                -SearchTerm $st `
                                -Project $config.Project)
                            $assigneePickerData.LastSearchText = $st
                            # Clamp SelectedIndex (0=Unassign, 1..Count=results)
                            $maxIdx = $assigneePickerData.Results.Count
                            if ($assigneePickerData.SelectedIndex -gt $maxIdx) {
                                $assigneePickerData.SelectedIndex = $maxIdx
                            }
                        }
                        elseif ($st.Length -lt 2 -and $assigneePickerData.Results.Count -gt 0) {
                            $assigneePickerData.Results = @()
                            $assigneePickerData.LastSearchText = ""
                            $assigneePickerData.SelectedIndex = 0
                        }
                    }
                }

                "commentpicker" {
                    Render-CommentPicker -Item $commentPickerData.Item `
                        -Comments $commentPickerData.Comments `
                        -SelectedIndex $commentPickerData.SelectedIndex `
                        -Action $commentPickerData.Action

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $commentPickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($commentPickerData.SelectedIndex -gt 0) {
                                $commentPickerData.SelectedIndex--
                            }
                        }
                        'DownArrow' {
                            if ($commentPickerData.SelectedIndex -lt ($commentPickerData.Comments.Count - 1)) {
                                $commentPickerData.SelectedIndex++
                            }
                        }
                        'Enter' {
                            $selectedComment = $commentPickerData.Comments[$commentPickerData.SelectedIndex]
                            $action = $commentPickerData.Action

                            if ($action -eq "edit") {
                                # Open full-screen text editor with original text
                                $oldText = Remove-Html -Html $selectedComment.text
                                $author = "Unknown"
                                if ($selectedComment.createdBy -and $selectedComment.createdBy.displayName) {
                                    $author = $selectedComment.createdBy.displayName
                                }

                                $newText = Edit-TextBlock -Title "Edit Comment by $author" -InitialText $oldText

                                [Console]::CursorVisible = $false

                                if ($null -ne $newText) {
                                    try {
                                        Update-WorkItemComment -Organization $config.Organization `
                                            -Project $config.Project -PAT $config.PAT `
                                            -WorkItemId $commentPickerData.Item.Id `
                                            -CommentId $selectedComment.id -Text $newText
                                        $detailData.Comments = @(Get-WorkItemComments -Organization $config.Organization `
                                            -Project $config.Project -PAT $config.PAT `
                                            -WorkItemId $commentPickerData.Item.Id)
                                        $statusMessage = "Comment updated"
                                    }
                                    catch {
                                        $statusMessage = "Error: $($_.Exception.Message)"
                                    }
                                }
                                else {
                                    $statusMessage = "Edit cancelled"
                                }
                            }
                            elseif ($action -eq "delete") {
                                # Confirm deletion
                                [Console]::Clear()
                                Write-Host ""
                                $preview = (Remove-Html -Html $selectedComment.text)
                                if ($preview.Length -gt 80) { $preview = $preview.Substring(0, 77) + "..." }
                                Write-Host "  Delete this comment?" -ForegroundColor Yellow
                                Write-Host "  $preview" -ForegroundColor Gray
                                Write-Host ""
                                Write-Host "  Press [y] to confirm, any other key to cancel" -ForegroundColor Yellow
                                $confirm = [Console]::ReadKey($true)
                                if ($confirm.KeyChar -eq 'y' -or $confirm.KeyChar -eq 'Y') {
                                    try {
                                        Remove-WorkItemComment -Organization $config.Organization `
                                            -Project $config.Project -PAT $config.PAT `
                                            -WorkItemId $commentPickerData.Item.Id `
                                            -CommentId $selectedComment.id
                                        $detailData.Comments = @(Get-WorkItemComments -Organization $config.Organization `
                                            -Project $config.Project -PAT $config.PAT `
                                            -WorkItemId $commentPickerData.Item.Id)
                                        $statusMessage = "Comment deleted"
                                    }
                                    catch {
                                        $statusMessage = "Error: $($_.Exception.Message)"
                                    }
                                }
                                else {
                                    $statusMessage = "Delete cancelled"
                                }
                            }

                            $commentPickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        default { }
                    }
                }

                "fieldpicker" {
                    Render-FieldPicker -Item $fieldPickerData.Item `
                        -Fields $fieldPickerData.Fields `
                        -SelectedIndex $fieldPickerData.SelectedIndex

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $fieldPickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($fieldPickerData.SelectedIndex -gt 0) {
                                $fieldPickerData.SelectedIndex--
                            }
                        }
                        'DownArrow' {
                            if ($fieldPickerData.SelectedIndex -lt ($fieldPickerData.Fields.Count - 1)) {
                                $fieldPickerData.SelectedIndex++
                            }
                        }
                        'Enter' {
                            $field = $fieldPickerData.Fields[$fieldPickerData.SelectedIndex]
                            $plainText = Remove-Html -Html $field.Value

                            $newText = Edit-TextBlock -Title "Edit $($field.Label)" -InitialText $plainText

                            [Console]::CursorVisible = $false

                            if ($null -ne $newText) {
                                try {
                                    Update-WorkItemField -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -WorkItemId $fieldPickerData.Item.Id `
                                        -FieldPath $field.FieldPath `
                                        -Value $newText

                                    # Update local detail data
                                    switch ($field.FieldPath) {
                                        "System.Title" {
                                            $fieldPickerData.Item['Title'] = $newText
                                            $detailData.Item['Title'] = $newText
                                        }
                                        "System.Description" {
                                            $detailData.Description = $newText
                                        }
                                        "Microsoft.VSTS.TCM.ReproSteps" {
                                            $detailData.ReproSteps = $newText
                                        }
                                        "Microsoft.VSTS.TCM.SystemInfo" {
                                            $detailData.SystemInfo = $newText
                                        }
                                    }
                                    $statusMessage = "$($field.Label) updated"
                                }
                                catch {
                                    $statusMessage = "Error: $($_.Exception.Message)"
                                }
                            }
                            else {
                                $statusMessage = "Edit cancelled"
                            }

                            $fieldPickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        default { }
                    }
                }

                "toolsmenu" {
                    Render-ToolsMenu -SelectedIndex $toolsMenuData.SelectedIndex `
                        -MenuItems $toolsMenuData.MenuItems

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $toolsMenuData = $null
                            $mode = "list"
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($toolsMenuData.SelectedIndex -gt 0) {
                                $toolsMenuData.SelectedIndex--
                            }
                        }
                        'DownArrow' {
                            if ($toolsMenuData.SelectedIndex -lt ($toolsMenuData.MenuItems.Count - 1)) {
                                $toolsMenuData.SelectedIndex++
                            }
                        }
                        'Enter' {
                            $selectedAction = $toolsMenuData.MenuItems[$toolsMenuData.SelectedIndex].Action

                            switch ($selectedAction) {
                                "reconfigure" {
                                    [Console]::Clear()
                                    [Console]::CursorVisible = $true
                                    $config = Request-TTConfig
                                    [Console]::CursorVisible = $false

                                    # Reset all tabs and reload current
                                    for ($ti = 0; $ti -lt $tabState.Count; $ti++) {
                                        $tabState[$ti].Items = $null
                                        $tabState[$ti].Loaded = $false
                                        $tabState[$ti].SelectedIndex = 0
                                        $tabState[$ti].ScrollOffset = 0
                                    }
                                    $queryData.HasSearched = $false
                                    $priData.HasParent  = $false
                                    $priData.ParentId   = 0
                                    $priData.ParentTitle = ""
                                    $priData.SearchInput = ""
                                    $priData.SearchResults = @()
                                    $priData.FormState = "input"
                                    $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                    $tabState[$activeTab].Items = $items
                                    $tabState[$activeTab].Loaded = $true
                                    $validItems = @($items | Where-Object { $_.Id -and $_.Title })
                                    $selectedIndex = 0
                                    $scrollOffset = 0
                                    $statusMessage = "Reconfigured - $($validItems.Count) items loaded"
                                }
                                "viewlog" {
                                    $logPath = Join-Path (Get-TTConfigDir) "debug.log"
                                    if (Test-Path $logPath) {
                                        [Console]::Clear()
                                        [Console]::CursorVisible = $true
                                        Write-Host ""
                                        Write-Host "  ── Debug Log ($logPath) ──" -ForegroundColor Cyan
                                        Write-Host ""
                                        Get-Content $logPath -Tail 50 | ForEach-Object { Write-Host "  $_" }
                                        Write-Host ""
                                        Write-Host "  Press any key to continue..." -ForegroundColor DarkGray
                                        $null = [Console]::ReadKey($true)
                                        [Console]::CursorVisible = $false
                                    }
                                    else {
                                        $statusMessage = "No debug log found yet"
                                    }
                                }
                                "readme" {
                                    $readmePath = Join-Path $PSScriptRoot '..' 'README.md'
                                    if (Test-Path $readmePath) {
                                        [Console]::Clear()
                                        [Console]::CursorVisible = $true
                                        $readmeLines = @(Get-Content $readmePath)
                                        $scrollPos = 0
                                        $pageSize = [Console]::WindowHeight - 3

                                        while ($true) {
                                            [Console]::SetCursorPosition(0, 0)
                                            $visibleEnd = [Math]::Min($scrollPos + $pageSize, $readmeLines.Count)
                                            for ($li = $scrollPos; $li -lt $visibleEnd; $li++) {
                                                $line = $readmeLines[$li]
                                                if ($line -match '^#{1,2}\s') {
                                                    Write-Host $line.PadRight([Console]::WindowWidth) -ForegroundColor Cyan
                                                }
                                                elseif ($line -match '^#{3,}\s') {
                                                    Write-Host $line.PadRight([Console]::WindowWidth) -ForegroundColor Yellow
                                                }
                                                elseif ($line -match '^```') {
                                                    Write-Host $line.PadRight([Console]::WindowWidth) -ForegroundColor DarkGray
                                                }
                                                elseif ($line -match '^\s*[-*]\s') {
                                                    Write-Host $line.PadRight([Console]::WindowWidth) -ForegroundColor White
                                                }
                                                else {
                                                    Write-Host $line.PadRight([Console]::WindowWidth) -ForegroundColor Gray
                                                }
                                            }
                                            # Clear any leftover lines
                                            for ($li = $visibleEnd; $li -lt $scrollPos + $pageSize; $li++) {
                                                Write-Host (' ' * [Console]::WindowWidth)
                                            }
                                            $pct = if ($readmeLines.Count -gt $pageSize) { [Math]::Round(($scrollPos / [Math]::Max(1, $readmeLines.Count - $pageSize)) * 100) } else { 100 }
                                            Write-Host "  README.md  Line $($scrollPos+1)/$($readmeLines.Count)  ${pct}%%  [Esc] close  [↑↓/PgUp/PgDn] scroll" -ForegroundColor DarkCyan -NoNewline

                                            $rk = [Console]::ReadKey($true)
                                            if ($rk.Key -eq 'Escape') { break }
                                            elseif ($rk.Key -eq 'UpArrow') { if ($scrollPos -gt 0) { $scrollPos-- } }
                                            elseif ($rk.Key -eq 'DownArrow') { if ($scrollPos -lt ($readmeLines.Count - $pageSize)) { $scrollPos++ } }
                                            elseif ($rk.Key -eq 'PageUp') { $scrollPos = [Math]::Max(0, $scrollPos - $pageSize) }
                                            elseif ($rk.Key -eq 'PageDown') { $scrollPos = [Math]::Min([Math]::Max(0, $readmeLines.Count - $pageSize), $scrollPos + $pageSize) }
                                            elseif ($rk.Key -eq 'Home') { $scrollPos = 0 }
                                            elseif ($rk.Key -eq 'End') { $scrollPos = [Math]::Max(0, $readmeLines.Count - $pageSize) }
                                        }
                                        [Console]::CursorVisible = $false
                                    }
                                    else {
                                        $statusMessage = "README.md not found"
                                    }
                                }
                                "deletetask" {
                                    $targetItem = if ($validItems.Count -gt 0) { $validItems[$selectedIndex] } else { $null }
                                    if ($null -eq $targetItem -or $targetItem.Type -ne 'Task') {
                                        $statusMessage = "Delete Task only works on Task items (selected item is '$($targetItem.Type)')"
                                    }
                                    else {
                                        [Console]::Clear()
                                        [Console]::CursorVisible = $true
                                        Write-Host ""
                                        Write-Host "  Delete Task" -ForegroundColor Red
                                        Write-Host ""
                                        Write-Host "  #$($targetItem.Id) $($targetItem.Title)" -ForegroundColor White
                                        Write-Host ""
                                        Write-Host "  This will move the task to the Azure DevOps recycle bin." -ForegroundColor Yellow
                                        Write-Host "  Press [y] to confirm, any other key to cancel." -ForegroundColor Yellow
                                        Write-Host ""
                                        $confirm = [Console]::ReadKey($true)
                                        [Console]::CursorVisible = $false
                                        if ($confirm.KeyChar -eq 'y' -or $confirm.KeyChar -eq 'Y') {
                                            try {
                                                Remove-WorkItem -Organization $config.Organization `
                                                    -Project $config.Project -PAT $config.PAT `
                                                    -WorkItemId $targetItem.Id
                                                # Remove from local list
                                                [void]$items.Remove($targetItem)
                                                if ($selectedIndex -ge $items.Count) {
                                                    $selectedIndex = [Math]::Max(0, $items.Count - 1)
                                                }
                                                $statusMessage = "Deleted Task #$($targetItem.Id)"
                                            }
                                            catch {
                                                $statusMessage = "Error: $($_.Exception.Message)"
                                            }
                                        }
                                        else {
                                            $statusMessage = "Delete cancelled"
                                        }
                                    }
                                }
                                "about" {
                                    $modInfo = $null
                                    $manifestPath = Join-Path $PSScriptRoot '..' 'AzDoTimeTracker.psd1'
                                    if (Test-Path $manifestPath) {
                                        $modInfo = Import-PowerShellDataFile $manifestPath
                                    }
                                    $ver = if ($modInfo) { $modInfo.ModuleVersion } else { '?' }
                                    $author = if ($modInfo) { $modInfo.Author } else { '?' }
                                    $desc = if ($modInfo) { $modInfo.Description } else { '?' }
                                    $ps = if ($modInfo) { $modInfo.PowerShellVersion } else { '?' }

                                    [Console]::Clear()
                                    [Console]::CursorVisible = $true
                                    Write-Host ""
                                    Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
                                    Write-Host "  ║           AzDoTimeTracker                           ║" -ForegroundColor Cyan
                                    Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
                                    Write-Host ""
                                    Write-Host "  Version:      $ver" -ForegroundColor White
                                    Write-Host "  Author:       $author" -ForegroundColor White
                                    Write-Host "  PowerShell:   $ps+" -ForegroundColor White
                                    Write-Host ""
                                    Write-Host "  $desc" -ForegroundColor Gray
                                    Write-Host ""
                                    Write-Host "  Track time, manage states, edit fields and comments" -ForegroundColor Gray
                                    Write-Host "  on your Azure DevOps work items from the terminal." -ForegroundColor Gray
                                    Write-Host ""
                                    Write-Host "  Press any key to continue..." -ForegroundColor DarkGray
                                    $null = [Console]::ReadKey($true)
                                    [Console]::CursorVisible = $false
                                }
                            }

                            $toolsMenuData = $null
                            $mode = "list"
                            [Console]::Clear()
                        }
                        default { }
                    }
                }

                "queryform" {
                    Render-QueryForm -QueryData $queryData `
                        -TabNames $tabNames -ActiveTabIndex $activeTab `
                        -StatusMessage $statusMessage
                    $statusMessage = ""

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $mode = "list"
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($queryData.FormSelectedIndex -gt 0) {
                                $queryData.FormSelectedIndex--
                            }
                        }
                        'DownArrow' {
                            if ($queryData.FormSelectedIndex -lt 5) {
                                $queryData.FormSelectedIndex++
                            }
                        }
                        'Enter' {
                            if ($queryData.FormSelectedIndex -eq 5) {
                                # Run search
                                [Console]::Clear()
                                Write-Host "`n  Searching..." -ForegroundColor Cyan
                                $wiId = 0
                                if ($queryData.WorkItemId -match '^\d+$') {
                                    $wiId = [int]$queryData.WorkItemId
                                }
                                $results = @(Search-WorkItemsByFilters `
                                    -Organization $config.Organization `
                                    -Project $config.Project -PAT $config.PAT `
                                    -TitleContains $queryData.TitleContains `
                                    -StateFilter $queryData.State `
                                    -TypeFilter $queryData.Type `
                                    -AssignedTo $queryData.AssignedTo `
                                    -WorkItemId $wiId)
                                $items = [System.Collections.ArrayList]@($results)
                                $tabState[4].Items = $items
                                $queryData.HasSearched = $true
                                $selectedIndex = 0
                                $scrollOffset = 0
                                $tabState[4].EmptyMessage = "No results found. Press / to search again."
                                $statusMessage = "Found $($items.Count) item(s)"
                                $mode = "list"
                                [Console]::Clear()
                            }
                            else {
                                # Edit selected field
                                $fieldKeys = @("TitleContains", "State", "Type", "AssignedTo", "WorkItemId")
                                $fieldLabels = @(
                                    "Title contains",
                                    "State (e.g. Active, New, Closed)",
                                    "Type (e.g. Bug, Task, User Story)",
                                    "Assigned to (@me or name)",
                                    "Work Item ID"
                                )
                                $fKey = $fieldKeys[$queryData.FormSelectedIndex]
                                $currentVal = $queryData[$fKey]
                                if (-not $currentVal) { $currentVal = "" }

                                if ($fKey -eq 'AssignedTo') {
                                    # Redraw form so the suggestion overlay has a clean background
                                    [Console]::Clear()
                                    Render-QueryForm -QueryData $queryData `
                                        -TabNames $tabNames -ActiveTabIndex $activeTab `
                                        -StatusMessage "Type 3+ characters to see name suggestions"
                                    $newVal = Read-WithSuggestions `
                                        -Prompt "Assigned to (@me or name)" `
                                        -InitialValue $currentVal `
                                        -Organization $config.Organization `
                                        -PAT $config.PAT
                                    # $null means Escape (keep old value); "" means cleared intentionally
                                    if ($null -ne $newVal) {
                                        $queryData[$fKey] = $newVal
                                    }
                                } else {
                                    [Console]::CursorVisible = $true
                                    [Console]::SetCursorPosition(0, [Console]::WindowHeight - 1)
                                    Write-Host (" " * [Console]::WindowWidth) -NoNewline
                                    [Console]::SetCursorPosition(0, [Console]::WindowHeight - 1)
                                    $newVal = Read-Host " $($fieldLabels[$queryData.FormSelectedIndex]) [$currentVal]"
                                    [Console]::CursorVisible = $false

                                    if ($newVal -ne '') {
                                        $queryData[$fKey] = $newVal
                                    }
                                }
                                [Console]::Clear()
                            }
                        }
                        'Tab' {
                            # Tab / Shift+Tab to cycle through tabs
                            if ($key.Modifiers -band [ConsoleModifiers]::Shift) {
                                $newTab = if ($activeTab -eq 0) { $tabNames.Count - 1 } else { $activeTab - 1 }
                            }
                            else {
                                $newTab = if ($activeTab -eq ($tabNames.Count - 1)) { 0 } else { $activeTab + 1 }
                            }
                            $tabState[$activeTab].SelectedIndex = $selectedIndex
                            $tabState[$activeTab].ScrollOffset = $scrollOffset
                            $tabState[$activeTab].Items = $items

                            $activeTab = $newTab

                            if (-not $tabState[$activeTab].Loaded) {
                                $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                $tabState[$activeTab].Items = $items
                                $tabState[$activeTab].Loaded = $true
                            }
                            else {
                                $items = $tabState[$activeTab].Items
                                if (-not $items) { $items = [System.Collections.ArrayList]::new() }
                            }

                            $selectedIndex = $tabState[$activeTab].SelectedIndex
                            $scrollOffset = $tabState[$activeTab].ScrollOffset

                            if ($activeTab -eq 4 -and (-not $queryData.HasSearched)) {
                                $mode = "queryform"
                            }
                            elseif ($activeTab -eq 5 -and (-not $priData.HasParent)) {
                                $mode = "priform"
                            }
                            else {
                                $mode = "list"
                            }

                            [Console]::Clear()
                        }
                        default {
                            # Check for tab switching (digit keys 1-6)
                            $keyChar = $key.KeyChar
                            if ($keyChar -ge '1' -and $keyChar -le '6') {
                                $newTab = [int]::Parse($keyChar.ToString()) - 1
                                if ($newTab -ne $activeTab -and $newTab -lt $tabNames.Count) {
                                    $tabState[$activeTab].SelectedIndex = $selectedIndex
                                    $tabState[$activeTab].ScrollOffset = $scrollOffset
                                    $tabState[$activeTab].Items = $items

                                    $activeTab = $newTab

                                    if (-not $tabState[$activeTab].Loaded) {
                                        $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                        $tabState[$activeTab].Items = $items
                                        $tabState[$activeTab].Loaded = $true
                                    }
                                    else {
                                        $items = $tabState[$activeTab].Items
                                        if (-not $items) { $items = [System.Collections.ArrayList]::new() }
                                    }

                                    $selectedIndex = $tabState[$activeTab].SelectedIndex
                                    $scrollOffset = $tabState[$activeTab].ScrollOffset

                                    if ($activeTab -eq 4 -and (-not $queryData.HasSearched)) {
                                        $mode = "queryform"
                                    }
                                    elseif ($activeTab -eq 5 -and (-not $priData.HasParent)) {
                                        $mode = "priform"
                                    }
                                    else {
                                        $mode = "list"
                                    }

                                    [Console]::Clear()
                                }
                            }
                        }
                    }
                }
                "priform" {
                    Render-PriSearchForm -PriData $priData `
                        -TabNames $tabNames -ActiveTabIndex $activeTab `
                        -StatusMessage $statusMessage
                    $statusMessage = ""

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            if ($priData.FormState -eq 'results') {
                                $priData.FormState = 'input'
                                $priData.SearchResults = @()
                                $priData.SearchResultIndex = 0
                            }
                            elseif ($priData.HasParent) {
                                $mode = 'list'
                                [Console]::Clear()
                            }
                            # else stay in form (no parent yet, can't go back to list)
                        }
                        'UpArrow' {
                            if ($priData.FormState -eq 'results' -and $priData.SearchResultIndex -gt 0) {
                                $priData.SearchResultIndex--
                            }
                        }
                        'DownArrow' {
                            if ($priData.FormState -eq 'results') {
                                $maxIdx = [Math]::Max(0, $priData.SearchResults.Count - 1)
                                if ($priData.SearchResultIndex -lt $maxIdx) {
                                    $priData.SearchResultIndex++
                                }
                            }
                        }
                        'Enter' {
                            if ($priData.FormState -eq 'results') {
                                # Select highlighted result as parent
                                $selectedParent = $priData.SearchResults[$priData.SearchResultIndex]
                                $priData.ParentId    = $selectedParent.Id
                                $priData.ParentTitle = $selectedParent.Title
                                $priData.HasParent   = $true
                                Save-PriParentConfig -ParentId $priData.ParentId -ParentTitle $priData.ParentTitle
                                $priData.FormState   = 'input'
                                $priData.SearchInput = ''
                                $priData.SearchResults = @()
                                $priData.SearchResultIndex = 0

                                [Console]::Clear()
                                Write-Host "`n  Loading items for parent #$($priData.ParentId)..." -ForegroundColor Cyan
                                $tabState[5].EmptyMessage = "No children or related items found for #$($priData.ParentId)."
                                $items = [System.Collections.ArrayList]@(
                                    Get-PriWorkItems -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -ParentId $priData.ParentId `
                                        -IncludeClosed:$tabState[5].ShowAll
                                )
                                $tabState[5].Items  = $items
                                $tabState[5].Loaded = $true
                                $selectedIndex = 0
                                $scrollOffset  = 0
                                $statusMessage = "Loaded $($items.Count) item(s) for parent #$($priData.ParentId)"
                                $mode = 'list'
                                [Console]::Clear()
                            }
                            else {
                                # 'input' state: search or fetch by ID
                                $searchInput = $priData.SearchInput.Trim()
                                if ($searchInput -eq '') {
                                    $statusMessage = 'Please enter a work item ID or title to search'
                                }
                                elseif ($searchInput -match '^\d+$') {
                                    # Fetch by ID directly
                                    [Console]::Clear()
                                    Write-Host "`n  Fetching work item #$searchInput..." -ForegroundColor Cyan
                                    $parentDetail = Get-WorkItemDetail -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT -WorkItemId ([int]$searchInput)
                                    if ($parentDetail) {
                                        $priData.ParentId    = [int]$searchInput
                                        $priData.ParentTitle = $parentDetail.fields.'System.Title'
                                        $priData.HasParent   = $true
                                        Save-PriParentConfig -ParentId $priData.ParentId -ParentTitle $priData.ParentTitle
                                        $priData.SearchInput = ''

                                        [Console]::Clear()
                                        Write-Host "`n  Loading items for parent #$($priData.ParentId)..." -ForegroundColor Cyan
                                        $tabState[5].EmptyMessage = "No children or related items found for #$($priData.ParentId)."
                                        $items = [System.Collections.ArrayList]@(
                                            Get-PriWorkItems -Organization $config.Organization `
                                                -Project $config.Project -PAT $config.PAT `
                                                -ParentId $priData.ParentId `
                                                -IncludeClosed:$tabState[5].ShowAll
                                        )
                                        $tabState[5].Items  = $items
                                        $tabState[5].Loaded = $true
                                        $selectedIndex = 0
                                        $scrollOffset  = 0
                                        $statusMessage = "Loaded $($items.Count) item(s) for parent #$($priData.ParentId)"
                                        $mode = 'list'
                                        [Console]::Clear()
                                    }
                                    else {
                                        $statusMessage = "Work item #$searchInput not found"
                                        [Console]::Clear()
                                    }
                                }
                                else {
                                    # Search by title
                                    [Console]::Clear()
                                    Write-Host "`n  Searching for '$searchInput'..." -ForegroundColor Cyan
                                    $searchResults = @(Search-WorkItemsByFilters `
                                        -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -TitleContains $searchInput)
                                    if ($searchResults.Count -eq 0) {
                                        $statusMessage = "No items found matching '$searchInput'"
                                        [Console]::Clear()
                                    }
                                    elseif ($searchResults.Count -eq 1) {
                                        # Auto-select the single match
                                        $priData.ParentId    = $searchResults[0].Id
                                        $priData.ParentTitle = $searchResults[0].Title
                                        $priData.HasParent   = $true
                                        Save-PriParentConfig -ParentId $priData.ParentId -ParentTitle $priData.ParentTitle
                                        $priData.SearchInput = ''

                                        [Console]::Clear()
                                        Write-Host "`n  Loading items for parent #$($priData.ParentId)..." -ForegroundColor Cyan
                                        $tabState[5].EmptyMessage = "No children or related items found for #$($priData.ParentId)."
                                        $items = [System.Collections.ArrayList]@(
                                            Get-PriWorkItems -Organization $config.Organization `
                                                -Project $config.Project -PAT $config.PAT `
                                                -ParentId $priData.ParentId `
                                                -IncludeClosed:$tabState[5].ShowAll
                                        )
                                        $tabState[5].Items  = $items
                                        $tabState[5].Loaded = $true
                                        $selectedIndex = 0
                                        $scrollOffset  = 0
                                        $statusMessage = "Loaded $($items.Count) item(s) for parent #$($priData.ParentId)"
                                        $mode = 'list'
                                        [Console]::Clear()
                                    }
                                    else {
                                        # Show results to pick from
                                        $priData.SearchResults     = $searchResults
                                        $priData.SearchResultIndex = 0
                                        $priData.FormState         = 'results'
                                        $statusMessage = "Found $($searchResults.Count) items - select a parent"
                                        [Console]::Clear()
                                    }
                                }
                            }
                        }
                        'Backspace' {
                            if ($priData.FormState -eq 'input' -and $priData.SearchInput.Length -gt 0) {
                                $priData.SearchInput = $priData.SearchInput.Substring(0, $priData.SearchInput.Length - 1)
                            }
                        }
                        'Tab' {
                            # Tab / Shift+Tab to cycle through tabs
                            if ($key.Modifiers -band [ConsoleModifiers]::Shift) {
                                $newTab = if ($activeTab -eq 0) { $tabNames.Count - 1 } else { $activeTab - 1 }
                            }
                            else {
                                $newTab = if ($activeTab -eq ($tabNames.Count - 1)) { 0 } else { $activeTab + 1 }
                            }
                            $tabState[$activeTab].SelectedIndex = $selectedIndex
                            $tabState[$activeTab].ScrollOffset  = $scrollOffset
                            $tabState[$activeTab].Items         = $items

                            $activeTab = $newTab

                            if (-not $tabState[$activeTab].Loaded) {
                                $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                $tabState[$activeTab].Items  = $items
                                $tabState[$activeTab].Loaded = $true
                            }
                            else {
                                $items = $tabState[$activeTab].Items
                                if (-not $items) { $items = [System.Collections.ArrayList]::new() }
                            }

                            $selectedIndex = $tabState[$activeTab].SelectedIndex
                            $scrollOffset  = $tabState[$activeTab].ScrollOffset

                            if ($activeTab -eq 4 -and (-not $queryData.HasSearched)) {
                                $mode = 'queryform'
                            }
                            elseif ($activeTab -eq 5 -and (-not $priData.HasParent)) {
                                $mode = 'priform'
                            }
                            else {
                                $mode = 'list'
                            }
                            [Console]::Clear()
                        }
                        default {
                            # Digit keys 1-6 switch tabs
                            $keyChar = $key.KeyChar
                            if ($keyChar -ge '1' -and $keyChar -le '6') {
                                $newTab = [int]::Parse($keyChar.ToString()) - 1
                                if ($newTab -ne $activeTab -and $newTab -lt $tabNames.Count) {
                                    $tabState[$activeTab].SelectedIndex = $selectedIndex
                                    $tabState[$activeTab].ScrollOffset  = $scrollOffset
                                    $tabState[$activeTab].Items         = $items

                                    $activeTab = $newTab

                                    if (-not $tabState[$activeTab].Loaded) {
                                        $items = [System.Collections.ArrayList]@(Refresh-TabItems -Config $config -TabIndex $activeTab -ShowAll $tabState[$activeTab].ShowAll -PriData $priData)
                                        $tabState[$activeTab].Items  = $items
                                        $tabState[$activeTab].Loaded = $true
                                    }
                                    else {
                                        $items = $tabState[$activeTab].Items
                                        if (-not $items) { $items = [System.Collections.ArrayList]::new() }
                                    }

                                    $selectedIndex = $tabState[$activeTab].SelectedIndex
                                    $scrollOffset  = $tabState[$activeTab].ScrollOffset

                                    if ($activeTab -eq 4 -and (-not $queryData.HasSearched)) {
                                        $mode = 'queryform'
                                    }
                                    elseif ($activeTab -eq 5 -and (-not $priData.HasParent)) {
                                        $mode = 'priform'
                                    }
                                    else {
                                        $mode = 'list'
                                    }
                                    [Console]::Clear()
                                }
                            }
                            elseif ($priData.FormState -eq 'input' -and $key.KeyChar -and -not [char]::IsControl($key.KeyChar)) {
                                # Append character to search input
                                $priData.SearchInput += $key.KeyChar
                            }
                        }
                    }
                }

                "statuspicker" {
                    Render-StatusPicker -Item $statusPickerData.Item `
                        -Statuses $statusPickerData.Statuses `
                        -SelectedStatusIndex $statusPickerData.SelectedIndex

                    $key = [Console]::ReadKey($true)

                    switch ($key.Key) {
                        'Escape' {
                            $statusPickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        'UpArrow' {
                            if ($statusPickerData.SelectedIndex -gt 0) {
                                $statusPickerData.SelectedIndex--
                            }
                        }
                        'DownArrow' {
                            if ($statusPickerData.SelectedIndex -lt ($statusPickerData.Statuses.Count - 1)) {
                                $statusPickerData.SelectedIndex++
                            }
                        }
                        'Enter' {
                            $newState = $statusPickerData.Statuses[$statusPickerData.SelectedIndex]
                            $item = $statusPickerData.Item

                            if ($newState -ne $item.State) {
                                try {
                                    Update-WorkItemState -Organization $config.Organization `
                                        -Project $config.Project -PAT $config.PAT `
                                        -WorkItemId $item.Id `
                                        -NewState $newState

                                    $item['State'] = $newState
                                    $statusMessage = "Status of #$($item.Id) changed to '$newState'"
                                }
                                catch {
                                    $statusMessage = "Error: $($_.Exception.Message)"
                                }
                            }
                            else {
                                $statusMessage = "Status unchanged"
                            }

                            $statusPickerData = $null
                            $mode = "detail"
                            [Console]::Clear()
                        }
                        default { }
                    }
                }
            }
        }
    }
    finally {
        # Save any remaining active timers on unexpected exit
        foreach ($timerId in @($activeTimers.Keys)) {
            $timerItem = $items | Where-Object { $_.Id -eq $timerId } | Select-Object -First 1
            if ($timerItem) {
                try { Save-Timer -Item $timerItem -Stopwatch $activeTimers[$timerId] -Config $config | Out-Null } catch { }
            }
        }
        [Console]::CursorVisible = $true
        [Console]::ResetColor()
    }
}
