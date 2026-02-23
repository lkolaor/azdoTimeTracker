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
    function Refresh-Items {
        param($Config)

        Write-Host "`n  Loading work items from Azure DevOps..." -ForegroundColor Cyan

        $result = Get-MyWorkItems -Organization $Config.Organization `
                                  -Project $Config.Project `
                                  -PAT $Config.PAT

        if (-not $result -or ($result.MyItems.Count -eq 0)) {
            return [System.Collections.ArrayList]::new()
        }

        $tree = Build-WorkItemTree -MyItems $result.MyItems -ParentItems $result.ParentItems
        return $tree
    }

    $items = [System.Collections.ArrayList]@(Refresh-Items -Config $config)

    # ── State ────────────────────────────────────────────────────────
    $selectedIndex = 0
    $scrollOffset = 0
    $statusMessage = ""
    $mode = "list"  # list | detail | statuspicker | commentpicker | fieldpicker | toolsmenu
    $detailScrollOffset = 0
    $detailData = $null
    $activeTimers = @{}          # hashtable: workItemId -> Stopwatch
    $statusPickerData = $null    # @{ Item; Statuses; SelectedIndex }
    $commentPickerData = $null   # @{ Item; Comments; SelectedIndex; Action }
    $fieldPickerData = $null     # @{ Item; Fields; SelectedIndex }
    $toolsMenuData = $null       # @{ SelectedIndex; MenuItems }

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
            # Always work with a filtered list of valid items
            $validItems = @($items | Where-Object { $_.Id -and $_.Title })

            # Ensure selectedIndex is always valid for the filtered list
            if ($selectedIndex -lt 0) { $selectedIndex = 0 }
            if ($selectedIndex -ge $validItems.Count) { $selectedIndex = [Math]::Max(0, $validItems.Count - 1) }

            switch ($mode) {
                "list" {
                    $hasTimers = $activeTimers.Count -gt 0

                    $scrollOffset = Render-WorkItemList -Items $validItems `
                        -SelectedIndex $selectedIndex -ScrollOffset $scrollOffset `
                        -StatusMessage $statusMessage -ActiveTimers $activeTimers
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
                            if ($selectedIndex -gt 0) { $selectedIndex-- }
                        }
                        'DownArrow' {
                            if ($selectedIndex -lt ($validItems.Count - 1)) { $selectedIndex++ }
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
                            $statusMessage = "Refreshing..."
                            [Console]::Clear()
                            $items = [System.Collections.ArrayList]@(Refresh-Items -Config $config)
                            $validItems = @($items | Where-Object { $_.Id -and $_.Title })
                            $selectedIndex = 0
                            $scrollOffset = 0
                            $statusMessage = "Refreshed - $($validItems.Count) items loaded"
                            [Console]::Clear()
                        }
                        'Enter' {
                            # Show detail view
                            if ($validItems.Count -gt 0) {
                                $item = $validItems[$selectedIndex]
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
                                                ParentId         = $itemId
                                                OriginalEstimate = 5.0
                                                CompletedWork    = $null
                                                RemainingWork    = 5.0
                                                IsMine           = $true
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
                                        $activeTimers[$itemId] = [System.Diagnostics.Stopwatch]::StartNew()
                                        $statusMessage = "Started timer on #$itemId"
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
                        default {
                            # Ignore other keys
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
                        default { }
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

                                    # Reload items with new config
                                    [Console]::Clear()
                                    $items = [System.Collections.ArrayList]@(Refresh-Items -Config $config)
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
