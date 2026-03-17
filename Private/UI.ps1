<#
.SYNOPSIS
    Terminal UI module for the Azure DevOps Time Tracker.
.DESCRIPTION
    Handles rendering the work item list, detail view, status picker, and keyboard input.
#>

# ── Type icons and colours ─────────────────────────────────────────
$script:TypeConfig = @{
    'Epic'                = @{ Icon = [char]0x2605; Color = 'Magenta'     }
    'Feature'             = @{ Icon = [char]0x25C6; Color = 'DarkMagenta' }
    'User Story'          = @{ Icon = [char]0x25CF; Color = 'Cyan'        }
    'Product Backlog Item'= @{ Icon = [char]0x25CF; Color = 'Cyan'        }
    'Bug'                 = @{ Icon = [char]0x25A0; Color = 'Red'         }
    'Task'                = @{ Icon = [char]0x25B6; Color = 'Green'       }
    'Incident'            = @{ Icon = [char]0x26A0; Color = 'Yellow'      }
    'Issue'               = @{ Icon = [char]0x26A0; Color = 'Yellow'      }
}

function Get-TypeIcon {
    param([string]$Type)
    $cfg = $script:TypeConfig[$Type]
    if ($cfg) { return $cfg.Icon } else { return [char]0x25CB } # ○
}

function Get-TypeColor {
    param([string]$Type)
    $cfg = $script:TypeConfig[$Type]
    if ($cfg) { return $cfg.Color } else { return 'White' }
}

# ── Strip HTML from description ─────────────────────────────────────
function Remove-Html {
    param([string]$Html)
    if (-not $Html) { return "" }
    # Replace common HTML entities
    $text = $Html -replace '<br\s*/?>', "`n"
    $text = $text -replace '<p[^>]*>', "`n"
    $text = $text -replace '</p>', ""
    $text = $text -replace '<div[^>]*>', "`n"
    $text = $text -replace '</div>', ""
    $text = $text -replace '<li[^>]*>', "  - "
    $text = $text -replace '</li>', "`n"
    $text = $text -replace '<[^>]+>', ''
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    # Clean up excess blank lines
    $text = $text -replace "(\r?\n){3,}", "`n`n"
    return $text.Trim()
}

# ── Format elapsed time from a Stopwatch ────────────────────────────
function Format-Elapsed {
    param([System.Diagnostics.Stopwatch]$SW)
    $e = $SW.Elapsed
    [int]$h = [Math]::Floor($e.TotalHours)
    [int]$m = $e.Minutes
    [int]$s = $e.Seconds
    return "{0:D2}:{1:D2}:{2:D2}" -f $h, $m, $s
}

# ── Calculate display width accounting for wide characters ──────────
function Get-DisplayWidth {
    param([string]$Text)
    if (-not $Text) { return 0 }
    $w = 0
    foreach ($c in $Text.ToCharArray()) {
        $cp = [int]$c
        # Common double-width ranges: CJK, fullwidth, some symbols/emoji
        if (($cp -ge 0x1100 -and $cp -le 0x115F) -or   # Hangul Jamo
            ($cp -ge 0x2E80 -and $cp -le 0x303E) -or   # CJK Radicals
            ($cp -ge 0x3040 -and $cp -le 0x33BF) -or   # Japanese/CJK
            ($cp -ge 0x3400 -and $cp -le 0x4DBF) -or   # CJK Ext A
            ($cp -ge 0x4E00 -and $cp -le 0x9FFF) -or   # CJK Unified
            ($cp -ge 0xA000 -and $cp -le 0xA4CF) -or   # Yi
            ($cp -ge 0xAC00 -and $cp -le 0xD7AF) -or   # Hangul
            ($cp -ge 0xF900 -and $cp -le 0xFAFF) -or   # CJK Compat
            ($cp -ge 0xFE30 -and $cp -le 0xFE6F) -or   # CJK Compat Forms
            ($cp -ge 0xFF01 -and $cp -le 0xFF60) -or   # Fullwidth
            ($cp -ge 0xFFE0 -and $cp -le 0xFFE6) -or   # Fullwidth signs
            ($cp -ge 0x2B50 -and $cp -le 0x2B55) -or   # Emoji stars/circles
            ($cp -ge 0x1F300 -and $cp -le 0x1F9FF) -or # Emoji block
            ($cp -ge 0xFE00 -and $cp -le 0xFE0F)) {    # Variation selectors
            $w += 2
        }
        else {
            $w += 1
        }
    }
    return $w
}

# ── Pad or truncate string to exact display width ───────────────────
function Format-FixedWidth {
    param([string]$Text, [int]$Width)
    $dw = Get-DisplayWidth -Text $Text
    if ($dw -gt $Width) {
        # Truncate: remove chars from end until display width fits
        $result = ""
        $currentWidth = 0
        foreach ($c in $Text.ToCharArray()) {
            $cp = [int]$c
            $cw = 1
            if (($cp -ge 0x1100 -and $cp -le 0x115F) -or
                ($cp -ge 0x2E80 -and $cp -le 0x303E) -or
                ($cp -ge 0x3040 -and $cp -le 0x33BF) -or
                ($cp -ge 0x3400 -and $cp -le 0x4DBF) -or
                ($cp -ge 0x4E00 -and $cp -le 0x9FFF) -or
                ($cp -ge 0xA000 -and $cp -le 0xA4CF) -or
                ($cp -ge 0xAC00 -and $cp -le 0xD7AF) -or
                ($cp -ge 0xF900 -and $cp -le 0xFAFF) -or
                ($cp -ge 0xFE30 -and $cp -le 0xFE6F) -or
                ($cp -ge 0xFF01 -and $cp -le 0xFF60) -or
                ($cp -ge 0xFFE0 -and $cp -le 0xFFE6) -or
                ($cp -ge 0x2600 -and $cp -le 0x27BF) -or
                ($cp -ge 0x2B50 -and $cp -le 0x2B55) -or
                ($cp -ge 0xFE00 -and $cp -le 0xFE0F)) {
                $cw = 2
            }
            if (($currentWidth + $cw) -gt ($Width - 3)) {
                $result += "..."
                break
            }
            $result += $c
            $currentWidth += $cw
        }
        return $result + (" " * [Math]::Max(0, $Width - (Get-DisplayWidth -Text $result)))
    }
    else {
        return $Text + (" " * [Math]::Max(0, $Width - $dw))
    }
}

# ── Word-wrap text to fit within a given width ─────────────────────
function Wrap-Text {
    param([string]$Text, [int]$MaxWidth, [string]$Prefix = "")
    $result = [System.Collections.ArrayList]::new()
    if (-not $Text) { [void]$result.Add($Prefix); return $result }
    $usable = $MaxWidth - $Prefix.Length
    if ($usable -lt 10) { $usable = 10 }
    foreach ($rawLine in ($Text -split "`n")) {
        $line = $rawLine.TrimEnd()
        if ($line.Length -eq 0) {
            [void]$result.Add($Prefix)
            continue
        }
        while ($line.Length -gt $usable) {
            # Find last space within usable width
            $breakAt = $line.LastIndexOf(' ', [Math]::Min($usable, $line.Length - 1))
            if ($breakAt -le 0) { $breakAt = $usable }
            [void]$result.Add($Prefix + $line.Substring(0, $breakAt).TrimEnd())
            $line = $line.Substring($breakAt).TrimStart()
        }
        if ($line.Length -gt 0) {
            [void]$result.Add($Prefix + $line)
        }
    }
    return $result
}

# ── Render tab bar line ─────────────────────────────────────────────
function Render-TabBarLine {
    param(
        [array]$TabNames,
        [int]$ActiveTabIndex,
        [int]$Width
    )

    $tabX = 0
    for ($ti = 0; $ti -lt $TabNames.Count; $ti++) {
        $label = " $($ti + 1):$($TabNames[$ti]) "
        if ($ti -eq $ActiveTabIndex) {
            Write-Host $label -ForegroundColor White -BackgroundColor DarkCyan -NoNewline
        }
        else {
            Write-Host $label -ForegroundColor Gray -NoNewline
        }
        if ($ti -lt $TabNames.Count - 1) {
            Write-Host "|" -ForegroundColor DarkGray -NoNewline
            $tabX += 1
        }
        $tabX += $label.Length
    }
    $remaining = [Math]::Max(0, $Width - $tabX)
    if ($remaining -gt 0) {
        Write-Host (" " * $remaining)
    }
    else {
        Write-Host ""
    }
}

# ── Render the main list ────────────────────────────────────────────
function Render-WorkItemList {
    param(
        [System.Collections.ArrayList]$Items,
        [int]$SelectedIndex,
        [int]$ScrollOffset,
        [string]$StatusMessage = "",
        [hashtable]$ActiveTimers = @{},
        [array]$TabNames = @(),
        [int]$ActiveTabIndex = -1,
        [bool]$ShowAllEnabled = $false,
        [bool]$ShowAllAvailable = $false,
        [string]$NoItemsMessage = "No work items found.",
        [bool]$SearchActive = $false,
        [string]$SearchQuery = ""
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " AZURE DEVOPS TIME TRACKER "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkBlue

    # Tab bar
    $tabBarLines = 0
    if ($TabNames.Count -gt 0 -and $ActiveTabIndex -ge 0) {
        $tabBarLines = 1
        Render-TabBarLine -TabNames $TabNames -ActiveTabIndex $ActiveTabIndex -Width $width
    }

    $helpParts = " [Arrows] Navigate  [Enter] Info  [t] Timer  [r] Refresh  [m] Tools"
    if ($ShowAllAvailable) {
        $helpParts += if ($ShowAllEnabled) { "  [x] Active only" } else { "  [x] Show all" }
    }
    if ($SearchActive) {
        $helpParts = " [ESC] Clear filter  [Arrows] Navigate  [Enter] Info  [t] Timer  [Backspace] Delete char "
    }
    elseif ($ActiveTabIndex -eq 4) { $helpParts += "  [/] Search" }
    elseif ($ActiveTabIndex -eq 5) { $helpParts += "  [/] Set Parent" }
    else { $helpParts += "  [/] Filter" }
    if (-not $SearchActive -and $TabNames.Count -gt 0) { $helpParts += "  [Tab] Switch tab" }
    $helpParts += "  [q] Quit "
    $helpLine = $helpParts
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    # Available lines for items
    $availableLines = $height - 4 - $tabBarLines  # header (1) + [tab bar] + help (1) + status (1) + bottom margin (1)

    if ($Items.Count -eq 0) {
        Write-Host ""
        if ($SearchActive -and $SearchQuery.Length -gt 0) {
            Write-Host "  No matches for '$SearchQuery'" -ForegroundColor Yellow
        } else {
            Write-Host "  $NoItemsMessage" -ForegroundColor Yellow
        }
        Write-Host "  Press 'r' to refresh or 'q' to quit." -ForegroundColor Gray
        $remaining = $availableLines - 2
        for ($l = 0; $l -lt $remaining; $l++) {
            Write-Host (" " * $width)
        }
    }
    else {
        # Filter out any items missing Id or Title (but keep separator rows)
        $validItems = $Items | Where-Object { ($_.Id -and $_.Title) -or $_.IsSeparator }
        # Adjust scroll offset so selected item is visible and valid
        if ($SelectedIndex -lt $ScrollOffset) {
            $ScrollOffset = $SelectedIndex
        }
        if ($SelectedIndex -ge ($ScrollOffset + $availableLines)) {
            $ScrollOffset = $SelectedIndex - $availableLines + 1
        }

        $rendered = 0
        for ($line = 0; $line -lt $availableLines; $line++) {
            $idx = $ScrollOffset + $line
            if ($idx -lt $validItems.Count) {
                $item = $validItems[$idx]

                # Render separator rows as dimmed section headers
                if ($item.IsSeparator) {
                    $sepLine = "  $($item.Title)"
                    Write-Host (Format-FixedWidth -Text $sepLine -Width $width) -ForegroundColor DarkYellow
                    $rendered++
                    continue
                }

                $indent = "  " * $item.Depth
                $icon = Get-TypeIcon -Type $item.Type
                $typeColor = Get-TypeColor -Type $item.Type

                # Active timer indicator
                $timerStr = ""
                $itemId = $item.Id
                $isTracking = $ActiveTimers.ContainsKey($itemId)
                if ($isTracking) {
                    $sw = $ActiveTimers[$itemId]
                    $timerStr = " [$(Format-Elapsed -SW $sw)]"
                }

                # Time info (completed/remaining)
                $timeStr = ""
                if ($null -ne $item.RemainingWork -or $null -ne $item.CompletedWork) {
                    $cw = if ($null -ne $item.CompletedWork) { "{0:N1}h" -f [double]$item.CompletedWork } else { "-" }
                    $rw = if ($null -ne $item.RemainingWork) { "{0:N1}h" -f [double]$item.RemainingWork } else { "-" }
                    $timeStr = " [$cw/$rw]"
                }

                $prefix = "  "
                if ($idx -eq $SelectedIndex) {
                    $prefix = "> "
                }

                $lineText = "$prefix$indent$icon $($item.Id)/$($item.State): $($item.Title)$timeStr$timerStr"

                $padded = Format-FixedWidth -Text $lineText -Width $width

                if ($isTracking) {
                    if ($idx -eq $SelectedIndex) {
                        Write-Host $padded -ForegroundColor White -BackgroundColor DarkYellow
                    }
                    else {
                        Write-Host $padded -ForegroundColor DarkYellow
                    }
                }
                elseif ($idx -eq $SelectedIndex) {
                    Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
                }
                elseif ($item.IsRelated) {
                    Write-Host $padded -ForegroundColor Yellow
                }
                elseif (-not $item.IsMine) {
                    Write-Host $padded -ForegroundColor DarkGray
                }
                else {
                    Write-Host $padded -ForegroundColor $typeColor
                }
                $rendered++
            }
            else {
                Write-Host (" " * $width)
            }
        }
    }

    # Status bar
    $timerCount = $ActiveTimers.Count
    $timerInfo = if ($timerCount -gt 0) { " | $timerCount timer(s) active" } else { "" }

    if ($SearchActive) {
        $cursor = "_"
        $matchWord = if ($Items.Count -eq 1) { "match" } else { "matches" }
        $matchInfo = if ($SearchQuery.Length -gt 0) { "  ($($Items.Count) $matchWord)" } else { "  (type to filter)" }
        $statusText = " /$SearchQuery$cursor$matchInfo$timerInfo"
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor White -BackgroundColor DarkMagenta -NoNewline
    }
    elseif ($StatusMessage) {
        $statusText = " $StatusMessage$timerInfo"
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor White -BackgroundColor DarkGreen -NoNewline
    }
    else {
        $countMsg = " $($Items.Count) items$timerInfo"
        $statusPadded = $countMsg + (" " * [Math]::Max(0, $width - $countMsg.Length))
        if ($timerCount -gt 0) {
            Write-Host $statusPadded -ForegroundColor White -BackgroundColor DarkRed -NoNewline
        }
        else {
            Write-Host $statusPadded -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline
        }
    }

    return $ScrollOffset
}

# ── Render the Tools menu ───────────────────────────────────────────
function Render-ToolsMenu {
    param(
        [int]$SelectedIndex,
        [array]$MenuItems
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " TOOLS "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkBlue

    $helpLine = " [Arrows] Navigate  [Enter] Select  [ESC] Back "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    Write-Host ""

    for ($i = 0; $i -lt $MenuItems.Count; $i++) {
        $prefix = "  "
        if ($i -eq $SelectedIndex) { $prefix = "> " }
        $lineText = "$prefix$($MenuItems[$i].Label)"
        $padded = Format-FixedWidth -Text $lineText -Width $width
        if ($i -eq $SelectedIndex) {
            Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
        }
        else {
            Write-Host $padded -ForegroundColor White
        }
    }

    # Fill remaining lines
    $usedLines = 3 + $MenuItems.Count  # header + help + blank + items
    $remaining = $height - $usedLines - 1
    for ($l = 0; $l -lt $remaining; $l++) {
        Write-Host (" " * $width)
    }

    # Status bar
    $statusText = " Select an option "
    $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
    Write-Host $statusPadded -ForegroundColor Gray -BackgroundColor DarkGray
}

# ── Render info/detail page ─────────────────────────────────────────
function Render-DetailView {
    param(
        [hashtable]$Item,
        [string]$Description,
        [string]$ReproSteps = "",
        [string]$SystemInfo = "",
        [array]$Comments,
        [int]$ScrollOffset,
        [string]$Organization = "",
        [string]$Project = ""
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Build content lines
    $lines = [System.Collections.ArrayList]::new()

    [void]$lines.Add("")
    [void]$lines.Add("  === $($Item.Type) #$($Item.Id): $($Item.Title) ===")
    [void]$lines.Add("")
    if ($Organization -and $Project) {
        $adoUrl = "https://dev.azure.com/$Organization/$([Uri]::EscapeDataString($Project))/_workitems/edit/$($Item.Id)"
        [void]$lines.Add("  URL:   $adoUrl")
    }
    [void]$lines.Add("  State:    $($Item.State)")
    $assignee = if ($Item.ContainsKey('AssignedTo') -and $Item.AssignedTo) { $Item.AssignedTo } else { "(unassigned)" }
    [void]$lines.Add("  Assigned: $assignee")

    if ($null -ne $Item.OriginalEstimate) {
        [void]$lines.Add("  Original Estimate: $($Item.OriginalEstimate) hours")
    }
    if ($null -ne $Item.CompletedWork) {
        [void]$lines.Add("  Completed Work:    $($Item.CompletedWork) hours")
    }
    if ($null -ne $Item.RemainingWork) {
        [void]$lines.Add("  Remaining Work:    $($Item.RemainingWork) hours")
    }

    [void]$lines.Add("")
    [void]$lines.Add("  --- Description ---")

    $descPlain = Remove-Html -Html $Description
    if ($descPlain) {
        $wrappedDesc = Wrap-Text -Text $descPlain -MaxWidth $width -Prefix "  "
        foreach ($wl in $wrappedDesc) { [void]$lines.Add($wl) }
    }
    else {
        [void]$lines.Add("  (no description)")
    }

    if ($ReproSteps) {
        [void]$lines.Add("")
        [void]$lines.Add("  --- Repro Steps ---")
        $reproPlain = Remove-Html -Html $ReproSteps
        $wrappedRepro = Wrap-Text -Text $reproPlain -MaxWidth $width -Prefix "  "
        foreach ($wl in $wrappedRepro) { [void]$lines.Add($wl) }
    }

    if ($SystemInfo) {
        [void]$lines.Add("")
        [void]$lines.Add("  --- System Info ---")
        $sysPlain = Remove-Html -Html $SystemInfo
        $wrappedSys = Wrap-Text -Text $sysPlain -MaxWidth $width -Prefix "  "
        foreach ($wl in $wrappedSys) { [void]$lines.Add($wl) }
    }

    [void]$lines.Add("")
    [void]$lines.Add("  --- Comments ($($Comments.Count)) ---")

    if ($Comments.Count -eq 0) {
        [void]$lines.Add("  (no comments)")
    }
    else {
        foreach ($comment in $Comments) {
            [void]$lines.Add("")
            $author = "Unknown"
            if ($comment.createdBy -and $comment.createdBy.displayName) {
                $author = $comment.createdBy.displayName
            }
            $date = ""
            if ($comment.createdDate) {
                $date = ([datetime]$comment.createdDate).ToString("yyyy-MM-dd HH:mm")
            }
            [void]$lines.Add("  [$date] ${author}:")
            $commentText = Remove-Html -Html $comment.text
            $wrappedComment = Wrap-Text -Text $commentText -MaxWidth $width -Prefix "    "
            foreach ($wl in $wrappedComment) { [void]$lines.Add($wl) }
        }
    }

    # Header
    $header = " WORK ITEM DETAILS "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkBlue

    $helpLine = " [ESC] Back  [Arrows] Scroll  [s] Status  [n] Assign  [h] Hours  [f] Edit Fields  [a] Add Comment  [e] Edit  [d] Delete "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    $availableLines = $height - 3

    for ($l = 0; $l -lt $availableLines; $l++) {
        $lineIdx = $ScrollOffset + $l
        if ($lineIdx -lt $lines.Count) {
            $lineText = $lines[$lineIdx]
            $padded = Format-FixedWidth -Text $lineText -Width $width
            Write-Host $padded -ForegroundColor White
        }
        else {
            Write-Host (" " * $width)
        }
    }

    # Status
    $scrollInfo = " Line $($ScrollOffset + 1) of $($lines.Count) "
    $padded3 = $scrollInfo + (" " * [Math]::Max(0, $width - $scrollInfo.Length))
    Write-Host $padded3 -ForegroundColor Gray -BackgroundColor DarkGray

    return @{
        ScrollOffset = $ScrollOffset
        TotalLines   = $lines.Count
    }
}

# ── Multi-line text editor ──────────────────────────────────────────
function Edit-TextBlock {
    <#
    .SYNOPSIS
        A simple full-screen multi-line text editor.
        Returns the edited text as a single string with newline line breaks,
        or $null if the user pressed Ctrl+Q to cancel.
    #>
    param(
        [string]$Title = "Edit Comment",
        [string]$InitialText = ""
    )

    [Console]::CursorVisible = $true
    [Console]::Clear()

    $width  = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Split initial text into editable line buffers
    $lines = [System.Collections.ArrayList]::new()
    $rawLines = @($InitialText -split "`n")
    foreach ($rl in $rawLines) {
        [void]$lines.Add([System.Collections.ArrayList]::new([char[]]$rl.TrimEnd()))
    }
    if ($lines.Count -eq 0) {
        [void]$lines.Add([System.Collections.ArrayList]::new())
    }

    $cursorRow = 0
    $cursorCol = $lines[0].Count   # start at end of first line
    $scrollRow = 0

    $headerLines = 2   # header + help bar
    $footerLines = 1   # status bar
    $editHeight  = $height - $headerLines - $footerLines

    function Redraw {
        [Console]::SetCursorPosition(0, 0)

        # Header
        $hdr = " $Title "
        $padH = [Math]::Max(0, $width - $hdr.Length)
        Write-Host ($hdr + (" " * $padH)) -ForegroundColor White -BackgroundColor DarkBlue

        $help = " [Ctrl+S] Save  [Ctrl+Q] Cancel  [Enter] New Line  [Arrows] Move "
        $padHelp = [Math]::Max(0, $width - $help.Length)
        Write-Host ($help + (" " * $padHelp)) -ForegroundColor Gray -BackgroundColor DarkGray

        for ($row = 0; $row -lt $editHeight; $row++) {
            $lineIdx = $scrollRow + $row
            if ($lineIdx -lt $lines.Count) {
                $lineText = ($lines[$lineIdx] -join "")
                if ($lineText.Length -gt ($width - 1)) {
                    $lineText = $lineText.Substring(0, $width - 1)
                }
                $padded = $lineText + (" " * [Math]::Max(0, $width - $lineText.Length))
            }
            else {
                $padded = "~" + (" " * [Math]::Max(0, $width - 1))
            }
            Write-Host $padded -ForegroundColor White -NoNewline
            # move to next line without scrolling
            if ($row -lt $editHeight - 1) { Write-Host "" }
        }

        # Footer
        [Console]::SetCursorPosition(0, $height - 1)
        $lineInfo = " Line $($cursorRow + 1)/$($lines.Count)  Col $($cursorCol + 1) "
        $padF = [Math]::Max(0, $width - $lineInfo.Length)
        Write-Host ($lineInfo + (" " * $padF)) -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline

        # Position real cursor
        $screenRow = $headerLines + ($cursorRow - $scrollRow)
        $screenCol = [Math]::Min($cursorCol, $width - 1)
        [Console]::SetCursorPosition($screenCol, $screenRow)
    }

    while ($true) {
        # Ensure cursor is within bounds
        if ($cursorRow -lt 0) { $cursorRow = 0 }
        if ($cursorRow -ge $lines.Count) { $cursorRow = $lines.Count - 1 }
        $lineLen = $lines[$cursorRow].Count
        if ($cursorCol -gt $lineLen) { $cursorCol = $lineLen }
        if ($cursorCol -lt 0) { $cursorCol = 0 }

        # Scroll to keep cursor visible
        if ($cursorRow -lt $scrollRow) { $scrollRow = $cursorRow }
        if ($cursorRow -ge $scrollRow + $editHeight) { $scrollRow = $cursorRow - $editHeight + 1 }

        Redraw

        $key = [Console]::ReadKey($true)

        # Ctrl+S = save
        if ($key.Key -eq 'S' -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
            $result = [System.Collections.ArrayList]::new()
            foreach ($line in $lines) {
                [void]$result.Add(($line -join ""))
            }
            [Console]::CursorVisible = $false
            return ($result -join "`n")
        }
        # Ctrl+Q = cancel
        if ($key.Key -eq 'Q' -and ($key.Modifiers -band [ConsoleModifiers]::Control)) {
            [Console]::CursorVisible = $false
            return $null
        }

        switch ($key.Key) {
            'UpArrow' {
                if ($cursorRow -gt 0) { $cursorRow-- }
            }
            'DownArrow' {
                if ($cursorRow -lt ($lines.Count - 1)) { $cursorRow++ }
            }
            'LeftArrow' {
                if ($cursorCol -gt 0) {
                    $cursorCol--
                }
                elseif ($cursorRow -gt 0) {
                    $cursorRow--
                    $cursorCol = $lines[$cursorRow].Count
                }
            }
            'RightArrow' {
                if ($cursorCol -lt $lines[$cursorRow].Count) {
                    $cursorCol++
                }
                elseif ($cursorRow -lt ($lines.Count - 1)) {
                    $cursorRow++
                    $cursorCol = 0
                }
            }
            'Home' {
                $cursorCol = 0
            }
            'End' {
                $cursorCol = $lines[$cursorRow].Count
            }
            'Enter' {
                # Split current line at cursor
                $tail = [System.Collections.ArrayList]::new()
                if ($cursorCol -lt $lines[$cursorRow].Count) {
                    for ($i = $cursorCol; $i -lt $lines[$cursorRow].Count; $i++) {
                        [void]$tail.Add($lines[$cursorRow][$i])
                    }
                    $lines[$cursorRow].RemoveRange($cursorCol, $lines[$cursorRow].Count - $cursorCol)
                }
                $cursorRow++
                $lines.Insert($cursorRow, $tail)
                $cursorCol = 0
            }
            'Backspace' {
                if ($cursorCol -gt 0) {
                    $lines[$cursorRow].RemoveAt($cursorCol - 1)
                    $cursorCol--
                }
                elseif ($cursorRow -gt 0) {
                    # Merge with previous line
                    $prevLen = $lines[$cursorRow - 1].Count
                    foreach ($ch in $lines[$cursorRow]) {
                        [void]$lines[$cursorRow - 1].Add($ch)
                    }
                    $lines.RemoveAt($cursorRow)
                    $cursorRow--
                    $cursorCol = $prevLen
                }
            }
            'Delete' {
                if ($cursorCol -lt $lines[$cursorRow].Count) {
                    $lines[$cursorRow].RemoveAt($cursorCol)
                }
                elseif ($cursorRow -lt ($lines.Count - 1)) {
                    # Merge next line into current
                    foreach ($ch in $lines[$cursorRow + 1]) {
                        [void]$lines[$cursorRow].Add($ch)
                    }
                    $lines.RemoveAt($cursorRow + 1)
                }
            }
            default {
                if ($key.KeyChar -ne "`0" -and -not [char]::IsControl($key.KeyChar)) {
                    $lines[$cursorRow].Insert($cursorCol, $key.KeyChar)
                    $cursorCol++
                }
            }
        }
    }
}

# ── Render comment picker (arrow-navigated) ───────────────────────
function Render-CommentPicker {
    param(
        [hashtable]$Item,
        [array]$Comments,
        [int]$SelectedIndex,
        [string]$Action   # "edit" or "delete"
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width  = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    $actionLabel = if ($Action -eq "edit") { "EDIT" } else { "DELETE" }

    # Header
    $header = " SELECT COMMENT TO $actionLabel "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkMagenta

    $helpLine = " [Up/Down] Select  [Enter] Confirm  [ESC] Cancel "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    $availableLines = $height - 3

    # Map each comment to the range of content-line indices it occupies
    $contentLines  = [System.Collections.ArrayList]::new()
    $commentFirstLine = @{}   # commentIndex -> first contentLine index

    [void]$contentLines.Add("")
    [void]$contentLines.Add("  $($Item.Type) #$($Item.Id): $($Item.Title)")
    [void]$contentLines.Add("")

    for ($ci = 0; $ci -lt $Comments.Count; $ci++) {
        $c = $Comments[$ci]
        $author = "Unknown"
        if ($c.createdBy -and $c.createdBy.displayName) { $author = $c.createdBy.displayName }
        $date = ""
        if ($c.createdDate) { $date = ([datetime]$c.createdDate).ToString("yyyy-MM-dd HH:mm") }

        $prefix = if ($ci -eq $SelectedIndex) { " >" } else { "  " }

        $commentFirstLine[$ci] = $contentLines.Count
        [void]$contentLines.Add("$prefix [$date] ${author}:")

        # Show ALL lines of the comment text, not just one preview line
        $commentPlain = Remove-Html -Html $c.text
        $textLines = @($commentPlain -split "`n")
        foreach ($tl in $textLines) {
            $tlTrimmed = $tl.TrimEnd()
            if ($tlTrimmed.Length -gt ($width - 8)) {
                $tlTrimmed = $tlTrimmed.Substring(0, [Math]::Max(0, $width - 11)) + "..."
            }
            [void]$contentLines.Add("$prefix   $tlTrimmed")
        }
        [void]$contentLines.Add("")
    }

    # Scroll so the selected comment is visible
    $scrollOffset = 0
    if ($SelectedIndex -ge 0 -and $commentFirstLine.ContainsKey($SelectedIndex)) {
        $selStart = $commentFirstLine[$SelectedIndex]
        if ($selStart -ge $availableLines) {
            $scrollOffset = $selStart - 2   # show a bit of context above
            $scrollOffset = [Math]::Max(0, $scrollOffset)
        }
    }

    # Build highlight set for the selected comment's lines
    $highlightSet = [System.Collections.Generic.HashSet[int]]::new()
    if ($SelectedIndex -ge 0 -and $commentFirstLine.ContainsKey($SelectedIndex)) {
        $hlStart = $commentFirstLine[$SelectedIndex]
        # Highlight until the next blank separator line (exclusive)
        $hlEnd = if ($commentFirstLine.ContainsKey($SelectedIndex + 1)) {
            $commentFirstLine[$SelectedIndex + 1] - 1   # blank line before next comment
        } else {
            $contentLines.Count - 1
        }
        for ($h = $hlStart; $h -le $hlEnd; $h++) {
            [void]$highlightSet.Add($h)
        }
    }

    for ($l = 0; $l -lt $availableLines; $l++) {
        $ci2 = $scrollOffset + $l
        if ($ci2 -lt $contentLines.Count) {
            $lineText = $contentLines[$ci2]
            if ($lineText.Length -gt $width) { $lineText = $lineText.Substring(0, $width - 3) + "..." }
            $padded = $lineText + (" " * [Math]::Max(0, $width - $lineText.Length))
            if ($highlightSet.Contains($ci2)) {
                Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
            }
            else {
                Write-Host $padded -ForegroundColor White
            }
        }
        else {
            Write-Host (" " * $width)
        }
    }

    $statusLine = " Comment $($SelectedIndex + 1) of $($Comments.Count) "
    $padded3 = $statusLine + (" " * [Math]::Max(0, $width - $statusLine.Length))
    Write-Host $padded3 -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline
}

# ── Render field picker ───────────────────────────────────────────
function Render-FieldPicker {
    param(
        [hashtable]$Item,
        [array]$Fields,
        [int]$SelectedIndex
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width  = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " EDIT FIELD "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkMagenta

    $helpLine = " [Up/Down] Select  [Enter] Edit  [ESC] Cancel "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    $availableLines = $height - 3

    $contentLines = [System.Collections.ArrayList]::new()
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  $($Item.Type) #$($Item.Id): $($Item.Title)")
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  Select a field to edit:")
    [void]$contentLines.Add("")

    $fieldStartLine = $contentLines.Count

    for ($i = 0; $i -lt $Fields.Count; $i++) {
        $prefix = if ($i -eq $SelectedIndex) { "  > " } else { "    " }
        $f = $Fields[$i]
        $preview = $f.Preview
        if ($preview.Length -gt ($width - 30)) {
            $preview = $preview.Substring(0, [Math]::Max(0, $width - 33)) + "..."
        }
        [void]$contentLines.Add("$prefix$($f.Label):  $preview")
    }

    for ($l = 0; $l -lt $availableLines; $l++) {
        if ($l -lt $contentLines.Count) {
            $lineText = $contentLines[$l]
            $padded = $lineText + (" " * [Math]::Max(0, $width - $lineText.Length))
            if ($l -ge $fieldStartLine -and ($l - $fieldStartLine) -eq $SelectedIndex) {
                Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
            }
            else {
                Write-Host $padded -ForegroundColor White
            }
        }
        else {
            Write-Host (" " * $width)
        }
    }

    $statusLine = " Choose a field "
    $padded3 = $statusLine + (" " * [Math]::Max(0, $width - $statusLine.Length))
    Write-Host $padded3 -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline
}

# ── Render status picker ───────────────────────────────────────────
function Render-StatusPicker {
    param(
        [hashtable]$Item,
        [array]$Statuses,
        [int]$SelectedStatusIndex
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " CHANGE STATUS "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkMagenta

    $helpLine = " [Up/Down] Select  [Enter] Confirm  [ESC] Cancel "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    $availableLines = $height - 3

    $contentLines = [System.Collections.ArrayList]::new()
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  $($Item.Type) #$($Item.Id): $($Item.Title)")
    [void]$contentLines.Add("  Current State: $($Item.State)")
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  Select new status:")
    [void]$contentLines.Add("")

    $statusStartLine = $contentLines.Count

    for ($i = 0; $i -lt $Statuses.Count; $i++) {
        $prefix = "    "
        if ($i -eq $SelectedStatusIndex) {
            $prefix = "  > "
        }
        [void]$contentLines.Add("$prefix$($Statuses[$i])")
    }

    for ($l = 0; $l -lt $availableLines; $l++) {
        if ($l -lt $contentLines.Count) {
            $lineText = $contentLines[$l]
            $padded = $lineText + (" " * [Math]::Max(0, $width - $lineText.Length))
            if ($l -ge $statusStartLine -and ($l - $statusStartLine) -eq $SelectedStatusIndex) {
                Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
            }
            else {
                Write-Host $padded -ForegroundColor White
            }
        }
        else {
            Write-Host (" " * $width)
        }
    }

    $statusLine = " Choose a status "
    $padded3 = $statusLine + (" " * [Math]::Max(0, $width - $statusLine.Length))
    Write-Host $padded3 -ForegroundColor Gray -BackgroundColor DarkGray
}

# ── Render assignee picker ──────────────────────────────────────────
function Render-AssigneePicker {
    param(
        [hashtable]$Item,
        [string]$SearchText,
        [array]$Results,
        [int]$SelectedIndex   # 0 = "(Unassign)", 1..n = Results entries
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " ASSIGN WORK ITEM "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkMagenta

    $helpLine = " [Type] Search  [Up/Down] Select  [Enter] Confirm  [Backspace] Delete  [ESC] Cancel "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    $availableLines = $height - 3

    $contentLines = [System.Collections.ArrayList]::new()
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  $($Item.Type) #$($Item.Id): $($Item.Title)")
    $currentAssignee = if ($Item.ContainsKey('AssignedTo') -and $Item.AssignedTo) { $Item.AssignedTo } else { "(unassigned)" }
    [void]$contentLines.Add("  Current Assignee: $currentAssignee")
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  Search: [$SearchText]_")
    [void]$contentLines.Add("")
    [void]$contentLines.Add("  Select new assignee:")
    [void]$contentLines.Add("")

    $optionsStartLine = $contentLines.Count

    # Option 0: Unassign
    $prefix0 = if ($SelectedIndex -eq 0) { "  > " } else { "    " }
    [void]$contentLines.Add("$prefix0(Unassign)")

    # Options 1..n: search results
    for ($i = 0; $i -lt $Results.Count; $i++) {
        $prefix = if (($i + 1) -eq $SelectedIndex) { "  > " } else { "    " }
        [void]$contentLines.Add("$prefix$($Results[$i])")
    }

    # Hint when no results yet
    if ($Results.Count -eq 0 -and $SearchText.Length -ge 2) {
        [void]$contentLines.Add("    (no matches - try a different name)")
    }
    elseif ($Results.Count -eq 0 -and $SearchText.Length -eq 1) {
        [void]$contentLines.Add("    (type one more character to search...)")
    }
    elseif ($Results.Count -eq 0) {
        [void]$contentLines.Add("    (type 2+ characters to search for a user)")
    }

    for ($l = 0; $l -lt $availableLines; $l++) {
        if ($l -lt $contentLines.Count) {
            $lineText = $contentLines[$l]
            $padded = $lineText + (" " * [Math]::Max(0, $width - $lineText.Length))
            $lineOptIdx = $l - $optionsStartLine   # 0 = Unassign, 1..n = results
            $isSelected = ($l -ge $optionsStartLine) -and ($lineOptIdx -eq $SelectedIndex)
            if ($isSelected) {
                Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
            }
            else {
                Write-Host $padded -ForegroundColor White
            }
        }
        else {
            Write-Host (" " * $width)
        }
    }

    $statusLine = " Select '(Unassign)' to clear the assignee, or type 2+ chars to search "
    $padded3 = $statusLine + (" " * [Math]::Max(0, $width - $statusLine.Length))
    Write-Host $padded3 -ForegroundColor Gray -BackgroundColor DarkGray
}

# ── Interactive text input with live autocomplete suggestions ──────
function Read-WithSuggestions {
    param(
        [string]$Prompt,
        [string]$InitialValue        = "",
        [string]$Organization        = "",
        [string]$PAT                 = "",
        [int]$MinCharsForSuggestions = 3,
        [int]$MaxSuggestions         = 8
    )

    $text             = $InitialValue
    $suggestions      = @()
    $suggestionIndex  = -1
    $lastFetchedText  = ""

    $screenWidth      = [Console]::WindowWidth
    $screenHeight     = [Console]::WindowHeight
    $inputRow         = $screenHeight - 1          # last row
    $suggestAreaTop   = $inputRow - $MaxSuggestions - 1

    [Console]::CursorVisible = $true

    function DrawInput {
        param([string]$T)
        [Console]::SetCursorPosition(0, $inputRow)
        $line   = " $Prompt [$T]"
        $padded = $line + (" " * [Math]::Max(0, $screenWidth - $line.Length))
        Write-Host $padded -NoNewline -ForegroundColor White -BackgroundColor DarkCyan
        # Park the cursor right after the typed text
        $col = [Math]::Min(" $Prompt [".Length + $T.Length, $screenWidth - 1)
        [Console]::SetCursorPosition($col, $inputRow)
    }

    function DrawSuggestions {
        param([array]$Suggs, [int]$SelIdx)
        $count = [Math]::Min($Suggs.Count, $MaxSuggestions)
        # Clear the entire suggestion area first
        for ($r = $suggestAreaTop; $r -lt $inputRow; $r++) {
            [Console]::SetCursorPosition(0, $r)
            Write-Host (" " * $screenWidth) -NoNewline
        }
        if ($count -gt 0) {
            [Console]::SetCursorPosition(0, $suggestAreaTop)
            $hdr = " Suggestions  (↑↓ select, Enter confirm, Esc cancel) "
            $hdrPad = $hdr + (" " * [Math]::Max(0, $screenWidth - $hdr.Length))
            Write-Host $hdrPad -NoNewline -ForegroundColor Cyan -BackgroundColor DarkGray
            for ($i = 0; $i -lt $count; $i++) {
                $row = $suggestAreaTop + 1 + $i
                if ($row -ge $inputRow) { break }
                [Console]::SetCursorPosition(0, $row)
                $lbl    = "  $($Suggs[$i])"
                $padded = $lbl + (" " * [Math]::Max(0, $screenWidth - $lbl.Length))
                if ($i -eq $SelIdx) {
                    Write-Host $padded -NoNewline -ForegroundColor White -BackgroundColor DarkBlue
                } else {
                    Write-Host $padded -NoNewline -ForegroundColor Gray
                }
            }
        }
    }

    # Initial draw
    DrawSuggestions -Suggs @() -SelIdx -1
    DrawInput -T $text

    while ($true) {
        # Refresh suggestions whenever the typed text changes
        if ($text.Length -ge $MinCharsForSuggestions -and $text -ne $lastFetchedText) {
            $suggestions     = @(Search-AzDoUsers -Organization $Organization -PAT $PAT -SearchTerm $text)
            $lastFetchedText = $text
            $suggestionIndex = -1
        } elseif ($text.Length -lt $MinCharsForSuggestions -and $suggestions.Count -gt 0) {
            $suggestions     = @()
            $lastFetchedText = ""
            $suggestionIndex = -1
        }

        DrawSuggestions -Suggs $suggestions -SelIdx $suggestionIndex
        DrawInput -T $text

        $key = [Console]::ReadKey($true)

        switch ($key.Key) {
            'Enter' {
                [Console]::CursorVisible = $false
                if ($suggestionIndex -ge 0 -and $suggestionIndex -lt $suggestions.Count) {
                    return $suggestions[$suggestionIndex]
                }
                return $text
            }
            'Escape' {
                [Console]::CursorVisible = $false
                return $null   # signals "cancelled"
            }
            'UpArrow' {
                if ($suggestions.Count -gt 0) {
                    if ($suggestionIndex -gt 0)  { $suggestionIndex-- }
                    elseif ($suggestionIndex -eq 0) { $suggestionIndex = -1 }
                }
            }
            'DownArrow' {
                if ($suggestions.Count -gt 0) {
                    $maxIdx = [Math]::Min($suggestions.Count, $MaxSuggestions) - 1
                    if ($suggestionIndex -lt $maxIdx) { $suggestionIndex++ }
                }
            }
            'Backspace' {
                if ($text.Length -gt 0) {
                    $text = $text.Substring(0, $text.Length - 1)
                    $suggestionIndex = -1
                    if ($text.Length -lt $MinCharsForSuggestions) {
                        $suggestions     = @()
                        $lastFetchedText = ""
                    }
                }
            }
            'Delete' {
                $text            = ""
                $suggestions     = @()
                $lastFetchedText = ""
                $suggestionIndex = -1
            }
            default {
                if ($key.KeyChar -and -not [char]::IsControl($key.KeyChar)) {
                    $text            += $key.KeyChar
                    $suggestionIndex  = -1
                }
            }
        }
    }
}

# ── Render query form ───────────────────────────────────────────────
function Render-QueryForm {
    param(
        [hashtable]$QueryData,
        [array]$TabNames = @(),
        [int]$ActiveTabIndex = -1,
        [string]$StatusMessage = ""
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " AZURE DEVOPS TIME TRACKER "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkBlue

    # Tab bar
    if ($TabNames.Count -gt 0 -and $ActiveTabIndex -ge 0) {
        Render-TabBarLine -TabNames $TabNames -ActiveTabIndex $ActiveTabIndex -Width $width
    }

    # Help
    $helpLine = " [Up/Down] Navigate  [Enter] Edit field / Search  [ESC] Back  [Tab] Switch tab "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    # Form content
    Write-Host ""
    Write-Host "  SEARCH WORK ITEMS" -ForegroundColor Cyan
    Write-Host ""

    $formFields = @(
        @{ Label = "Title contains"; Value = $QueryData.TitleContains; Hint = "" }
        @{ Label = "State";          Value = $QueryData.State;         Hint = "e.g. Active, New, Closed, or blank for any" }
        @{ Label = "Type";           Value = $QueryData.Type;          Hint = "e.g. Bug, Task, User Story, or blank for any" }
        @{ Label = "Assigned to";    Value = $QueryData.AssignedTo;    Hint = "e.g. @me, a name, or blank for any" }
        @{ Label = "Work Item ID";   Value = $QueryData.WorkItemId;    Hint = "Search by specific ID (ignores other filters)" }
    )

    $maxLabelLen = ($formFields | ForEach-Object { $_.Label.Length } | Measure-Object -Maximum).Maximum

    for ($i = 0; $i -lt $formFields.Count; $i++) {
        $f = $formFields[$i]
        $prefix = if ($i -eq $QueryData.FormSelectedIndex) { " > " } else { "   " }
        $label = $f.Label.PadRight($maxLabelLen)
        $val = if ($f.Value) { $f.Value } else { "" }
        $lineText = "$prefix${label}:  [$val]"
        $padded = Format-FixedWidth -Text $lineText -Width $width
        if ($i -eq $QueryData.FormSelectedIndex) {
            Write-Host $padded -ForegroundColor White -BackgroundColor DarkCyan
        }
        else {
            Write-Host $padded -ForegroundColor White
        }
    }

    # Search button
    Write-Host ""
    $searchIdx = $formFields.Count
    $prefix = if ($QueryData.FormSelectedIndex -eq $searchIdx) { " > " } else { "   " }
    $searchLabel = "${prefix}>>> Run Search <<<"
    $padded = Format-FixedWidth -Text $searchLabel -Width $width
    if ($QueryData.FormSelectedIndex -eq $searchIdx) {
        Write-Host $padded -ForegroundColor White -BackgroundColor DarkGreen
    }
    else {
        Write-Host $padded -ForegroundColor Green
    }

    # Hint for selected field
    Write-Host ""
    if ($QueryData.FormSelectedIndex -lt $formFields.Count) {
        $hint = $formFields[$QueryData.FormSelectedIndex].Hint
        if ($hint) {
            Write-Host (Format-FixedWidth -Text "  $hint" -Width $width) -ForegroundColor DarkGray
        }
        else {
            Write-Host (" " * $width)
        }
    }
    else {
        Write-Host (Format-FixedWidth -Text "  Press Enter to execute the search" -Width $width) -ForegroundColor DarkGray
    }

    # Fill remaining lines
    $usedLines = 3 + 3 + $formFields.Count + 3 + 1  # header area (3) + blank/title/blank (3) + fields + blank/button/blank (3) + hint (1)
    $remaining = $height - $usedLines - 1
    for ($l = 0; $l -lt $remaining; $l++) {
        Write-Host (" " * $width)
    }

    # Status bar
    if ($StatusMessage) {
        $statusText = " $StatusMessage "
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor White -BackgroundColor DarkGreen -NoNewline
    }
    else {
        $statusText = " Fill in filters and press Enter on 'Run Search' "
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline
    }
}
# ── Render parent-item search form (Pri tab) ────────────────────────
function Render-PriSearchForm {
    param(
        [hashtable]$PriData,
        [array]$TabNames       = @(),
        [int]$ActiveTabIndex   = -1,
        [string]$StatusMessage = ""
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width  = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " AZURE DEVOPS TIME TRACKER "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkBlue

    # Tab bar
    if ($TabNames.Count -gt 0 -and $ActiveTabIndex -ge 0) {
        Render-TabBarLine -TabNames $TabNames -ActiveTabIndex $ActiveTabIndex -Width $width
    }

    # Help line
    $inResults = $PriData.FormState -eq 'results'
    $helpLine = if ($inResults) {
        " [Up/Down] Select parent  [Enter] Confirm  [ESC] Back to search  [Tab] Switch tab "
    }
    else {
        " [Type] Work item ID or title  [Enter] Fetch/Search  [ESC] Cancel  [Tab] Switch tab "
    }
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    # Section title
    Write-Host ""
    Write-Host "  PARENT WORK ITEM" -ForegroundColor Cyan
    Write-Host ""

    $usedLines = 6  # header(1) + tabbar(1) + helpline(1) + blank(1) + title(1) + blank(1)

    # Current parent info
    if ($PriData.HasParent) {
        $parentLine = "  Current: #$($PriData.ParentId) - $($PriData.ParentTitle)"
        Write-Host (Format-FixedWidth -Text $parentLine -Width $width) -ForegroundColor Green
        Write-Host ""
        $usedLines += 2
    }

    # Search input line
    $inputLabel = "  Search (ID or title): "
    $inputVal   = if ($PriData.SearchInput) { $PriData.SearchInput } else { "" }
    $inputLine  = "$inputLabel[$inputVal]_"
    $paddedInput = Format-FixedWidth -Text $inputLine -Width $width
    if (-not $inResults) {
        Write-Host $paddedInput -ForegroundColor White -BackgroundColor DarkCyan
    }
    else {
        Write-Host $paddedInput -ForegroundColor White
    }
    $usedLines++

    # Results list
    if ($inResults -and $PriData.SearchResults.Count -gt 0) {
        Write-Host ""
        Write-Host "  Results ($($PriData.SearchResults.Count) found):" -ForegroundColor Gray
        $usedLines += 2

        $maxResults = $height - $usedLines - 2  # leave 2 lines: blank+status
        if ($maxResults -lt 1) { $maxResults = 1 }

        for ($i = 0; $i -lt [Math]::Min($PriData.SearchResults.Count, $maxResults); $i++) {
            $r = $PriData.SearchResults[$i]
            $icon = Get-TypeIcon -Type $r.Type
            $prefix = if ($i -eq $PriData.SearchResultIndex) { "  > " } else { "    " }
            $stateStr = if ($r.State) { "[$($r.State)] " } else { "" }
            $lineText = "$prefix$icon #$($r.Id) $stateStr$($r.Title)"
            $padded2 = Format-FixedWidth -Text $lineText -Width $width
            if ($i -eq $PriData.SearchResultIndex) {
                Write-Host $padded2 -ForegroundColor White -BackgroundColor DarkCyan
            }
            else {
                $typeColor = Get-TypeColor -Type $r.Type
                Write-Host $padded2 -ForegroundColor $typeColor
            }
            $usedLines++
        }
    }

    # Fill remaining lines
    $remaining = $height - $usedLines - 2
    for ($l = 0; $l -lt $remaining; $l++) {
        Write-Host (" " * $width)
    }

    # Status bar
    if ($StatusMessage) {
        $statusText = " $StatusMessage "
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor White -BackgroundColor DarkGreen -NoNewline
    }
    else {
        $statusText = if ($PriData.HasParent) {
            " Parent: #$($PriData.ParentId) $($PriData.ParentTitle) - type to search for a different parent "
        }
        else {
            " Enter a work item ID (e.g. 12345) or part of a title, then press Enter "
        }
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline
    }
}

# ── Copy text to system clipboard ──────────────────────────────────
function Set-TTClipboard {
    param([string]$Text)
    try {
        Set-Clipboard -Value $Text -ErrorAction Stop
        return $true
    } catch {}
    # Fallback for Linux
    if ($IsLinux) {
        foreach ($tool in @('xclip', 'xsel', 'wl-copy')) {
            if (Get-Command $tool -ErrorAction SilentlyContinue) {
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
                    $psi = [System.Diagnostics.ProcessStartInfo]::new($tool)
                    if ($tool -eq 'xclip') { $psi.Arguments = '-selection clipboard' }
                    elseif ($tool -eq 'xsel') { $psi.Arguments = '--clipboard --input' }
                    $psi.RedirectStandardInput = $true
                    $psi.UseShellExecute = $false
                    $psi.CreateNoWindow = $true
                    $p = [System.Diagnostics.Process]::Start($psi)
                    $p.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
                    $p.StandardInput.Close()
                    $p.WaitForExit(5000)
                    if ($p.ExitCode -eq 0) { return $true }
                } catch {}
            }
        }
    }
    elseif ($IsMacOS) {
        try {
            $psi = [System.Diagnostics.ProcessStartInfo]::new('pbcopy')
            $psi.RedirectStandardInput = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $p = [System.Diagnostics.Process]::Start($psi)
            $p.StandardInput.Write($Text)
            $p.StandardInput.Close()
            $p.WaitForExit(5000)
            return $true
        } catch {}
    }
    return $false
}

# ── Build plain text scrum report ──────────────────────────────────
function Build-ScrumReportText {
    param([hashtable]$ReportData)

    $sb = [System.Text.StringBuilder]::new()

    [void]$sb.AppendLine("** What I did yesterday")
    [void]$sb.AppendLine()

    if ($ReportData.ClosedItems.Count -gt 0) {
        [void]$sb.AppendLine("* Closed/Resolved:")
        foreach ($item in $ReportData.ClosedItems) {
            [void]$sb.AppendLine("  - [$($item.Type)] #$($item.Id): $($item.Title)")
        }
        [void]$sb.AppendLine()
    }

    if ($ReportData.CommentedItems.Count -gt 0) {
        [void]$sb.AppendLine("* Commented on:")
        foreach ($item in $ReportData.CommentedItems) {
            [void]$sb.AppendLine("  - [$($item.Type)] #$($item.Id): $($item.Title)")
        }
        [void]$sb.AppendLine()
    }

    if ($ReportData.EditedItems.Count -gt 0) {
        [void]$sb.AppendLine("* Edited/Updated hours:")
        foreach ($item in $ReportData.EditedItems) {
            [void]$sb.AppendLine("  - [$($item.Type)] #$($item.Id): $($item.Title)")
        }
        [void]$sb.AppendLine()
    }

    $yesterdayCount = $ReportData.ClosedItems.Count + $ReportData.CommentedItems.Count + $ReportData.EditedItems.Count
    if ($yesterdayCount -eq 0) {
        [void]$sb.AppendLine("  (no activity found)")
        [void]$sb.AppendLine()
    }

    [void]$sb.AppendLine("** What I am going to do today")
    [void]$sb.AppendLine()

    if ($ReportData.FollowUpItems.Count -gt 0) {
        [void]$sb.AppendLine("* Follow up (unanswered comments/mentions):")
        foreach ($item in $ReportData.FollowUpItems) {
            [void]$sb.AppendLine("  - [$($item.Type)] #$($item.Id): $($item.Title)")
        }
        [void]$sb.AppendLine()
    }

    if ($ReportData.ActiveItems.Count -gt 0) {
        [void]$sb.AppendLine("* Active column on my Teams Boards:")
        foreach ($item in $ReportData.ActiveItems) {
            [void]$sb.AppendLine("  - [$($item.Type)] #$($item.Id): $($item.Title)")
        }
        [void]$sb.AppendLine()
    }

    $todayCount = $ReportData.FollowUpItems.Count + $ReportData.ActiveItems.Count
    if ($todayCount -eq 0) {
        [void]$sb.AppendLine("  (no items found)")
        [void]$sb.AppendLine()
    }

    return $sb.ToString().TrimEnd()
}

# ── Render SCRUM report view ──────────────────────────────────────
function Render-ScrumReport {
    param(
        [array]$Lines,
        [int]$ScrollOffset,
        [array]$TabNames = @(),
        [int]$ActiveTabIndex = -1,
        [string]$StatusMessage = ""
    )

    [Console]::CursorVisible = $false
    [Console]::SetCursorPosition(0, 0)

    $width = [Console]::WindowWidth
    $height = [Console]::WindowHeight

    # Header
    $header = " AZURE DEVOPS TIME TRACKER "
    $padLen = [Math]::Max(0, $width - $header.Length)
    Write-Host ($header + (" " * $padLen)) -ForegroundColor White -BackgroundColor DarkBlue

    # Tab bar
    $tabBarLines = 0
    if ($TabNames.Count -gt 0 -and $ActiveTabIndex -ge 0) {
        $tabBarLines = 1
        Render-TabBarLine -TabNames $TabNames -ActiveTabIndex $ActiveTabIndex -Width $width
    }

    # Help bar
    $helpLine = " [Arrows/PgUp/PgDn] Scroll  [c] Copy to clipboard  [r] Refresh  [Tab] Switch tab  [q] Quit "
    $padLen2 = [Math]::Max(0, $width - $helpLine.Length)
    Write-Host ($helpLine + (" " * $padLen2)) -ForegroundColor Gray -BackgroundColor DarkGray

    # Available lines for content
    $availableLines = $height - 3 - $tabBarLines

    for ($l = 0; $l -lt $availableLines; $l++) {
        $lineIdx = $ScrollOffset + $l
        if ($lineIdx -lt $Lines.Count) {
            $lineText = $Lines[$lineIdx]
            $padded = Format-FixedWidth -Text "  $lineText" -Width $width
            # Color coding based on content
            if ($lineText -match '^\*\*\s') {
                Write-Host $padded -ForegroundColor Cyan
            }
            elseif ($lineText -match '^\*\s') {
                Write-Host $padded -ForegroundColor Yellow
            }
            elseif ($lineText -match '^\s+-\s\[') {
                Write-Host $padded -ForegroundColor White
            }
            elseif ($lineText -match '^\s+\(no ') {
                Write-Host $padded -ForegroundColor DarkGray
            }
            else {
                Write-Host $padded -ForegroundColor Gray
            }
        }
        else {
            Write-Host (" " * $width)
        }
    }

    # Status bar
    if ($StatusMessage) {
        $statusText = " $StatusMessage "
        $statusPadded = $statusText + (" " * [Math]::Max(0, $width - $statusText.Length))
        Write-Host $statusPadded -ForegroundColor White -BackgroundColor DarkGreen -NoNewline
    }
    else {
        $scrollInfo = " Line $($ScrollOffset + 1) of $($Lines.Count)  |  Press [c] to copy report to clipboard "
        $statusPadded = $scrollInfo + (" " * [Math]::Max(0, $width - $scrollInfo.Length))
        Write-Host $statusPadded -ForegroundColor Gray -BackgroundColor DarkGray -NoNewline
    }

    return $ScrollOffset
}