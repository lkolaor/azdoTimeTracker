# AzDoTimeTracker

An interactive PowerShell module for tracking time against Azure DevOps work items.
Please note that this whole thing is vibe coded, use at your own peril.

## Features

- Displays all work items (Epics, Features, User Stories, Tasks, Bugs, Incidents) assigned to you
- Hierarchical view showing parent-child relationships
- Keyboard-driven navigation
- View work item details and comments inline
- Built-in stopwatch for time tracking
- Automatically updates Completed Work and Remaining Work fields in Azure DevOps
- Full-screen text editor for comments and field editing
- Tools menu for reconfiguring settings

## Requirements

- **PowerShell 7.0+** (cross-platform)
- **Azure DevOps Personal Access Token (PAT)** with **Work Items (Read & Write)** scope

## Installation

### From source (local install)

Clone or copy the module folder, then import it:

```powershell
# Option 1: Import directly from the source folder
Import-Module /path/to/AzDoTimeTracker

# Option 2: Copy to your PowerShell modules directory for auto-discovery
$dest = Join-Path ($env:PSModulePath -split ':')[0] 'AzDoTimeTracker'
Copy-Item -Path /path/to/AzDoTimeTracker -Destination $dest -Recurse
Import-Module AzDoTimeTracker
```

### Symlink into modules path (recommended for development)

```powershell
$modulesDir = ($env:PSModulePath -split ':')[0]
if (-not (Test-Path $modulesDir)) { New-Item -ItemType Directory -Path $modulesDir -Force }
New-Item -ItemType SymbolicLink -Path (Join-Path $modulesDir 'AzDoTimeTracker') -Target /path/to/AzDoTimeTracker
```

Then simply:

```powershell
Import-Module AzDoTimeTracker
Start-TimeTracker
```

## Setup

### 1. Create a Personal Access Token

1. Go to `https://dev.azure.com/{your-org}/_usersSettings/tokens`
2. Click **New Token**
3. Give it a name (e.g. "Time Tracker")
4. Set the scope to **Work Items → Read & Write**
5. Copy the generated token

### 2. Run the application

```powershell
Import-Module AzDoTimeTracker
Start-TimeTracker
```

On first run, you'll be prompted to enter:
- **Organization** – your Azure DevOps org name (from `https://dev.azure.com/{org}`)
- **Project** – the project name
- **PAT** – the token you created above

Configuration is saved to:
- **Linux/macOS**: `~/.config/AzDoTimeTracker/config.json`
- **Windows**: `%APPDATA%\AzDoTimeTracker\config.json`

To force reconfiguration:
```powershell
Start-TimeTracker -Reconfigure
```

## Controls

| Key | Action |
|------|--------|
| `↑` / `↓` | Navigate the work item list |
| `Page Up` / `Page Down` | Scroll by page |
| `Home` / `End` | Jump to first/last item |
| `Enter` | View item details (description + comments) |
| `t` | Start/stop time tracking on selected item (on User Stories, creates a child Task first) |
| `r` | Refresh work item list from Azure DevOps |
| `m` | Open Tools menu |
| `q` | Quit (saves all active timers) |

### Detail View

| Key | Action |
|------|--------|
| `ESC` | Go back to list |
| `↑` / `↓` | Scroll |
| `h` | Edit hours (Original Estimate, Completed, Remaining) |
| `s` | Change status |
| `f` | Edit fields (title, description, etc.) |
| `a` | Add comment |
| `e` | Edit comment |
| `d` | Delete comment |

### Text Editor

| Key | Action |
|------|--------|
| `Ctrl+S` | Save |
| `Ctrl+Q` | Cancel |

### Tools Menu

Press `m` from the list view to open the Tools menu. Available options:

| Option | Description |
|--------|-------------|
| Reconfigure | Re-enter organization, project, and PAT |
| Delete selected Task | Permanently deletes the currently selected Task (Tasks only — requires confirmation) |
| View debug log | Shows the last 50 lines of the debug log |
| View README | Browse this README inside the TUI |
| About | Show version and module info |

#### Deleting a Task

1. Select the Task you want to delete in the list
2. Press `m` to open the Tools menu
3. Choose **Delete selected Task**
4. A confirmation screen shows the task ID and title
5. Press `y` to confirm — the task is moved to the Azure DevOps recycle bin
6. Any other key cancels the operation

> Only **Task** type items can be deleted through this menu. Selecting any other type (User Story, Bug, etc.) will show an error and take no action.

## Time Tracking

1. Select a work item that has time fields (Original Estimate, Completed Work, Remaining Work)
2. Press `t` to start the timer
3. Work on your task — the elapsed time is displayed live
4. Press `t` again to stop the timer and save
5. The elapsed time is added to **Completed Work** and subtracted from **Remaining Work**
6. Changes are saved to Azure DevOps automatically

> Items without time tracking fields (e.g., Features, Epics) will show a message that time tracking is not supported.

### Time Tracking on User Stories

User Stories do not carry time tracking fields themselves. When you press `t` on a User Story, the tracker automatically:

1. Creates a new **Task** child item with the title `<ID> <User Story title>` (e.g. `143748 Implement login page`)
2. Sets **Original Estimate** and **Remaining Work** to **5 hours**
3. Assigns the task to the same person as the User Story and sets its state to **Active**
4. Links it as a child of the User Story in Azure DevOps
5. Selects the new task in the list and starts the timer on it immediately

## Work Item Display

Items are shown with type-specific icons:

- ★ Epic
- ◆ Feature
- ● User Story / Product Backlog Item
- ▶ Task
- ■ Bug
- ⚠ Incident / Issue

Parent items that aren't assigned to you are shown in gray for context.

## Module Structure

```
AzDoTimeTracker/
├── AzDoTimeTracker.psd1     # Module manifest
├── AzDoTimeTracker.psm1     # Root module (dot-sources all files)
├── Build-Package.ps1        # Builds a .nupkg for offline install
├── Public/
│   └── Start-TimeTracker.ps1  # Exported entry-point function
├── Private/
│   ├── Config.ps1            # Configuration management
│   ├── AzureDevOps.ps1       # Azure DevOps REST API functions
│   └── UI.ps1                # Terminal UI rendering
└── README.md
```

## Exported Commands

| Command | Description |
|---------|-------------|
| `Start-TimeTracker` | Launch the interactive time tracker TUI |

### Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Reconfigure` | Switch | Force re-entry of organization, project, and PAT |

## Publishing to PowerShell Gallery

To publish this module to the PowerShell Gallery:

```powershell
# Update ProjectUri and LicenseUri in AzDoTimeTracker.psd1 first
Publish-Module -Path ./AzDoTimeTracker -NuGetApiKey $apiKey
```

## Building a .nupkg for Offline Install

The included `Build-Package.ps1` script creates a `.nupkg` file — the same
format used by the PowerShell Gallery — so the module can be shared and
installed offline with `Install-Module`.

### Build the package

```powershell
# From the module directory
./Build-Package.ps1

# Or specify a custom output directory
./Build-Package.ps1 -OutputDir ~/Desktop
```

This creates `out/AzDoTimeTracker.<version>.nupkg`.

### Install from the .nupkg

**Option 1: Register a local repository** (recommended)

```powershell
# Point a local repo at the folder containing the .nupkg
Register-PSRepository -Name Local -SourceLocation /path/to/out -InstallationPolicy Trusted

# Install the module
Install-Module -Name AzDoTimeTracker -Repository Local

# Clean up the repo registration when done
Unregister-PSRepository -Name Local
```

**Option 2: Extract manually**

A `.nupkg` is just a zip file. You can extract it directly into your modules path:

```powershell
$dest = Join-Path ($env:PSModulePath -split '[;:]')[0] 'AzDoTimeTracker'
Expand-Archive -Path out/AzDoTimeTracker.1.0.0.nupkg -DestinationPath $dest -Force
```

After installation:

```powershell
Import-Module AzDoTimeTracker
Start-TimeTracker
```

### Uninstall

```powershell
# If installed via Install-Module
Uninstall-Module AzDoTimeTracker

# Or remove manually
Remove-Item -Recurse (Join-Path ($env:PSModulePath -split '[;:]')[0] 'AzDoTimeTracker')
```
