# M365 Package Manager

A complete PowerShell GUI application for managing, removing, and reinstalling Microsoft 365 (M365) applications, Office components, Teams integrations, and Zivver software on Windows computers.

## 📋 Overview

This tool enables IT administrators and end users to:
- scan and display Microsoft 365 and Office applications
- detect and remove Teams add-ins (Office Outlook add-in)
- discover and clean up Zivver applications and configurations
- remove selected M365 components (ODT-supported)
- clean Teams cache and log files
- fully reinstall M365/Office
- manually reinstall Teams Add-in
- manually reinstall Zivver application
- get detailed operation output

## 🚀 Quick start

### Requirements
- **PowerShell 5.1+** (Windows PowerShell)
- **Administrator rights** (the script requests elevation automatically)
- **Windows 10/11**
- **Internet access** (for Office Deployment Tool download)

### Run the script

```powershell
.\M365-Package-Manager.ps1
```

The script will request admin elevation automatically when needed.

## 📖 User guide

### Main interface

The application uses a WPF-based graphical interface with:

1. **Application list** - detailed DataGrid of discovered entries:
   - checkbox per item
   - display name (for example `Microsoft 365 Apps for enterprise`)
   - version
   - source
   - uninstall command
   - type (M365, Teams, Zivver)

2. **Status area** - live execution status:
   - status message
   - progress bar for long-running actions

3. **Action buttons**:
   - **🔍 Scan M365** – detect M365/Office applications
   - **🔍 Scan Teams Addin** – detect Teams Outlook Add-in
   - **🔍 Scan Zivver** – detect Zivver software
   - **🗑️ Remove Selected** – delete selected entries
   - **🧹 Clear Teams Cache** – clean Teams files
   - **🔄 Reinstall M365** – reinstall M365
   - **🔧 Reinstall Teams Addin** – repair Teams Outlook add-in
   - **🔧 Reinstall Zivver** – repair Zivver application
   - **🗑️ Remove All** – full M365 removal

### Step 1: Scan applications

1. Click **"🔍 Scan M365"** to discover installed M365/Office applications.
2. The DataGrid lists each result with version and source.
3. Optionally scan for specific groups:
   - **Teams Addin** – Teams Outlook integration
   - **Zivver** – e-mail security software

### Step 2: Select items

- Select items with the checkboxes
- Select multiple entries for bulk removal
- Only selected items are eligible for removal

### Step 3: Remove entries

1. Select one or more entries in the DataGrid
2. Click **"🗑️ Remove Selected"**
3. Follow the on-screen status updates
4. Status is refreshed live

### Step 4: Clean cache

1. Click **"🧹 Clear Teams Cache"** to remove Teams cache and logs
2. This removes:
   - local Teams cache (`AppData\Local\Microsoft\Teams`)
   - roaming Teams data (`AppData\Roaming\Microsoft\Teams`)
   - logs and temporary files

### Step 5: Reinstall

#### Reinstall M365
1. Click **"🔄 Reinstall M365"**
2. The script downloads Office Deployment Tool (ODT) if required
3. It creates a deployment config containing:
   - bitness based on system architecture
   - English (`en-us`) and Dutch (`nl-nl`) language packs
   - default update channel
4. Runs installation

#### Reinstall Teams Add-in
1. Click **"🔧 Reinstall Teams Addin"**
2. The script detects Outlook installation files
3. Registers Outlook add-in registry keys

#### Reinstall Zivver
1. Click **"🔧 Reinstall Zivver"**
2. Detects available Zivver install files
3. Runs the install workflow

### Step 6: Full removal

1. Click **"🗑️ Remove All"**
2. The script will:
   - download ODT if missing
   - create a removal configuration
   - remove all M365 components
   - perform cleanup cleanup steps

## 🔧 Technical architecture

### Core components

#### `AppItem` class
```powershell
class AppItem {
    [string]$DisplayName      # display name
    [string]$Version          # version
    [string]$Source           # source path/identifier
    [string]$UninstallString  # uninstall command
    [string]$Type             # type: M365, Teams, Zivver
    [bool]$IsSelected         # selection state
}
```

#### Scan functions

**Find-M365Item**
- scans Windows uninstall registry keys for M365/Office entries
- filters by Office/M365 related keys and excludes Zivver entries

**Find-TeamsAddin**
- checks per-user registry keys
- scans user AppData for add-in folders
- excludes Zivver locations

**Find-ZivverItem**
- scans APPDATA, LOCALAPPDATA, ProgramData and known cache/log/config paths
- includes `C:\Program Files (x86)\Zivver B.V.` if present

#### Removal functions

**Remove-M365Item**
- stops dependent processes (Teams, Outlook, Excel, Word, etc.)
- executes ODT-based uninstall
- removes stale files if required
- cleans registry leftovers where possible

**Remove-FilesInFolder**
- attempts process termination
- removes files/folders recursively
- retries when files are locked
- safe cleanup of in-use items

**Clear-TeamsCache**
- scans all user profiles
- removes Teams cache and logs
- closes Teams safely first

#### Reinstall logic

**Reinstall-TeamsAddin-Logic**
- locates Outlook installation
- registers Teams Add-in keys
- runs Outlook repair/reload when needed

**Reinstall-Zivver-Logic**
- discovers installer payloads
- executes Zivver setup flow
- validates installation state where possible

### Data flow

```
GUI (WPF)
  └─ Scan-M365Item / Scan-TeamsAddin / Scan-ZivverItem
       └─ AppList (ObservableCollection[AppItem])
            └─ DataGrid binding refresh

  └─ Remove selected entries
       ├─ Stop-ProcessByName
       ├─ Remove-M365Item / Remove-PathSafe
       ├─ Remove-FilesInFolder
       └─ UI status refresh via Dispatcher

  └─ Reinstall actions
       ├─ ODT download when needed
       ├─ config file generation
       └─ setup.exe execution
```

### Threading model

- Scan and install actions run asynchronously with background jobs
- UI updates are marshaled through `$window.Dispatcher.Invoke()`
- button states and list refreshes are synchronized on UI thread

## 🔒 Safety / reliability

- automatic admin elevation
- process termination with safety checks
- retry handling for locked files
- registry error handling
- detailed status output

## 📝 Logging

Execution output appears in the PowerShell host:
- scan progress
- per-app removal status
- ODT download progress
- install/reinstall result lines

Example output:
```
Microsoft 365 Apps for enterprise (version 2312) - Scan complete
Removal complete: Word (Microsoft 365)
Teams cache cleanup complete.
M365 reinstall started...
M365 reinstall completed.
```

## 📊 Supported operations

### Application categories
- **M365**: Microsoft 365 Apps and Office products
- **Teams**: Teams client and Teams add-ins
- **Zivver**: E-mail security integrations

### Removal methods
1. Registry-based removal (when available)
2. File-system cleanup
3. Process termination before deletion
4. ODT-based deployment tool removal flow

### Reinstall channels
- **Current** channel used by ODT configuration
- Multi-language setup enabled for `en-us` and `nl-nl`

## ⚙️ Configuration

### Office Deployment Tool

The script auto-downloads ODT from Microsoft when missing:
- URL: `https://www.microsoft.com/en-us/download/details.aspx?id=49117`
- Extracted to: `C:\ODT`
- Logs: `C:\ODT\Logs`

### Customizable settings

Tune these variables in the script as needed:

```powershell
# scan filters
$m365Keywords = @('Microsoft 365', 'Office')

# custom Zivver paths
$zivverPaths = @(
    'C:\Program Files (x86)\Zivver B.V',
    "$env:APPDATA\Zivver"
    # add more locations as needed
)

# processes to stop before removal
$stopProcesses = @('Teams', 'Outlook', 'WINWORD', 'EXCEL', 'POWERPNT')
```

## 🐛 Troubleshooting

### "Administrator rights are required"
- Run PowerShell as Administrator
- Check UAC policy

### ODT download fails
- Verify internet access and proxy
- Disable blocking firewall rules temporarily
- Download ODT manually and place it in `C:\ODT`

### Item not found during scan
- The app may not be installed
- Run a new scan
- Check uninstall keys manually:
  - `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`
  - `HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall`

### Removal fails
- Close the application manually (for example Teams/Outlook)
- Run the script again with admin rights
- Verify folder permissions
- Try Safe Mode with Networking if needed

### Teams Add-in reinstall fails
- Ensure Outlook is installed
- Restart Outlook after reinstall
- Validate registry keys under `HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins`

## 📋 Feature matrix

| Feature | Supported | Notes |
|---------|---|---|
| M365 scan | ✅ Yes | Registry + WOW6432Node |
| Teams Add-in scan | ✅ Yes | Outlook integration detection |
| Zivver scan | ✅ Yes | Multi-location search |
| Remove selected items | ✅ Yes | ODT-assisted removal |
| Cache cleanup | ✅ Yes | Teams data removal |
| Reinstall | ✅ Yes | ODT based |
| Logging | ✅ Yes | PowerShell host output |
| WPF GUI | ✅ Yes | Full user interface |
| Async operations | ✅ Yes | Non-blocking background jobs |
| Bulk operations | ✅ Yes | Multi-select support |

## 📞 Support

If you run into issues:
1. Review the PowerShell host output for exact errors
2. Re-run scans to verify current state
3. Check Windows Firewall / security policies
4. Run as Administrator
5. Remove manually via Control Panel if automatic flow fails

## 📅 Version history

### Current (Dec 2025)
- ✅ WPF GUI implementation
- ✅ M365/Office scan support
- ✅ Teams Add-in detection
- ✅ Zivver management
- ✅ Async scan operations
- ✅ ODT based reinstall flow
- ✅ Cache and logs cleanup
- ✅ Detailed status updates

### Earlier versions
- Basic M365 removal actions
- manual registry scan
- single-item removals

## 📄 License / attribution

Built for internal Microsoft 365 device management use.

---

**Last updated:** March 16, 2026
**PowerShell version:** 5.1+
**ODT support:** Enabled
**Administrative rights required:** Yes
