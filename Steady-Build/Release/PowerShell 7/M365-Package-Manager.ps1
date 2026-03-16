##############################
# M365 Package Manager
# COMPLETE FIX – PART 1/5: INITIALIZATION & CLASSES
# PowerShell 5.1 Compatible
##############################

# --- Check whether the script is running as Administrator ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    try {
        # Use $PSCommandPath to get the current script path (works in PS 5.1).
        Start-Process pwsh.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    } catch {
        [System.Windows.Forms.MessageBox]::Show("This script requires Administrator rights to run.", "Permission error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    exit
}


# --- Add required GUI assemblies ---
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- Define the AppItem class ---
class AppItem {
    [string]$DisplayName
    [string]$Version
    [string]$Source
    [string]$UninstallString
    [string]$Type
    [bool]$IsSelected
}

# --- ObservableCollection for items (data binding) ---
$appList = [System.Collections.ObjectModel.ObservableCollection[AppItem]]::new()

##############################
# PART 2/5: FUNCTION DEFINITIONS
##############################

# ==========================================
# WPF HELPER FUNCTIONS
# ==========================================

# NOTE: The variables $dgApps, $lblStatus, etc. do not exist yet here,
# but functions must be defined before they are called.

# WPF DataGrid refresh function
function Update-DataGridWPF {
    param(
        [System.Collections.ArrayList]$DataList
    )
    # Use the WPF DataGrid variable ($dgApps) and the Dispatcher
    # to access the UI thread.
    if ($null -ne $dgApps -and $null -ne $dgApps.Dispatcher) {
        $dgApps.Dispatcher.Invoke([action]{
            $dgApps.ItemsSource = $null
            $dgApps.ItemsSource = $DataList
        })
    }
}

function Stop-ProcessByName {
    param(
        [string]$procName
    )

    try {
        $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($proc) {
            Stop-Process -Name $procName -Force -ErrorAction SilentlyContinue
            Write-Host "Proces $procName gestopt."
        }
    } catch {
        # Use string concatenation to format an error message
        Write-Host ("Fout bij stoppen van proces " + $procName + ": " + $_.Exception.Message)
    }
}

function Write-WPFStatus {
    param(
        [string]$StatusText,
        [int]$ProgressValue = 0
    )
    
    # Gets the global $window variable that contains the WPF elements
    $window = Get-Variable -Name 'window' -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    
    # If no WPF window is available, use Write-Host as fallback
    if (-not $window) {
        Write-Host "Geen WPF-venster beschikbaar om de status bij te werken. Huidige status: $StatusText"
        return
    }

    # FIX: Filter unwanted MSI output or other strings (for example '[INT]75')
    $FilteredText = $StatusText -replace "\[INT\]\d+" -replace "Download: \d+% voltooid\.\.\." 
    $FilteredText = $FilteredText.Trim()

    # Only update when meaningful text remains
    if ([string]::IsNullOrWhiteSpace($FilteredText)) {
        return # Sla de update over als de tekst leeg is na filteren
    }
    
    # If a WPF window exists, update status text and progress
    $window.Dispatcher.Invoke([action]{
        # Find WPF elements (must be defined in your GUI code)
        $txtStatus = $window.FindName("txtStatus")
        $prgStatus = $window.FindName("prgStatus")

        if ($txtStatus) {
            $txtStatus.Text = $FilteredText # GEBRUIK DE GEFILTERDE TEKST
        }
        
        if ($prgStatus) {
            $prgStatus.Value = $ProgressValue
        }
    })
}

# Function to display status in the WPF window (replaces erroneous Write-Host usage).
function Write-ConsoleStatus {
    param(
        [string]$StatusText,
        [int]$ProgressValue = -1
    )
    
    if ($ProgressValue -ge 0) {
        Write-Host "$StatusText - Progress: $ProgressValue%"
    } else {
        Write-Host "$StatusText"
    }
}

# Function to start asynchronous tasks and keep the UI responsive
function Start-AsyncJob {
    param(
        [ScriptBlock]$ScriptBlock,
        [string]$JobName,
        $ArgumentList,
        [System.Windows.Window]$window
    )

    # Start the background task
    $job = Start-Job -Name $JobName -ScriptBlock {
        param($sb, $argsList, $window)

        # Function to update status and progress
        function Write-WPFStatus {
            param([string]$StatusText, [int]$ProgressValue = -1)

            # Update the UI via Dispatcher to safely send progress to the GUI
            $window.Dispatcher.Invoke([action]{
                # Update the status label when available
                if ($window.FindName("lblStatus")) {
                    $lblStatus = $window.FindName("lblStatus")
                    $lblStatus.Text = $StatusText
                }

                # Update the progress bar when available
                if ($window.FindName("progressBar")) {
                    $progressBar = $window.FindName("progressBar")
                    $progressBar.Value = $ProgressValue
                }
            })
        }

        try {
            # Execute the task and send progress updates to the GUI
            & $sb @argsList
        } catch {
            # Send error message to the UI thread
            $window.Dispatcher.Invoke([action]{
                $window.Title = "Fout bij het uitvoeren van de taak: $($_.Exception.Message)"
            })
        }
    } -ArgumentList $ScriptBlock, $ArgumentList, $window

    return $job
}

# Function to re-enable buttons
function Enable-AllButtons {
    try {

        # Keep reinstall buttons enabled
        foreach ($btn in @($btnReinstallTeamsAddin,$btnReinstallZivver,$btnScanM365,$btnScanZivver)) {
            if ($btn) { 
                $btn.Dispatcher.Invoke([action]{ 
                    $btn.IsEnabled = $true
                    $btn.Visibility = [System.Windows.Visibility]::Visible
                })
            }
        }

        # Remove button depends on the current WPF selection
        if ($btnRemove -and $lstResults) {
            $btnRemove.Dispatcher.Invoke([action]{
                $btnRemove.IsEnabled = ($lstResults.SelectedItems.Count -gt 0)
                $btnRemove.Visibility = [System.Windows.Visibility]::Visible
            })
        }

    } catch {
        Write-Host "Fout bij het inschakelen van knoppen: $($_.Exception.Message)"
    }
}

# Function to find visual children (for header checkbox logic)
function Find-VisualChild {
    param(
        [System.Windows.DependencyObject]$parent,
        [Type]$childType
    )

    for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($parent); $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($parent, $i)

        if ($child -is $childType) {
            return $child
        }

        $result = Find-VisualChild -parent $child -childType $childType
        if ($result) { return $result }
    }
}

# --- Helper status update function ---
function Update-Status($text) {
	
    $window.Dispatcher.Invoke([action]{
        ($window.FindName("lblStatus")).Text = $text
    })
}

# --- Function: Remove-Path with retry ---
function Remove-PathSafe {
    param([string]$Path, [string]$DisplayName)
    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            $retry = 0
            while (Test-Path $Path -and $retry -lt 5) {
                Start-Sleep -Milliseconds 500
                Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
                $retry++
            }
            return !(Test-Path $Path)
        } catch {
            Write-WPFStatus "FOUT: Kon $DisplayName ($Path) niet verwijderen. ($($_.Exception.Message))"
            return $false
        }
    } else {
        return $true
    }
}

# ==========================================
# CORE LOGIC FUNCTIONS (unchanged from Part 1)
# ==========================================

function Find-M365Item {
    # Cache folders to scan for M365 (excluding Zivver paths)
    $cachePaths = @(
        "AppData\Local\Microsoft\Teams",
        "AppData\Roaming\Microsoft\Teams",
        "AppData\Local\Microsoft\TeamsMeetingAdd-in",
        "AppData\Local\Microsoft\TeamsMeetingAddinMsis",
        "AppData\Local\Microsoft\Office\16.0\OfficeFileCache",
        "AppData\Local\Microsoft\Outlook"
    )

    $users = Get-ChildItem "C:\Users" -Directory
    foreach ($user in $users) {
        foreach ($path in $cachePaths) {
            $fullPath = Join-Path $user.FullName $path
            if (Test-Path $fullPath) {
                $item = [AppItem]::new()
                $item.DisplayName = "M365 Cache: $path"
                $item.Version = "N/A"
                $item.Source = $fullPath
                $item.UninstallString = "Remove manually"
                $item.Type = "Cache"
                $item.IsSelected = $false
                $appList.Add($item)
            }
        }
    }

    Write-Host "`n--- M365 scan voltooid ---`n"
}

function Find-TeamsAddin {
    param([string[]]$Users)

    $results = @()

    foreach ($user in $Users) {
        # Old code used Get-LocalUser; now using a more robust SID lookup.
        $sid = $null
        try { $sid = (New-Object System.Security.Principal.NTAccount($user)).Translate([System.Security.Principal.SecurityIdentifier]).Value } catch {}
        
        if (-not $sid) {
             Write-Host "WAARSCHUWING: Kon geen SID vinden voor gebruiker: $user. Overslaan."
             continue
        }

        # Registry check for Teams add-in (Zivver excluded)
        $regKey = "HKU:\$sid\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
        if (Test-Path $regKey) {
            $results += [PSCustomObject]@{ User=$user; Type="Registry"; Path=$regKey }
        }

        # AppData check for Teams Meeting add-in folder (Zivver excluded)
        $addinPath = Join-Path "C:\Users\$user" "AppData\Local\Microsoft\TeamsMeetingAddin"
        if (Test-Path $addinPath) {
            $results += [PSCustomObject]@{ User=$user; Type="AppData"; Path=$addinPath }
        }
    }

    return $results
}

function Find-ZivverItem {
    $zivverPaths = @(
        # AppData locations (left as-is)
        [System.IO.Path]::Combine($env:APPDATA, 'Zivver'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver'),
        [System.IO.Path]::Combine($env:ProgramData, 'Zivver'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver\Cache'),
        [System.IO.Path]::Combine($env:APPDATA, 'Zivver\Cache'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver\Logs'),
        [System.IO.Path]::Combine($env:APPDATA, 'Zivver\Config'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver\B.V.'),

        # New Program Files (x86) location for Zivver B.V.
        [System.IO.Path]::Combine('C:\Program Files (x86)', 'Zivver B.V')
    )

    foreach ($path in $zivverPaths) {
        Write-Host "Controleren Zivver pad: $path"
        
        # Extra debugging: verify path for access errors
        if (Test-Path $path) {
            Write-Host "Zivver item gevonden: $path"
            $item = [AppItem]::new()
            $item.DisplayName = switch -Regex ($path) {
                'Cache'  { "Zivver Cache" }
                'Logs'   { "Zivver Logs" }
                'Config' { "Zivver Config" }
                'B\.V'   { "Zivver B.V." }
                default  { "Zivver App" }
            }
            $item.Version = "N/A"
            $item.Source = $path
            $item.UninstallString = "Remove manually"
            $item.Type = "Zivver"
            $item.IsSelected = $false
            $appList.Add($item)
        } else {
            Write-Host "Zivver item NIET gevonden op pad: $path"
            # Additional error check: report when the path is not found
            Write-Host "Fout bij toegang tot pad: $path"
            try {
                $test = Test-Path -Path $path
                if ($test -eq $false) {
                    Write-Host "Pad niet toegankelijk: $path"
                }
            } catch {
                Write-Host "Er is een fout opgetreden bij het testen van het pad: $path"
            }
        }
    }
}

function Clear-TeamsCache {
    Write-Host "Opruimen van Teams cache en logs..."

    $users = Get-ChildItem "C:\Users" -Directory
    foreach ($user in $users) {
        $teamsCachePaths = @(
            "AppData\Local\Microsoft\Teams",
            "AppData\Roaming\Microsoft\Teams"
        )
        foreach ($path in $teamsCachePaths) {
            $fullPath = Join-Path $user.FullName $path
            if (Test-Path $fullPath) {
                Write-Host "Verwijderen van Teams cache voor gebruiker $($user.Name): $fullPath"
                Remove-Item -Path $fullPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    Write-Host "Teams cache opruimen voltooid."
}

function Stop-TeamsOutlook {
    Write-WPFStatus "Sluit Outlook en Teams af..." 15
    
    $processesToKill = @('outlook', 'teams')
    
    foreach ($proc in $processesToKill) {
        # Force-close processes
        Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Remove-FilesInFolder($folder, $processName = $null) {
    if (Test-Path $folder) {

        # --- Disable PowerShell confirmation for the whole function scope ---
        $oldConfirm = $ConfirmPreference
        $ConfirmPreference = 'None'

        try {
            # Stop the process when a process name is provided
            if ($processName) {
                Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object {
                    try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
                }
            }

            # Remove files and subfolders
            Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try { Remove-Item $_.FullName -Force -Recurse -Confirm:$false -ErrorAction Stop } catch {}
            }

            # Remove the root folder itself
            try { Remove-Item -Path $folder -Recurse -Force -Confirm:$false -ErrorAction Stop } catch {}
        } finally {
            # --- Restore confirmation preference to the original value ---
            $ConfirmPreference = $oldConfirm
        }
    }
}

function Remove-M365Item($item) {
    try {
        if ($item.Type -eq "M365") {

            # --- Special case: Teams Machine-Wide installer ---
            if ($item.DisplayName -match "Teams Machine-wide Installer") {
                $msi = Get-WmiObject Win32_Product | Where-Object { $_.Name -eq "Teams Machine-wide Installer" }
                if ($msi) {
                    Write-Host "Stil verwijderen van Teams Machine-Wide Installer..."
                    Start-Process msiexec.exe -ArgumentList "/x `"$($msi.IdentifyingNumber)`" /qn /norestart" -Wait
                    Write-Host "Teams Machine-Wide Installer verwijderd."
                } else {
                    Write-Host "MSI voor Teams Machine-Wide Installer niet gevonden, overslaan."
                }
                return
            }

            # --- Remove Microsoft 365 Apps ---
            if ($item.DisplayName -match "Microsoft 365 Apps") {
                Write-Host "Verwijderen: $($item.DisplayName)"
                $clickToRunExe = "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
                if (Test-Path $clickToRunExe) {
                    $arguments = "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=$($item.ProductId) /quiet /norestart"
                    Start-Process -FilePath $clickToRunExe -ArgumentList $arguments -Wait
                    Write-Host "$($item.DisplayName) succesvol verwijderd."
                } else {
                    Write-Host "OfficeClickToRun.exe niet gevonden, kan $($item.DisplayName) niet verwijderen."
                }
                return
            }

            # --- Remove Teams Meeting add-in ---
            if ($item.DisplayName -match "Teams Meeting Add-in") {
                try {
                    $uninstallString = $item.UninstallString
                    if ($uninstallString -match "MsiExec.exe /I\{(.+?)\}") {
                        $msiGuid = $matches[1]
                        Write-Host "Verwijderen van Teams Meeting Add-in via MSI: $msiGuid"
                        Start-Process msiexec.exe -ArgumentList "/x $msiGuid /qn /norestart" -Wait
                        Write-Host "Teams Meeting Add-in for Microsoft Office verwijderd via MSI."
                    } else {
                        Write-Host "Geen geldige MSI GUID gevonden voor Teams Meeting Add-in."
                    }
                } catch {
                    Write-Host "Fout bij verwijderen van Teams Meeting Add-in via MSI: $($_.Exception.Message)"
                }

                # --- Remove all Teams Meeting add-in folders for all users ---
                $users = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }
                foreach ($user in $users) {
                    foreach ($folderName in @("TeamsMeetingAdd-in", "TeamsMeetingAddinMsis")) {
                        $teamsAddInPath = Join-Path $user.FullName "AppData\Local\Microsoft\$folderName"
                        if (Test-Path $teamsAddInPath) {
                            Write-Host "Verwijderen van $folderName map voor gebruiker: $($user.Name)..."
                            Remove-Item -Path $teamsAddInPath -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Host "$folderName map verwijderd voor gebruiker: $($user.Name)"
                        }
                    }
                }

                # --- Remove registry settings ---
                try {
                    $regKeyPath = "HKCU:\Software\Microsoft\Office\Teams Meeting Add-in"
                    if (Test-Path $regKeyPath) {
                        Remove-Item -Path $regKeyPath -Recurse -Force
                        Write-Host "Teams Meeting Add-in registry-instellingen verwijderd."
                    }
                } catch {
                    Write-Host "Fout bij verwijderen van Teams Meeting Add-in registry: $($_.Exception.Message)"
                }

                # --- Disable Teams Meeting add-in in Outlook ---
                try {
                    $outlook = New-Object -ComObject Outlook.Application
                    $addin = $outlook.COMAddIns | Where-Object { $_.Description -match "Teams Meeting Add-in" }
                    if ($addin) { $addin.Connect = $false; Write-Host "Teams Meeting Add-in uitgeschakeld in Outlook." }
                } catch { Write-Host "Outlook niet beschikbaar of fout bij uitschakelen add-in: $($_.Exception.Message)" }
            }

        }
    } catch {
        Write-Host "Fout bij verwijderen van $($item.DisplayName): $($_.Exception.Message)"
    }

    # --- Scan again after removal ---
    Find-M365Item

    # --- Clean Teams cache after removal ---
    $users = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }

    foreach ($user in $users) {
        $teamsPackages = Get-ChildItem (Join-Path $user.FullName "AppData\Local\Packages") -Directory -Filter "*MSTeams*" -ErrorAction SilentlyContinue
        foreach ($package in $teamsPackages) {
            $cacheRoot = Join-Path $package.FullName "LocalCache\Microsoft\MSTeams\EBWebView"
            if (Test-Path $cacheRoot) {
                Get-ChildItem $cacheRoot -Directory -Filter "WV2Profile_*" -ErrorAction SilentlyContinue | ForEach-Object {
                    $profileFolder = $_.FullName

                    # --- Safety check: Teams or Outlook still running? ---
                    $userProcesses = Get-Process -IncludeUserName -ErrorAction SilentlyContinue | Where-Object { $_.UserName -like "*\$($user.Name)" }
                    $teamsRunning = $userProcesses | Where-Object { $_.Name -match "Teams|OUTLOOK" }

                    if ($teamsRunning) {
                        Write-Host "Teams of Outlook draait nog voor gebruiker $($user.Name), cache verwijderen wordt overgeslagen."
                        return
                    }

                    # --- Set all attributes to Normal so deletion succeeds ---
                    Get-ChildItem $profileFolder -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                        try { $_.Attributes = 'Normal' } catch {}
                    }

                    # --- Remove the entire profile folder ---
                    try {
                        Remove-Item $profileFolder -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Host "Cache verwijderd: $profileFolder"
                    } catch {
                        Write-Host "Kan map niet verwijderen: $profileFolder - $($_.Exception.Message)"
                    }

                    # --- Check and remove empty parent folders ---
                    $parent = Split-Path $profileFolder -Parent
                    while (Test-Path $parent -and (Get-ChildItem $parent -Force | Measure-Object).Count -eq 0) {
                        try {
                            Remove-Item $parent -Force -ErrorAction SilentlyContinue
                            $parent = Split-Path $parent -Parent
                        } catch { break }
                    }
                }
            }
        }
    }
}

function Remove-AppItems {
    param (
        [Parameter(Mandatory=$true)] [System.Collections.ArrayList]$appList,
        [Parameter(Mandatory=$false)] [bool]$RemoveAll = $false
    )

    # --- Determine which items should be removed ---
    $itemsToRemove = if ($RemoveAll) {
        $appList | Where-Object { $_.Type -match "M365|TeamsCache|TeamsAddin|Zivver|LanguagePack" }
    } else {
        $appList | Where-Object { $_.IsSelected }
    }

    if ($itemsToRemove.Count -eq 0) {
        Write-WPFStatus "No items selected for removal."
        return
    }

    Write-WPFStatus "Start verwijderen van $($itemsToRemove.Count) items..."

    # --- Stop relevant processes ---
    $processesToStop = @("Teams","MsTeams","OUTLOOK","WINWORD","EXCEL","POWERPNT","WebViewHost")
    foreach ($proc in $processesToStop | Select-Object -Unique) {
        Get-Process -Name $proc -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
    Start-Sleep -Seconds 2

    $itemsRemovedSuccessfully = 0
    $itemsToKeep = [System.Collections.ArrayList]::new()

    foreach ($item in $itemsToRemove) {
        $success = $false
        Write-WPFStatus "Verwerken: $($item.DisplayName)..."

        try {
            switch ($item.Type) {
                "M365" {
                    Remove-M365Item $item
                    $success = $true
                }

                "TeamsCache" {
                    if ($item.Source -and (Test-Path $item.Source)) {
                        Get-ChildItem -Path $item.Source -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                            try { $_.Attributes = 'Normal' } catch {}
                        }

                        $retry = 0
                        do {
                            try { Remove-Item -Path $item.Source -Recurse -Force -ErrorAction SilentlyContinue } catch {}
                            Start-Sleep -Milliseconds 500
                            $retry++
                        } while ((Test-Path $item.Source) -and $retry -lt 10)

                        $success = !(Test-Path $item.Source)
                        if ($success) { Write-WPFStatus "Map succesvol verwijderd: $($item.Source)" } 
                        else { Write-WPFStatus "Kon map niet verwijderen (mogelijk open bestanden): $($item.Source)" }
                    } else {
                        $success = $true
                    }
                }

                "TeamsAddin" {
                    if ($item.Source -and (Test-Path $item.Source)) {
                        Get-ChildItem -Path $item.Source -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                            try { $_.Attributes = 'Normal' } catch {}
                        }

                        $retry = 0
                        do {
                            try { Remove-Item -Path $item.Source -Recurse -Force -ErrorAction SilentlyContinue } catch {}
                            Start-Sleep -Milliseconds 500
                            $retry++
                        } while ((Test-Path $item.Source) -and $retry -lt 10)

                        $success = !(Test-Path $item.Source)
                        if ($success) { Write-WPFStatus "Map succesvol verwijderd: $($item.Source)" } 
                        else { Write-WPFStatus "Kon map niet verwijderen (mogelijk open bestanden): $($item.Source)" }
                    } else {
                        $success = $true
                    }

                    # Keep consistent display name for Teams add-in items
                    $item.DisplayName = "Teams Meeting Add-in"
                }

                "Zivver" {
                    if ($item.Source -and (Test-Path $item.Source)) {
                        try {
                            Start-Process $item.Source -ArgumentList "/quiet /norestart" -Wait -NoNewWindow -PassThru | Out-Null
                            $success = $true
                        } catch {
                            Write-WPFStatus "Fout bij verwijderen Zivver: $($_.Exception.Message)"
                        }
                    } else {
                        $success = $true
                    }
                }

                "LanguagePack" {
                    $success = Remove-OfficeLanguagePack $item
                }

                default {
                    Write-WPFStatus "Geen verwijderlogica voor $($item.DisplayName)"
                }
            }
        } catch {
            Write-WPFStatus "Fout bij verwijderen $($item.DisplayName): $($_.Exception.Message)"
            $success = $false
        }

        if ($success) { $itemsRemovedSuccessfully++ } else { $itemsToKeep.Add($item) | Out-Null }
    }

    # --- Update appList and GUI ---
    $appList.Clear()
    $appList.AddRange($itemsToKeep)
    $window.Dispatcher.Invoke([action]{ Update-DataGridWPF -DataList $appList })

    # --- Final status ---
    if ($itemsRemovedSuccessfully -gt 0) {
        Write-WPFStatus "Removal completed. $itemsRemovedSuccessfully items were removed successfully."
    } elseif ($itemsToKeep.Count -gt 0) {
        Write-WPFStatus "Removal failed for some items. Check the log for details."
    } else {
        Write-WPFStatus "Geen items verwijderd."
    }
}

function Clear-Folder($folder) {
    if (-not (Test-Path $folder)) { return }

    # Stop processes that may hold the folder
    $processes = @("Teams", "msedge", "WebViewHost")
    foreach ($p in $processes) {
        Get-Process -Name $p -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    Start-Sleep -Seconds 2 # even wachten tot processen echt gestopt zijn

    # Remove all content from the folder
    Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try { Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue } catch {}
    }

    # Remove the folder itself
    try { Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue } catch {}

    # Check whether the folder still exists and report this
    if (Test-Path $folder) {
        Write-Host "Map $folder bestaat nog. Waarschijnlijk door Teams/WebView opnieuw aangemaakt."
    } else {
        Write-Host "Map $folder succesvol verwijderd."
    }
}

function Invoke-TeamsAddinReinstall {
    Write-WPFStatus "Start herinstallatie Teams Add-in...", 0

    # 1. Stop all relevant M365 processes (Teams, Outlook, etc.)
    $m365Processes = @("Teams", "OUTLOOK", "Word", "Excel", "PowerPoint", "OneNote", "Skype")
    foreach ($process in $m365Processes) {
        Write-WPFStatus "Stoppen van proces: $process...", 5
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    # 2. Remove Teams add-in and cache
    Write-WPFStatus "Teams Add-in en cache verwijderen...", 15
    $users = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notmatch "Public|Default|Default User|All Users" }
    
    foreach ($user in $users) {
        $regPath = "Registry::HKEY_USERS\$((Get-LocalUser -Name $user.Name -ErrorAction SilentlyContinue).SID)\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
        if (Test-Path $regPath) {
            try { Remove-Item -Path $regPath -Recurse -Force } catch { Write-WPFStatus "Fout bij verwijderen van registry voor $($user.Name)", 25 }
        }

        $cachePath = Join-Path $user.FullName "AppData\Local\Microsoft\Teams"
        if (Test-Path $cachePath) {
            try { Remove-Item -Path $cachePath -Recurse -Force } catch { Write-WPFStatus "Fout bij verwijderen van cache voor $($user.Name)", 25 }
        }
    }

    Write-WPFStatus "Teams Add-in en cache verwijderd.", 30

    # 3. Reinstall the Teams add-in via ClickToRun
    Write-WPFStatus "Teams Add-in opnieuw installeren...", 40
    $clickToRunPaths = @(
        "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe",
        "$env:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    )
    $clickToRunExe = $clickToRunPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($clickToRunExe) {
        $productId = "TeamsAddinForOutlook"
        $arguments = "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=$productId /quiet /norestart"
        Start-Process -FilePath $clickToRunExe -ArgumentList $arguments -Wait
        Write-WPFStatus "Teams Add-in installatie voltooid.", 80
    } else { 
        Write-WPFStatus "OfficeClickToRun.exe niet gevonden.", 40 
    }

    # 4. Register the add-in for all users
    Write-WPFStatus "Registeren van Teams Add-in voor alle gebruikers...", 60
    $users = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notmatch "Public|Default|Default User|All Users" }
    foreach ($user in $users) {
        try {
            $sid = (Get-LocalUser -Name $user.Name -ErrorAction SilentlyContinue).SID
            if ($sid) {
                $regPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
                if (-not (Test-Path $regPath)) { 
                    New-Item -Path $regPath -Force | Out-Null 
                }
                Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3
                Set-ItemProperty -Path $regPath -Name "Description" -Value "Teams Meeting Add-in"
                Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Teams Meeting Add-in"
            }
        } catch { 
            Write-WPFStatus "Kon registry voor $($user.Name) niet instellen: $($_.Exception.Message)", 70 
        }
    }

    Write-WPFStatus "Teams Add-in herinstallatie voltooid.", 90

    # Add success popup
    [System.Windows.Forms.MessageBox]::Show(
        "Teams Add-in Installatie Voltooid.", 
        "Teams Add-in Installatie Voltooid", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

function Invoke-ZivverReinstall {
    Write-WPFStatus "Start herinstallatie Zivver Plugin...", 0

    # URL to download the Zivver installer
    $zivverDownloadUrl = "https://downloads.zivver.com/officeplugin/latest/Zivver.OfficePlugin.Installer.msi"
    $tempDir = Join-Path $env:TEMP "ZivverInstaller"
    $zivverInstallerPath = Join-Path $tempDir "Zivver.OfficePlugin.Installer.msi"
    
    # Build the AppData\Local\Zivver B.V. path dynamically
    $appDataInstallDir = Join-Path $env:LOCALAPPDATA "Zivver B.V."

    # Stop all M365 processes (Teams, Outlook, etc.)
    $m365Processes = @("Teams", "Outlook", "Word", "Excel", "PowerPoint", "OneNote", "Skype")
    foreach ($process in $m365Processes) {
        Write-WPFStatus "Stoppen van proces: $process...", 5
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    # -----------------------------------------------------------
    # 2. Find and silently remove current Zivver installation
    # -----------------------------------------------------------
    Write-WPFStatus "Start detectie en stille verwijdering van Zivver B.V. installatie...", 15
    $zivverEntry = $null
    $uninstallPaths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($path in $uninstallPaths) {
        $zivverEntry = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | Where-Object { 
            $_.DisplayName -like "*Zivver*" -and $_.Publisher -like "*Zivver B.V.*"
        }
        if ($zivverEntry) { break }
    }

    if ($zivverEntry -and $zivverEntry.UninstallString) {
        $uninstallCommand = $zivverEntry.UninstallString
        
        if ($uninstallCommand -match 'msiexec\.exe') {
            $finalUninstallCommand = $uninstallCommand.Replace("/I", "/X") + " /quiet /norestart"
        } elseif ($uninstallCommand -match '\.exe') {
            $finalUninstallCommand = "$uninstallCommand /S /uninstall /quiet /norestart"
        } else {
            $finalUninstallCommand = "$uninstallCommand /quiet"
        }

        Write-WPFStatus "Zivver installatie gevonden. Start stille verwijdering.", 20
        Start-Process -FilePath cmd -ArgumentList "/c $finalUninstallCommand" -Wait -WindowStyle Hidden -ErrorAction Stop

        Write-WPFStatus "Removal process completed. Waiting for system...", 25
        Start-Sleep -Seconds 5
    } else {
        Write-WPFStatus "Zivver installatie niet gevonden. Overslaan van de-installatie.", 25
    }

    # -----------------------------------------------------------
    # 3. Download Zivver MSI installer
    # -----------------------------------------------------------
    Write-WPFStatus "Start downloaden van Zivver MSI Installer...", 40

    try {
        # Check if the temp folder exists and create it if needed
        if (-not (Test-Path $tempDir)) { 
            New-Item -Path $tempDir -ItemType Directory -Force | Out-Null 
        }

        # Download the MSI
        Invoke-WebRequest -Uri $zivverDownloadUrl -OutFile $zivverInstallerPath -ErrorAction Stop

        Write-WPFStatus "Download voltooid: Zivver MSI Installer.", 50
    } catch {
        $errorMessage = $_.Exception.Message
        Write-WPFStatus "FOUT bij downloaden: $errorMessage", 50

        # Show an error message to the user
        [System.Windows.Forms.MessageBox]::Show(
            "Er is een fout opgetreden bij het downloaden van de Zivver MSI Installer: $errorMessage", 
            "Fout bij Downloaden", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # -----------------------------------------------------------
    # 4. Start installation under AppData\Local
    # -----------------------------------------------------------
    Write-WPFStatus "Start directe stille installatie naar AppData\Local...", 75

    # Ensure the folder exists
    if (-not (Test-Path $appDataInstallDir)) {
        Write-Host "AppData installatiedirectory bestaat niet, maken..."
        New-Item -Path $appDataInstallDir -ItemType Directory -Force | Out-Null
    }

    # Specify the install path in AppData\Local
    $arguments = @(
        "/i", "`"$zivverInstallerPath`"",
        "MSIINSTALLPERUSER=1",          # Installeer per gebruiker
        "I_ACCEPT_LICENSE_AGREEMENT=1", # Licentie akkoord
        "USEIMPERSONATE=0",             # Geen impersonation gebruiken
        "/quiet",                       # Stille installatie
        "/norestart",                   # Geen herstart na installatie
        "INSTALLDIR=`"$appDataInstallDir`""  # Specificeer het installatiedirectory in AppData\Local\Zivver B.V.
    )

    # Start the installation using msiexec
    Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -WindowStyle Hidden -ErrorAction Stop

    Write-WPFStatus "Installatie voltooid.", 90

    # Add success popup
    [System.Windows.Forms.MessageBox]::Show(
        "Zivver Installatie Voltooid.", 
        "Zivver Installatie Voltooid", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

function Update-DataGridWPF {
    param(
        [System.Collections.ObjectModel.ObservableCollection[AppItem]]$DataList
    )
    
    # Use the WPF DataGrid variable ($dgApps) and the Dispatcher
    $dgApps.Dispatcher.Invoke([action]{
        $dgApps.ItemsSource = $null
        $dgApps.ItemsSource = $DataList
        $dgApps.Items.Refresh()
    })
}

##############################
# PART 3/5: XAML DEFINITION
##############################
      
# XAML GUI (DEFINITIEVE LAYOUT)
$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="deRolfgroep Package Manager" Height="750" Width="1000" WindowStartupLocation="CenterScreen"
        Background="#1E1E1E" WindowStyle="SingleBorderWindow">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="M365 Installations, Language Packs &amp; Components"
                   FontSize="24" FontWeight="Bold" Foreground="White" Margin="0,0,0,20"/>

        <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,0,0,10">
            <Button Name="btnScanM365" Content="🔍 Scan M365/Teams" Width="180" Margin="5" Background="#007AFF" Foreground="White"/>
            <Button Name="btnScanZivver" Content="🔍 Scan Zivver" Width="150" Margin="5" Background="#FF9500" Foreground="White"/>
            <Button Name="btnRemove" Content="❌ Remove Selected" Width="200" Margin="5" Background="#FF3B30" Foreground="White" IsEnabled="False"/>
        </StackPanel>

        <DataGrid Name="dgApps" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False"
                  SelectionMode="Extended" FontSize="14" Background="#2D2D2D" HeadersVisibility="Column" 
                  ColumnHeaderHeight="30" SelectionUnit="FullRow">
            <DataGrid.Resources>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#444444"/>
                    <Setter Property="Foreground" Value="White"/>
                    <Setter Property="FontWeight" Value="Bold"/>
                    <Setter Property="HorizontalContentAlignment" Value="Center"/>
                </Style>
                <Style TargetType="DataGridRow">
                    <Setter Property="Background" Value="#2D2D2D"/>
                    <Style.Triggers>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="#007AFF"/>
                            <Setter Property="Foreground" Value="White"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
                <Style x:Key="TextBlockWhite" TargetType="TextBlock">
                    <Setter Property="Foreground" Value="White"/>
                </Style>
            </DataGrid.Resources>
            <DataGrid.Columns>
                <DataGridTemplateColumn Width="150">
                    <DataGridTemplateColumn.HeaderTemplate>
                        <DataTemplate>
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                <CheckBox x:Name="chkSelectAll" VerticalAlignment="Center" Margin="0,0,5,0"/>
                                <TextBlock Text="Select" Foreground="White" VerticalAlignment="Center"/>
                            </StackPanel>
                        </DataTemplate>
                    </DataGridTemplateColumn.HeaderTemplate>
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="*" ElementStyle="{StaticResource TextBlockWhite}"/>
                <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="120" ElementStyle="{StaticResource TextBlockWhite}"/>
                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="120" ElementStyle="{StaticResource TextBlockWhite}"/>
                <DataGridTextColumn Header="UninstallString" Binding="{Binding UninstallString}" Width="250" ElementStyle="{StaticResource TextBlockWhite}"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Voeg de ProgressBar toe voor de voortgangsweergave -->
        <ProgressBar Name="progressBar" Grid.Row="3" Height="20" Width="300" Minimum="0" Maximum="100" Visibility="Hidden" Margin="0,10,0,0"/>

        <!-- Voeg een Label toe voor statusupdates -->
        <TextBlock Name="lblStatus" Grid.Row="4" Foreground="White" FontSize="14" Margin="0,10,0,0"/>

        <StackPanel Orientation="Vertical" Grid.Row="5">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,5,0,5">
                <Button Name="btnRemoveAll" Content="🧹 Remove M365 Completely" Width="250" Margin="5,0" Background="#FF3B30" Foreground="White"/>
            </StackPanel>

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="btnReinstall" Content="🔄 Reinstall M365" Width="200" Margin="5" Background="#34C759" Foreground="White"/>
                <Button Name="btnReinstallZivver" Content="🔄 Reinstall Zivver Plugin" Width="200" Margin="5" Background="#34C759" Foreground="White"/>
                <Button Name="btnReinstallTeamsAddin" Content="🔄 Reinstall Teams Add-in" Width="200" Margin="5" Background="#34C759" Foreground="White"/>
            </StackPanel>
            
        </StackPanel>

        <TextBlock Text="v1.0" Foreground="White" FontSize="12" HorizontalAlignment="Right" VerticalAlignment="Bottom" Grid.Row="5" Margin="0,0,10,5"/>
    </Grid>
</Window>
"@

##############################
# PART 4/5: XAML LOAD, WIRING & EVENT HANDLERS
##############################

# --- 1. Load XAML and resolve elements (CRITICAL) ---
$reader = New-Object System.Xml.XmlNodeReader ([xml]$XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# --- 2. Resolve GUI elements via FindName ---
$dgApps = $window.FindName("dgApps")
$btnRemove = $window.FindName("btnRemove")
$btnReinstall = $window.FindName("btnReinstall")
$btnRemoveAll = $window.FindName("btnRemoveAll")
$btnScanM365 = $window.FindName("btnScanM365")
$btnScanZivver = $window.FindName("btnScanZivver")
$btnReinstallZivver = $window.FindName("btnReinstallZivver")
$btnReinstallTeamsAddin = $window.FindName("btnReinstallTeamsAddin")
$lblStatus = $window.FindName("lblStatus") # <== Status TextBlock
$progressBar = $window.FindName("progressBar") # <== Progress Bar toegevoegd aan XAML
$chkSelectAll = $window.FindName("chkSelectAll") # <== Header CheckBox

# --- 3. Data binding (prevents ArgumentNullException at ShowDialog) ---
$dgApps.ItemsSource = $appList

# ==========================================
# WIRE EVENT HANDLERS
# ==========================================

# --- Header checkbox functionality ---
$window.Dispatcher.Invoke([action]{
    # Zoek de header checkbox in de DataGrid
    $dgApps.Dispatcher.InvokeAsync({
        $headerCheckBox = Find-VisualChild -parent $dgApps -childType ([type]"System.Windows.Controls.CheckBox")
        if ($headerCheckBox) {
            # Select all items
            $headerCheckBox.Add_Checked({
                # Update the selection state for all items in the list
                foreach ($item in $appList) {
                    $item.IsSelected = $true
                }
                # Reload the DataGrid to reflect selection changes
                $dgApps.ItemsSource = $null  # Reset de ItemsSource
                $dgApps.ItemsSource = $appList
            })

            # Deselect all items
            $headerCheckBox.Add_Unchecked({
                # Update the selection state for all items in the list
                foreach ($item in $appList) {
                    $item.IsSelected = $false
                }
                # Reload the DataGrid to reflect selection changes
                $dgApps.ItemsSource = $null  # Reset de ItemsSource
                $dgApps.ItemsSource = $appList
            })
        }
    }, [System.Windows.Threading.DispatcherPriority]::Loaded)
})

# --- Make header checkbox functional ---
if ($chkSelectAll) {
    $chkSelectAll.Add_Checked({
        # Select all items in the list
        foreach ($item in $appList) { 
            $item.IsSelected = $true 
        }
        # Reload the DataGrid to reflect selection changes
        $dgApps.ItemsSource = $null  # Reset de ItemsSource
        $dgApps.ItemsSource = $appList
    })

    $chkSelectAll.Add_Unchecked({
        # Deselect all items in the list
        foreach ($item in $appList) { 
            $item.IsSelected = $false 
        }
        # Reload the DataGrid to reflect selection changes
        $dgApps.ItemsSource = $null  # Reset de ItemsSource
        $dgApps.ItemsSource = $appList
    })
}

# --- Find header checkbox after rendering ---
$window.Add_Loaded({
    $headerCheckBox = Find-VisualChild -parent $dgApps -childType ([Type]"System.Windows.Controls.CheckBox")
    if ($headerCheckBox) {
        $headerCheckBox.Add_Checked({
            # Select all items in the list
            foreach ($item in $appList) {
                $item.IsSelected = $true
            }
            # Reload the DataGrid to reflect selection changes
            $dgApps.ItemsSource = $null  # Reset de ItemsSource
            $dgApps.ItemsSource = $appList
        })

        $headerCheckBox.Add_Unchecked({
            # Deselect all items in the list
            foreach ($item in $appList) {
                $item.IsSelected = $false
            }
            # Reload the DataGrid to reflect selection changes
            $dgApps.ItemsSource = $null  # Reset de ItemsSource
            $dgApps.ItemsSource = $appList
        })
    }
})

# --- Init paths ---
$localAppDataRoot = [string]$env:LocalAppData

# --- Button: Remove All M365/Teams items ---
$btnRemoveAll.Add_Click({
    Start-AsyncJob -JobName "Full M365 Removal" -ScriptBlock {
        param($appList)

        Update-Status "Removing all M365/Teams components..."

        # Stop Teams and Office processes
        $procs = "Teams","MsTeams","OUTLOOK","WINWORD","EXCEL","POWERPNT"
        foreach ($p in $procs | Select-Object -Unique) {
            Get-Process -Name $p -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force }
        }
        Start-Sleep -Seconds 1

        # Teams and add-in paths
        $teamsPaths = @(
            # Teams 1.x
            Join-Path $localAppDataRoot "Microsoft\Teams\Application Cache\Cache",
            Join-Path $localAppDataRoot "Microsoft\Teams\Blob_storage",
            Join-Path $localAppDataRoot "Microsoft\Teams\Cache",
            Join-Path $localAppDataRoot "Microsoft\Teams\databases",
            Join-Path $localAppDataRoot "Microsoft\Teams\GPUCache",
            Join-Path $localAppDataRoot "Microsoft\Teams\IndexedDB",
            Join-Path $localAppDataRoot "Microsoft\Teams\Local Storage",
            Join-Path $localAppDataRoot "Microsoft\Teams\tmp",
            # Teams 2.x
            Join-Path $localAppDataRoot "Microsoft\Teams\Logs",
            # Teams Add-in
            Join-Path $localAppDataRoot "Microsoft\TeamsMeetingAddin"
        )

        $itemsRemovedSuccessfully = 0
        $itemsToKeep = [System.Collections.ArrayList]::new()

        # Remove items from the list
        $itemsToRemove = $appList | Where-Object { $_.Type -match "M365|TeamsCache|TeamsAddin|Zivver" }
        foreach ($item in $itemsToRemove) {
            $success = $false
            Write-WPFStatus "Verwerking: $($item.DisplayName)..."
            try {
                switch ($item.Type) {
                    'M365' { Remove-M365Item $item; $success = $true }
                    'TeamsCache' {
                        $success = Remove-PathSafe $item.Source $item.DisplayName
                    }
                    'TeamsAddin' {
                        $success = Remove-PathSafe $item.Source "Teams Meeting Add-in"
                        $item.DisplayName = "Teams Meeting Add-in"
                    }
                    'Zivver' {
                        Start-Process $item.Source -ArgumentList "/quiet /norestart" -Wait -NoNewWindow -PassThru | Out-Null
                        $success = $true
                    }
                    default { Write-WPFStatus "Geen verwijderlogica voor $($item.DisplayName)" }
                }
            } catch {
                Write-WPFStatus "FOUT: Kon $($item.DisplayName) niet verwijderen. ($($_.Exception.Message))"
            }
            if ($success) { 
                $itemsRemovedSuccessfully++
                # Remove the item from the list and refresh the DataGrid
                $appList.Remove($item)
                $window.Dispatcher.Invoke([action]{ Update-DataGridWPF -DataList $appList })
            } else { 
                $itemsToKeep.Add($item) | Out-Null 
            }
        }

        # Also remove standard Teams cache folders
        foreach ($path in $teamsPaths) {
            if (Remove-PathSafe $path $path) { $itemsRemovedSuccessfully++ }
        }

        # Update list and DataGrid after processing all items
        $appList.Clear()
        $appList.AddRange($itemsToKeep)
        Write-WPFStatus "Complete verwijdering voltooid. $itemsRemovedSuccessfully items succesvol verwijderd."

        # Ensure all buttons are reactivated
        $window.Dispatcher.Invoke([action]{ Enable-AllButtons })
        $window.Dispatcher.Invoke([action]{ Update-DataGridWPF -DataList $appList })
    } -ArgumentList $appList
})

# --- Button: Remove selected items ---
$btnRemove.Add_Click({
    # Get selected items from the list
    $selectedItems = $appList | Where-Object { $_.IsSelected }

    if ($selectedItems.Count -eq 0) {
        Write-WPFStatus "No items selected for removal."
        return
    }

    # Start removing selected items
    Write-WPFStatus "Removing selected items..."

    foreach ($item in $selectedItems) {
        try {
            Write-WPFStatus "Removing item: $($item.DisplayName)"

            # Stop all related Teams processes
            if ($item.DisplayName -match "Teams") {
                Write-WPFStatus "Stoppen van Teams-gerelateerde processen..."

                # Find processes named ms-teams and stop them
                $teamsProcesses = Get-Process | Where-Object { $_.Name -eq "ms-teams" }

                foreach ($process in $teamsProcesses) {
                    Write-WPFStatus "Stoppen van proces: $($process.Name) met PID: $($process.Id)"
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                }

                # Stop Teams Updater if it is running
                $teamsUpdater = Get-Process | Where-Object { $_.Name -eq "TeamsUpdater" }
                if ($teamsUpdater) {
                    Write-WPFStatus "Stoppen van Teams Updater"
                    Stop-Process -Name "TeamsUpdater" -Force -ErrorAction SilentlyContinue
                }

                # Force-stop Teams via Taskkill
                Write-WPFStatus "Forceer afsluiten van Teams via Taskkill"
                try {
                    Start-Process "taskkill" -ArgumentList "/F /IM ms-teams.exe" -NoNewWindow -ErrorAction SilentlyContinue
                } catch {
                    Write-WPFStatus "Kan ms-teams niet dwingen af te sluiten via Taskkill: $($_.Exception.Message)"
                }
            }

            # Remove files in the selected item folder
            Write-WPFStatus "Removing files in folder: $($item.Source)"
            if (Test-Path $item.Source) {
                try {
                    # Remove file or folder
                    Remove-Item -Path $item.Source -Recurse -Force -ErrorAction Stop
                    Write-WPFStatus "Bestanden succesvol verwijderd: $($item.Source)"
                    
                    # Check whether the parent folder is empty and remove it if possible
                    $parentDir = Split-Path $item.Source
                    if (Test-Path $parentDir -and (Get-ChildItem $parentDir).Count -eq 0) {
                        Remove-Item -Path $parentDir -Force -ErrorAction SilentlyContinue
                        Write-WPFStatus "Lege map verwijderd: $parentDir"
                    }
                } catch {
                    Write-WPFStatus "Fout bij verwijderen van bestanden in $($item.Source): $($_.Exception.Message)"
                }
            }
            else {
                Write-WPFStatus "Fout: Pad bestaat niet voor $($item.DisplayName)"
            }

            # After successful removal, remove the item from the list
            $appList.Remove($item)
            Write-WPFStatus "Item succesvol verwijderd: $($item.DisplayName)"
        }
        catch {
            Write-WPFStatus "Fout bij verwijderen van $($item.DisplayName): $($_.Exception.Message)"
        }
    }

    # Update UI and show completion message after removal is done
    Write-WPFStatus "Removal completed. All selected items were removed successfully."
    Update-DataGridWPF -DataList $appList
})

# --- Button: Scan M365/Teams including Teams Meeting Add-in ---
$btnScanM365.Add_Click({
    $appList.Clear()
    Write-WPFStatus "Starting scan: M365/Teams components..."
    Find-M365Item

    # --- 1. Per-machine registry scan ---
    $regPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($path in $regPaths) {
        Get-ItemProperty $path -ErrorAction SilentlyContinue 2>$null |
        Where-Object { $_.DisplayName -match 'Office|Microsoft 365|Language Pack|Proofing|Teams|OneDrive|Outlook|Word|Excel|PowerPoint' -and $_.DisplayName -notmatch 'Zivver' } | # Filter Zivver out
        ForEach-Object {
            $item = [AppItem]::new()
            $item.DisplayName      = $_.DisplayName
            $item.Version          = $_.DisplayVersion
            $item.Source           = if ($_.DisplayName -match "Language Pack") {'Language Pack'} else {'Registry'}
            $item.UninstallString  = $_.UninstallString
            $item.Type             = "M365"
            $item.IsSelected       = $false
            $appList.Add($item)
        }
    }

    # --- 2. Per-user scan ---
    $users = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue 2>$null |
             Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }

    foreach ($user in $users) {
        Write-WPFStatus "Scannen gebruiker $($user.Name)"
        $sid = $null
        try { $sid = (New-Object System.Security.Principal.NTAccount($user.Name)).Translate([System.Security.Principal.SecurityIdentifier]).Value 2>$null } catch {}

        # Teams add-in registry
        if ($sid) {
            $regKey = "HKU\$sid\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
            if ((Test-Path $regKey -ErrorAction SilentlyContinue)) {
                $item = [AppItem]::new()
                $item.DisplayName      = "Teams Meeting Add-in (Registry)"
                $item.Version          = "N/A"
                $item.Source           = $regKey
                $item.UninstallString  = "Registry key"
                $item.Type             = "TeamsAddin"
                $item.IsSelected       = $false
                $appList.Add($item)
            }
        }

        # Teams add-in AppData
        $localAppDataRoot = Join-Path $user.FullName "AppData\Local"
        $teamsAddinPaths = @()
        $teamsAddinPaths += Join-Path $localAppDataRoot "Microsoft\TeamsMeetingAddin"
        $teamsAddinPaths += Join-Path $localAppDataRoot "Microsoft\TeamsMeetingAddinMsis"

        foreach ($path in $teamsAddinPaths) {
            if ((Test-Path $path -ErrorAction SilentlyContinue)) {
                $item = [AppItem]::new()
                $item.DisplayName      = "Teams Meeting Add-in"
                $item.Version          = "N/A"
                $item.Source           = $path
                $item.UninstallString  = $path
                $item.Type             = "TeamsAddin"
                $item.IsSelected       = $false
                $appList.Add($item)
            }
        }

        # Legacy Teams cache
        $oldTeamsCachePaths = @()
        $oldTeamsCachePaths += Join-Path $localAppDataRoot "Microsoft\Teams"
        $oldTeamsCachePaths += Join-Path $localAppDataRoot "Microsoft\TeamsDesktopClient"
        $oldTeamsCachePaths += Join-Path $localAppDataRoot "Microsoft\Teams\Cache"
        $oldTeamsCachePaths += Join-Path $localAppDataRoot "Microsoft\Teams\Service Worker\CacheStorage"
        $oldTeamsCachePaths += Join-Path $localAppDataRoot "Microsoft\Teams\IndexedDB"
        $oldTeamsCachePaths += Join-Path $localAppDataRoot "Microsoft\Office\16.0\OfficeFileCache"

        foreach ($fullPath in $oldTeamsCachePaths) {
            if ((Test-Path $fullPath -ErrorAction SilentlyContinue)) {
                $item = [AppItem]::new()
                $item.DisplayName      = "Cache: $(Split-Path $fullPath -Leaf)"
                $item.Version          = "N/A"
                $item.Source           = $fullPath
                $item.UninstallString  = $fullPath
                $item.Type             = "TeamsCache"
                $item.IsSelected       = $false
                $appList.Add($item)
            }
        }

        # Latest Teams UWP cache
        $teamsCacheRoot = Join-Path $localAppDataRoot "Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"
        if ((Test-Path $teamsCacheRoot -PathType Container -ErrorAction SilentlyContinue)) {
            $item = [AppItem]::new()
            $item.DisplayName      = "Teams Cache (Main Folder)"
            $item.Version          = "N/A"
            $item.Source           = $teamsCacheRoot
            $item.UninstallString  = $teamsCacheRoot
            $item.Type             = "TeamsCache"
            $item.IsSelected       = $false
            $appList.Add($item)
        }
    }

    Write-WPFStatus "✅ M365/Teams scan completed - $($appList.Count) items found."

    $window.Dispatcher.Invoke([action]{ $btnRemove.IsEnabled = $true; $btnRemove.Visibility = [System.Windows.Visibility]::Visible })
    Update-DataGridWPF $appList
})

# --- Scan Zivver including Zivver B.V. folders ---
$btnScanZivver.Add_Click({
    $appList.Clear()
    
    # Send scan start message to WPF status
    Write-WPFStatus "Starting scan: Zivver folders..."
    
    # Step 1: Check Zivver registry locations (COM add-in)
    $zivverRegPaths = @(
        "HKCU:\Software\Microsoft\Office\Outlook\Addins\Zivver",  # HKEY_CURRENT_USER voor de gebruiker
        "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\Zivver",  # HKEY_LOCAL_MACHINE voor systeemwijde installaties
        "HKCU:\Software\Microsoft\Office\Teams\Addins\Zivver"    # Check Teams als Zivver ook daar wordt geladen
    )

    foreach ($regPath in $zivverRegPaths) {
        if (Test-Path $regPath) {
            # Do not send debug messages to the console, only to WPF status
            Write-WPFStatus "Zivver add-in gevonden in het register op: $regPath"
            $item = [AppItem]::new()
            $item.DisplayName      = "Zivver Add-in (Registry)"
            $item.Version          = "N/A"
            $item.Source           = $regPath
            $item.UninstallString  = "Remove manually"
            $item.Type             = "Zivver"
            $item.IsSelected       = $false
            $appList.Add($item)
        }
    }

    # Step 2: Check file locations for Zivver
    $zivverPaths = @(
        [System.IO.Path]::Combine($env:APPDATA, 'Zivver'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver'),
        [System.IO.Path]::Combine($env:ProgramData, 'Zivver'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver\Cache'),
        [System.IO.Path]::Combine($env:APPDATA, 'Zivver\Cache'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver\Logs'),
        [System.IO.Path]::Combine($env:APPDATA, 'Zivver\Config'),
        [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Zivver\B.V.'),

        # Include Program Files (x86) location for Zivver B.V.
        [System.IO.Path]::Combine('C:\Program Files (x86)', 'Zivver B.V')
    ) | Sort-Object -Unique

    foreach ($path in $zivverPaths) {
        if (Test-Path $path) {
            # Send only important messages to WPF
            Write-WPFStatus "Zivver item gevonden: $path"
            $item = [AppItem]::new()
            $item.DisplayName      = switch -Regex ($path) {
                'Cache'  { "Zivver Cache" }
                'Logs'   { "Zivver Logs" }
                'Config' { "Zivver Config" }
                'B\.V'   { "Zivver B.V." }
                default  { "Zivver App" }
            }
            $item.Version          = "N/A"
            $item.Source           = $path
            $item.UninstallString  = "Remove manually"
            $item.Type             = "Zivver"
            $item.IsSelected       = $false
            $appList.Add($item)
        }
    }

    # Step 3: Update UI with scan completion status to WPF
    Write-WPFStatus "Zivver scan completed - $($appList.Count) items found."
    $btnRemove.Dispatcher.Invoke([action]{ $btnRemove.IsEnabled = ($appList.Count -gt 0) })
    Update-DataGridWPF $appList
})

# --- Reinstall M365/Office ---
$btnReinstall.Add_Click({
	Write-Host "Herinstallatie van M365 wordt gestart..."

    $odtFolder = "C:\ODT"
    if (-not (Test-Path $odtFolder)) { New-Item -Path $odtFolder -ItemType Directory -Force | Out-Null }
    $setupExe = Join-Path $odtFolder "setup.exe"
    $logFolder = Join-Path $odtFolder "Logs"
    if (-not (Test-Path $logFolder)) { New-Item -Path $logFolder -ItemType Directory -Force | Out-Null }

    # Download ODT if not present
    if (-not (Test-Path $setupExe)) {
		try {
			Write-Host "Download Office Deployment Tool..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $page = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
            $link = $page.Links | Where-Object { $_.href -match "officedeploymenttool.*\.exe" } | Select-Object -First 1
            if (-not $link) { throw "Kan ODT downloadlink niet vinden." }
            $odtExe = Join-Path $odtFolder "OfficeDeploymentTool.exe"
            Invoke-WebRequest -Uri $link.href -OutFile $odtExe -UseBasicParsing
            Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:$odtFolder" -Wait
		} catch {
            Write-Host "Fout bij ODT download: $($_.Exception.Message)"
            return
        }
    }

    $bitness = if ([Environment]::Is64BitOperatingSystem) { "64" } else { "32" }

    # Create installation XML configuration
    $configFile = Join-Path $odtFolder "install365.xml"
    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="$bitness" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us"/>
      <Language ID="nl-nl"/>
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE"/>
  <Logging Level="Standard" Path="$logFolder"/>
</Configuration>
"@
    $xmlContent | Set-Content $configFile -Force

    # Start installation
    $arguments = @("/configure", $configFile)
    Start-Process -FilePath $setupExe -ArgumentList $arguments -Wait
    Write-Host "✅ Herinstallatie M365 voltooid."
    Write-WPFStatus "Microsoft 365 is succesvol herinstalleerd.", 100
})

# --- Full M365 removal ---
$btnRemoveAll.Add_Click({
    Update-Status "🧹 Removing all M365 components..."

    $odtFolder = "C:\ODT"
    $logFolder = Join-Path $odtFolder "Logs"

    if (-not (Test-Path $odtFolder)) { New-Item $odtFolder -ItemType Directory -Force | Out-Null }
    if (-not (Test-Path $logFolder)) { New-Item $logFolder -ItemType Directory -Force | Out-Null }

    $setupExe = Join-Path $odtFolder "setup.exe"

    if (-not (Test-Path $setupExe)) {
        try {
            Update-Status "Download Office Deployment Tool..."
            $page = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
            $link = $page.Links | Where-Object { $_.href -match "officedeploymenttool.*\.exe" } | Select-Object -First 1
            $odtExe = Join-Path $odtFolder "OfficeDeploymentTool.exe"
            Invoke-WebRequest -Uri $link.href -OutFile $odtExe -UseBasicParsing
            Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:$odtFolder" -Wait
        } catch {
            Update-Status "Fout bij ODT download: $($_.Exception.Message)"
            return
        }
    }

    # Create XML for full removal
    $configFile = Join-Path $odtFolder "remove365.xml"
    $configXml = @"
<Configuration>
  <Remove All="TRUE"/>
  <Display Level="Full" AcceptEULA="TRUE"/>
  <Logging Level="Standard" Path="$logFolder"/>
</Configuration>
"@
    $configXml | Set-Content $configFile -Force

    # Start removal
    Start-Process -FilePath $setupExe -ArgumentList "/configure `"$configFile`"" -Wait
    Update-Status "Alle M365 componenten verwijderd."
})


# --- Extra event handlers for reinstall buttons ---
$btnReinstallZivver.Add_Click({ Invoke-ZivverReinstall })
$btnReinstallTeamsAddin.Add_Click({ Invoke-TeamsAddinReinstall })


##############################
# PART 5/5: SHOW WINDOW
##############################

# --- Show Window (Triggers WPF rendering) ---
$window.ShowDialog()
