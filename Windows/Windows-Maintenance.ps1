﻿<#
    Copyright (C) 2022  Stolpe.io
    <https://stolpe.io>
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#>

$PSVersion = $Host.Version.Major

Function Find-NeededModules {
    Write-Output "`n=== Making sure that all modules are installad and up to date ===`n"
    # Modules to check if it's installed and imported
    $NeededModules = @("PowerShellGet", "MSIPatches", "PSWindowsUpdate", "NuGet")
    $NeededPackages = @("NuGet", "PowerShellGet")
    # Collects all of the installed modules on the system
    $CurrentModules = Get-InstalledModule | Select-Object Name, Version | Sort-Object Name
    # Collects all of the installed packages
    $AllPackageProviders = Get-PackageProvider -ListAvailable | Select-Object Name -ExpandProperty Name

    # Making sure that TLS 1.2 is used.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    # Installing needed packages if it's missing.
    Write-Output "Making sure that all of the PackageProviders that are needed are installed..."
    foreach ($Provider in $NeededPackages) {
        if ($Provider -NotIn $AllPackageProviders) {
            Try {
                Write-Output "Installing $($Provider) as it's missing..."
                Install-PackageProvider -Name $provider -Scope AllUsers -Force
                Write-Output "$($Provider) is now installed" -ForegroundColor Green
            }
            catch {
                Write-Error "Error installing $($Provider)"
                Write-Error "$($PSItem.Exception)"
                continue
            }
        }
        else {
            Write-Output "$($provider) is already installed." -ForegroundColor Green
        }
    }

    # Setting PSGallery as trusted if it's not trusted
    Write-Output "Making sure that PSGallery is set to Trusted..."
    if ((Get-PSRepository -name PSGallery | Select-Object InstallationPolicy -ExpandProperty InstallationPolicy) -eq "Untrusted") {
        try {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
            Write-Output "PSGallery is now set to trusted" -ForegroundColor Green
        }
        catch {
            Write-Error "Error could not set PSGallery to trusted"
            Write-Error "$($PSItem.Exception)"
            continue
        }
    }
    else {
        Write-Output "PSGallery is already trusted" -ForegroundColor Green
    }

    # Checks if all modules in $NeededModules are installed and up to date.
    Write-Output "Making sure that all of the needed modules are installed and up to date..."
    foreach ($m in $NeededModules) {
        if ($m -in $CurrentModules.Name) {
            # Collects the latest version of module
            $NewestVersion = Find-Module -Name $m | Sort-Object Version -Descending | Select-Object Version -First 1
            # Get all the installed modules and versions
            $AllVersions = Get-InstalledModule -Name $m -AllVersions | Sort-Object PublishedDate -Descending
            $MostRecentVersion = $AllVersions[0].Version

            # Check if the module are up to date
            if ($NewestVersion.Version -gt $AllVersions.Version) {
                try {
                    Write-Output "Updating $($m) to version $($NewestVersion.Version)..."
                    Update-Module -Name $($m) -Force
                    Write-Output "$($m) has been updated!" -ForegroundColor Green
                }
                catch {
                    Write-Error "$($PSItem.Exception)"
                    continue
                }

                # Remove old versions of the modules
                if ($AllVersions.Count -gt 1 ) {
                    Foreach ($Version in $AllVersions) {
                        if ($Version.Version -ne $MostRecentVersion) {
                            try {
                                Write-Output "Uninstalling previous version $($Version.Version) of module $($m)..."
                                Uninstall-Module -Name $m -RequiredVersion $Version.Version -Force -ErrorAction SilentlyContinue
                                Write-Output "$($m) are not uninstalled!" -ForegroundColor Green
                            }
                            catch {
                                Write-Output "Error uninstalling previous version $($Version.Version) of module $($m)" -ForegroundColor Red
                                Write-Output "$($PSItem.Exception)" -ForegroundColor Red
                                continue
                            }
                        }
                    }
                }
            }
            else {
                Write-Output "$($m) don't need to be updated as it's on the latest version" -ForegroundColor Green
            }
        }
        else {
            # Installing missing module
            Write-Output "Installing module $($m) as it's missing..."
            try {
                Install-Module -Name $m -Scope AllUsers -Force
                Write-Output "$($m) are now installed!" -ForegroundColor Green
            }
            catch {
                Write-Output "Could not install $($m)" -ForegroundColor Red
                Write-Output "$($PSItem.Exception)" -ForegroundColor Red
                continue
            }
        }
    }
    # Collect all of the imported modules.
    $ImportedModules = get-module | Select-Object Name, Version
    
    # Import module if it's not imported
    foreach ($module in $NeededModules) {
        if ($module -eq "MSIPatches" -and $PSVersion -gt 5) {
            Write-Output "Remove-MSPatches only works with PowerShell 5.1, skipping it." -ForegroundColor Yellow
        }
        else {
            if ($module -in $ImportedModules.Name) {
                Write-Output "$($Module) are already imported!" -ForegroundColor Green
            }
            else {
                try {
                    Write-Output "Importing $($module) module..."
                    Import-Module -Name $module -Force
                    Write-Output "$($module) are now imported!" -ForegroundColor Green
                }
                catch {
                    Write-Output "Could not import module $($module)" -ForegroundColor Red
                    Write-Output "$($PSItem.Exception)" -ForegroundColor Red
                    continue
                }
            }
        }
    }
}

Function Remove-MSPatches {
    if ($PSVersion -gt 5) {
        Write-Warning "Remove-MSPatches only works with PowerShell 5.1, skipping this function."
    }
    else {
        Write-Output "`n=== Delete all orphaned patches ===`n"
        $OrphanedPatch = Get-OrphanedPatch
        if ($Null -ne $OrphanedPatch) {
            $FreeUp = Get-MsiPatch | select-object OrphanedPatchSize -ExpandProperty OrphanedPatchSize
            Write-Output "This will free up: $($FreeUp)GB"
            try {
                Write-Output "Deleting all of the orphaned patches..."
                Get-OrphanedPatch | Remove-Item
                Write-Output "Success, all of the orphaned patches has been deleted!" -ForegroundColor Green
            }
            catch {
                Write-Error "$($PSItem.Exception)"
                continue
            }
        }
        else {
            Write-Output "No orphaned patches was found." -ForegroundColor Green
        }
    }
}

Function Update-MSUpdates {
    Write-Output "`n=== Windows Update and Windows Store ===`n"
    #Update Windows Store apps!
    if ($PSVersion -gt 5) {
        Write-Warning "Microsoft store updates only works with PowerShell 5.1, skipping this function."
    }
    else {
        try {
            Write-Output "Checking if Windows Store has any updates..."
            $namespaceName = "root\cimv2\mdm\dmmap"
            $className = "MDM_EnterpriseModernAppManagement_AppManagement01"
            $wmiObj = Get-WmiObject -Namespace $namespaceName -Class $className
            $result = $wmiObj.UpdateScanMethod()
            Write-Output "$($result)" -ForegroundColor Green
            Write-Output "Success, checking and if needed updated Windows Store apps!" -ForegroundColor Green
        }
        catch {
            Write-Error "$($PSItem.Exception)"
            continue
        }
    }

    # Checking after Windows Updates
    try {
        Write-Output "Starting to search after Windows Updates..."
        $WSUSUpdates = Get-WindowsUpdate
        if ($Null -ne $WSUSUpdates) {
            Install-WindowsUpdate -AcceptAll
            Write-Output "All of the Windows Updates has been installed!" -ForegroundColor Green
        }
        else {
            Write-Output "All of the latest updates has been installed already! Your up to date!" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "$($PSItem.Exception)"
        continue
    }
}

Function Update-MSDefender {
    Write-Output "`n=== Microsoft Defender ===`n"
    try {
        Write-Output "Update signatures from Microsoft Update Server..."
        Update-MpSignature -UpdateSource MicrosoftUpdateServer
        Write-Output "Updated signatures complete!" -ForegroundColor Green
    }
    catch {
        Write-Error "$($PSItem.Exception)"
        continue
    }


    try {
        Write-Output "Starting Defender Quick Scan, please wait..."
        Start-MpScan -ScanType QuickScan -ErrorAction SilentlyContinue
        Write-Output "Defender quick scan is completed!" -ForegroundColor Green
    }
    catch {
        Write-Error "$($PSItem.Exception)"
        continue
    }
}

function Remove-TempFolderFiles {
    Write-Output "`n=== Starting to delete temp files and folders ===`n"
    $WindowsOld = "C:\Windows.old"
    $Users = Get-ChildItem -Path C:\Users | select-object name -ExpandProperty Name
    $TempFolders = @("C:\Temp", "C:\Tmp", "C:\Windows\Temp", "C:\Windows\Prefetch", "C:\Windows\SoftwareDistribution\Download")
    $SpecialFolders = @("C:\`$Windows`.~BT", "C:\`$Windows`.~WS")

    try {
        Write-Output "Stopping wuauserv..."
        Stop-Service -Name 'wuauserv'
        do {
            Write-Output 'Waiting for wuauserv to stop...'
            Start-Sleep -s 1

        } while (Get-Process wuauserv -ErrorAction SilentlyContinue)
        Write-Output "Wuauserv is now stopped!" -ForegroundColor Green
    }
    catch {
        Write-Error "$($PSItem.Exception)"
        continue   
    }

    foreach ($TempFolder in $TempFolders) {
        if (Test-Path -Path $TempFolder) {
            try {
                Write-Output "Deleting all files in $($TempFolder)..."
                Remove-Item "$($TempFolder)\*" -Recurse -Force -Confirm:$false
                Write-Output "All files in $($TempFolder) has been deleted!" -ForegroundColor Green
            }
            catch {
                Write-Error "$($PSItem.Exception)"
                continue   
            }
        }  
    }

    Try {
        Write-Output "Starting wuauserv again..."
        Start-Service -Name 'wuauserv'
        Write-Output "Wuauserv has started again!" -ForegroundColor Green
    }
    catch {
        Write-Error "$($PSItem.Exception)"
        continue   
    }

    foreach ($usr in $Users) {
        $UsrTemp = "C:\Users\$($usr)\AppData\Local\Temp"
        if (Test-Path -Path $UsrTemp) {
            try {
                Write-Output "Deleting all files in $($UsrTemp)..."
                Remove-Item "$($UsrTemp)\*" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                Write-Output "All files in $($UsrTemp) has been deleted!" -ForegroundColor Green
            }
            catch {
                Write-Error "$($PSItem.Exception)"
                continue   
            }
        } 
    }

    if (Test-Path -Path $WindowsOld) {
        try {
            Write-Output "Deleting folder $($WindowsOld)..."
            Remove-Item "$($WindowsOld)\" -Recurse -Force -Confirm:$false
            Write-Output "The folder $($WindowsOld) has been deleted!" -ForegroundColor Green
        }
        catch {
            Write-Error "$($PSItem.Exception)"
            continue   
        }
    }

    foreach ($sFolder in $SpecialFolders) {
        if (Test-Path -Path $sFolder) {
            try {
                takeown /F "$($sFolder)\*" /R /A
                icacls "$($sFolder)\*.*" /T /grant administrators:F
                Write-Output "Deleting folder $($sFolder)\..."
                Remove-Item "$($sFolder)\" -Recurse -Force -Confirm:$False
                Write-Output "Folder $($sFolder)\* has been deleted!" -ForegroundColor Green
            }
            catch {
                Write-Error "$($PSItem.Exception)"
                continue   
            }
        }
    }

}

Function Start-CleanDisk {
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]$Section
    )

    $sections = @(
        'Active Setup Temp Folders',
        'BranchCache',
        'Content Indexer Cleaner',
        'Device Driver Packages',
        'Downloaded Program Files',
        'GameNewsFiles',
        'GameStatisticsFiles',
        'GameUpdateFiles',
        'Internet Cache Files',
        'Memory Dump Files',
        'Offline Pages Files',
        'Old ChkDsk Files',
        'Previous Installations',
        'Recycle Bin',
        'Service Pack Cleanup',
        'Setup Log Files',
        'System error memory dump files',
        'System error minidump files',
        'Temporary Files',
        'Temporary Setup Files',
        'Temporary Sync Files',
        'Thumbnail Cache',
        'Update Cleanup',
        'Upgrade Discarded Files',
        'User file versions',
        'Windows Defender',
        'Windows Error Reporting Archive Files',
        'Windows Error Reporting Queue Files',
        'Windows Error Reporting System Archive Files',
        'Windows Error Reporting System Queue Files',
        'Windows ESD installation files',
        'Windows Upgrade Log Files'
    )

    if ($PSBoundParameters.ContainsKey('Section')) {
        if ($Section -notin $sections) {
            throw "The section [$($Section)] is not available. Available options are: [$($Section -join ',')]."
        }
    }
    else {
        $Section = $sections
    }

    Write-Verbose -Message 'Clearing CleanMgr.exe automation settings.'

    $getItemParams = @{
        Path        = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\*'
        Name        = 'StateFlags0001'
        ErrorAction = 'SilentlyContinue'
    }
    Get-ItemProperty @getItemParams | Remove-ItemProperty -Name StateFlags0001 -ErrorAction SilentlyContinue

    Write-Verbose -Message 'Adding enabled disk cleanup sections...'
    foreach ($keyName in $Section) {
        $newItemParams = @{
            Path         = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\$keyName"
            Name         = 'StateFlags0001'
            Value        = 1
            PropertyType = 'DWord'
            ErrorAction  = 'SilentlyContinue'
        }
        $null = New-ItemProperty @newItemParams
    }

    Write-Verbose -Message 'Starting CleanMgr.exe...'
    Start-Process -FilePath CleanMgr.exe -ArgumentList '/sagerun:1' -NoNewWindow

    Write-Verbose -Message 'Waiting for CleanMgr and DismHost processes...'
    Get-Process -Name cleanmgr, dismhost -ErrorAction SilentlyContinue | Wait-Process
}


Find-NeededModules
Remove-MSPatches
Remove-TempFolderFiles
Start-CleanDisk
Update-MSDefender
Update-MSUpdates

Write-Output "The script is finished!"
$RebootNeeded = Get-WURebootStatus | Select-Object RebootRequired -ExpandProperty RebootRequired
if ($RebootNeeded -eq "true") {
    Write-Warning "Windows Update want you to reboot your computer, so please do that!" -ForegroundColor Yellow
}