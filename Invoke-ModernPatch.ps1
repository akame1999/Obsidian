<#
.SYNOPSIS
    OpsCenter - Unified Windows Patch Management Script
.DESCRIPTION
    All-in-one: config, deployment, rollback, and server checks.
    Replaces: Invoke-ModernPatch.ps1 + patch_config.ps1 + patch_rollback.ps1
.ACTIONS
    PreCheck          - Verify server connectivity and OS info
    GetSystemInfo     - Get OS, disk, and uptime details
    Deploy            - Copy and install a patch file
    Rollback          - Uninstall a KB by number
    CancelReboot      - Remove a scheduled reboot task
    GetInstalledPatches - List installed hotfixes
#>

param(
    [Parameter(Mandatory=$true)][string]$Action,
    [Parameter(Mandatory=$true)][PSCredential]$Cred,
    [string]$TargetServer,
    [string]$FileName,
    [string]$RebootTime,
    [string]$KBNumber   # Used by Rollback action
)

# ─── INLINE CONFIG ────────────────────────────────────────────────────────────
# Edit server list and patch share here — single source of truth
$Config = @{
    Servers = @(
        "L11SGRIFP001",
        "L11SGRIFP002",
        "L11SGRIFP003",
        "L11SGRIFP005",
        "L11SGRIVMHDC01",
        "L11SGRIWEB001"
    )
    # Override with env var PATCH_SOURCE_ROOT in production
    PatchSourceRoot = if ($env:PATCH_SOURCE_ROOT) {
        $env:PATCH_SOURCE_ROOT
    } else {
        "\\10.87.60.2\IT Asset Detail\Naseer\Patch"
    }
    Domain = "ZL"
}

# ─── HELPER ───────────────────────────────────────────────────────────────────
function Write-JsonOutput {
    param($Data)
    Write-Output ($Data | ConvertTo-Json -Compress)
}

# ─── GET SYSTEM INFO ──────────────────────────────────────────────────────────
if ($Action -eq "GetSystemInfo") {
    try {
        $info = Invoke-Command -ComputerName $TargetServer -Credential $Cred -ScriptBlock {
            $osObj = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            if ($osObj) {
                $os        = $osObj.Caption
                $uptime    = [math]::Round(((Get-Date) - $osObj.LastBootUpTime).TotalDays, 1)
            } else {
                $os        = "Unknown"
                $uptime    = 0
            }
            $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
            return @{
                Online      = $true
                OS          = $os
                DiskFreeGB  = if ($disk) { [math]::Round($disk.FreeSpace / 1GB, 2) } else { 0 }
                DiskTotalGB = if ($disk) { [math]::Round($disk.Size / 1GB, 2) }      else { 0 }
                UptimeDays  = $uptime
            }
        } -ErrorAction Stop
        Write-JsonOutput $info
    }
    catch {
        Write-JsonOutput @{ Online=$false; OS="ERROR"; DiskFreeGB=0; DiskTotalGB=0; UptimeDays=0; Error=$_.Exception.Message }
        exit 1
    }
    exit
}

# ─── PRE-CHECK ────────────────────────────────────────────────────────────────
if ($Action -eq "PreCheck") {
    try {
        $info = Invoke-Command -ComputerName $TargetServer -Credential $Cred -ScriptBlock {
            $os = "Unknown"
            try {
                $osObj = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
                if ($osObj.Caption) { $os = $osObj.Caption }
            } catch {
                try {
                    $osObj = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
                    if ($osObj.Caption) { $os = $osObj.Caption }
                } catch {}
            }
            $reboot = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -or
                      (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
            return @{ OS = $os; RebootRequired = $reboot }
        } -ErrorAction Stop

        $compat = "Yes"
        if ($FileName) {
            if ($FileName -like "*2016*" -and $info.OS -notlike "*2016*") { $compat = "No" }
            if ($FileName -like "*2019*" -and $info.OS -notlike "*2019*") { $compat = "No" }
            if ($FileName -like "*2022*" -and $info.OS -notlike "*2022*") { $compat = "No" }
        }

        Write-JsonOutput @{
            Server         = $TargetServer
            OS             = $info.OS
            Compatible     = $compat
            RebootRequired = if ($info.RebootRequired) { "Yes" } else { "No" }
            Online         = $true
        }
    }
    catch {
        Write-JsonOutput @{ Server=$TargetServer; OS="ERROR"; Compatible="N/A"; RebootRequired="N/A"; Online=$false; Error=$_.Exception.Message }
        exit 1
    }
    exit
}

# ─── DEPLOY ───────────────────────────────────────────────────────────────────
if ($Action -eq "Deploy") {
    try {
        $LocalPath = Join-Path $Config.PatchSourceRoot $FileName
        if (-not (Test-Path $LocalPath)) {
            Write-Error "Patch file not found: $LocalPath"
            exit 1
        }

        $JustTheFile = Split-Path $FileName -Leaf
        $RemoteTemp  = "C:\Windows\Temp\$JustTheFile"
        $FileExt     = [System.IO.Path]::GetExtension($JustTheFile).ToLower()
        $KBFromFile  = if ($JustTheFile -match "kb(\d+)") { "KB" + $matches[1] } else { $null }
        $TaskName    = "TempPatchInstall_$(Get-Date -Format 'yyyyMMddHHmmss')"

        $session = New-PSSession -ComputerName $TargetServer -Credential $Cred -ErrorAction Stop
        Copy-Item -Path $LocalPath -Destination $RemoteTemp -ToSession $session -ErrorAction Stop

        Invoke-Command -Session $session -ScriptBlock {
            param($path, $taskName, $ext)
            $exe  = if ($ext -eq ".msu") { "wusa.exe" } else { $path }
            $args = if ($ext -eq ".msu") { "$path /quiet /norestart" } else { "/quiet /norestart" }
            $action    = New-ScheduledTaskAction -Execute $exe -Argument $args
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
            Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Force | Out-Null
            Start-ScheduledTask -TaskName $taskName
            $waited = 0
            do {
                Start-Sleep -Seconds 5; $waited += 5
                $state = (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue).State
            } while ($state -eq "Running" -and $waited -lt 600)
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Remove-Item $path -Force -ErrorAction SilentlyContinue
        } -ArgumentList $RemoteTemp, $TaskName, $FileExt

        if ($KBFromFile) {
            $verify = Invoke-Command -Session $session -ScriptBlock {
                param($kb) Get-HotFix -Id $kb -ErrorAction SilentlyContinue
            } -ArgumentList $KBFromFile
            if ($verify) { Write-Host "[INFO] $KBFromFile verified installed" }
        }

        if ($RebootTime) {
            Invoke-Command -Session $session -ScriptBlock {
                param($rt)
                $dt        = [datetime]::ParseExact($rt, "yyyy-MM-dd HH:mm", $null)
                $action    = New-ScheduledTaskAction -Execute "shutdown.exe" -Argument "/r /f /t 0"
                $trigger   = New-ScheduledTaskTrigger -Once -At $dt
                $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
                Unregister-ScheduledTask -TaskName "ScheduledReboot_PatchMaint" -Confirm:$false -ErrorAction SilentlyContinue
                Register-ScheduledTask -TaskName "ScheduledReboot_PatchMaint" -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
            } -ArgumentList $RebootTime
            Write-Host "[INFO] Reboot scheduled for $RebootTime"
        }

        Remove-PSSession $session
        Write-Output (ConvertTo-Json @{ Message="Deployment completed"; Success=$true } -Compress)
    }
    catch {
        Write-Output (ConvertTo-Json @{ Message=$_.Exception.Message; Success=$false } -Compress)
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        exit 1
    }
    exit
}

# ─── ROLLBACK ─────────────────────────────────────────────────────────────────
if ($Action -eq "Rollback") {
    if (-not $KBNumber) { Write-Error "KBNumber is required for Rollback"; exit 1 }

    try {
        $session = New-PSSession -ComputerName $TargetServer -Credential $Cred -ErrorAction Stop

        # Verify KB is installed
        $installed = Invoke-Command -Session $session -ScriptBlock {
            param($kb) Get-HotFix -Id $kb -ErrorAction SilentlyContinue
        } -ArgumentList $KBNumber

        if (-not $installed) {
            Write-Output (ConvertTo-Json @{ Message="$KBNumber not found on $TargetServer"; Success=$false } -Compress)
            Remove-PSSession $session
            exit 0
        }

        $kbNum    = $KBNumber -replace "KB", ""
        $taskName = "TempRollback_$(Get-Date -Format 'yyyyMMddHHmmss')"

        Invoke-Command -Session $session -ScriptBlock {
            param($kb, $taskName)
            $action    = New-ScheduledTaskAction -Execute "wusa.exe" -Argument "/uninstall /kb:$kb /quiet /norestart"
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
            Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Force | Out-Null
            Start-ScheduledTask -TaskName $taskName
            $waited = 0
            do {
                Start-Sleep -Seconds 5; $waited += 5
                $state = (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue).State
            } while ($state -eq "Running" -and $waited -lt 600)
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        } -ArgumentList $kbNum, $taskName

        Start-Sleep -Seconds 15
        $stillThere = Invoke-Command -Session $session -ScriptBlock {
            param($kb) Get-HotFix -Id $kb -ErrorAction SilentlyContinue
        } -ArgumentList $KBNumber

        Remove-PSSession $session

        if ($stillThere) {
            Write-Output (ConvertTo-Json @{ Message="$KBNumber may still be present - verify manually"; Success=$false } -Compress)
        } else {
            Write-Output (ConvertTo-Json @{ Message="$KBNumber successfully removed from $TargetServer"; Success=$true } -Compress)
        }
    }
    catch {
        Write-Output (ConvertTo-Json @{ Message=$_.Exception.Message; Success=$false } -Compress)
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        exit 1
    }
    exit
}

# ─── CANCEL REBOOT ────────────────────────────────────────────────────────────
if ($Action -eq "CancelReboot") {
    try {
        Invoke-Command -ComputerName $TargetServer -Credential $Cred -ScriptBlock {
            $t = "ScheduledReboot_PatchMaint"
            if (Get-ScheduledTask -TaskName $t -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $t -Confirm:$false
                Write-Output "Reboot task cancelled"
            } else {
                Write-Output "No reboot task found"
            }
        }
        Write-Output (ConvertTo-Json @{ Message="Done"; Success=$true } -Compress)
    }
    catch {
        Write-Output (ConvertTo-Json @{ Message=$_.Exception.Message; Success=$false } -Compress)
        exit 1
    }
    exit
}

# ─── GET INSTALLED PATCHES ────────────────────────────────────────────────────
if ($Action -eq "GetInstalledPatches") {
    try {
        $hotfixes = Invoke-Command -ComputerName $TargetServer -Credential $Cred -ScriptBlock {
            Get-HotFix | Select-Object HotFixID, InstalledOn, Description
        } -ErrorAction Stop
        $result = $hotfixes | ForEach-Object {
            @{ HotFixID=$_.HotFixID; InstalledOn=$_.InstalledOn.ToString("yyyy-MM-dd"); Description=$_.Description }
        }
        Write-Output (ConvertTo-Json $result -Compress)
    }
    catch {
        Write-Error "Failed: $($_.Exception.Message)"
        exit 1
    }
    exit
}

# ─── UNKNOWN ACTION ───────────────────────────────────────────────────────────
Write-Error "Unknown action '$Action'. Valid: PreCheck, GetSystemInfo, Deploy, Rollback, CancelReboot, GetInstalledPatches"
exit 1