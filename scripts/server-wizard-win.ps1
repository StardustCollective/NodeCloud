# ============================================================
#  STARDUST COLLECTIVE
#  SERVER SETUP WIZARD (WINDOWS - LOCAL)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms

$Cyan   = "Cyan"
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"
$Gray   = "Gray"

$script:Config = @{
    OldServerHost = $null
    OldServerUser = $null
    NewServerHost = $null
    NewServerUser = $null
    UseSshKey     = $null
    SshKeyPath    = $null
    LocalP12Path  = $null
}

function Get-RealSshExe {
    return (Get-Command ssh -ErrorAction Stop).Source
}

function Get-RealScpExe {
    return (Get-Command scp -ErrorAction Stop).Source
}

function Show-Banner {
    Clear-Host
    Write-Host "==============================================================" -ForegroundColor $Cyan
    Write-Host "                    STARDUST COLLECTIVE" -ForegroundColor $Cyan
    Write-Host "              SERVER SETUP WIZARD (WINDOWS)" -ForegroundColor $Cyan
    Write-Host "==============================================================" -ForegroundColor $Cyan
    Write-Host ""
}

function Pause-AnyKey {
    Write-Host ""
    Write-Host "Press any key to return to the menu..." -ForegroundColor $Gray
    [void][System.Console]::ReadKey($true)
}

function Read-NonEmpty {
    param(
        [string]$Prompt,
        [string]$Default = $null
    )
    while ($true) {
        if ($Default) {
            $input = Read-Host "$Prompt [$Default]"
            if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
        } else {
            $input = Read-Host $Prompt
        }
        if (-not [string]::IsNullOrWhiteSpace($input)) {
            return $input.Trim()
        }
    }
}

function Confirm-YesNo {
    param(
        [string]$Prompt,
        [bool]$DefaultYes = $true
    )

    $hint = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    $answer = Read-Host "$Prompt $hint"

    if ([string]::IsNullOrWhiteSpace($answer)) {
        return $DefaultYes
    }

    return $answer.ToLower() -in @("y","yes")
}

function Choose-File {
    param([string]$Title = "Select a file")

    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = $Title
    $ofd.Filter = "All files (*.*)|*.*"

    if ($ofd.ShowDialog() -eq "OK") { return $ofd.FileName }
    return $null
}

function Ensure-SshTools {
    if (-not (Get-Command ssh -ErrorAction SilentlyContinue)) {
        Show-Banner
        Write-Host "ERROR: ssh not found in PATH." -ForegroundColor $Red
        Read-Host "Install OpenSSH Client and reopen PowerShell. Press Enter to exit."
        exit 1
    }
}

function Get-KnownHostsPath {
    $dir = Join-Path $env:USERPROFILE ".ssh"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return (Join-Path $dir "known_hosts")
}

function Ensure-SshKeySelection {
    if ($null -eq $script:Config.UseSshKey) {
        $script:Config.UseSshKey = Confirm-YesNo "Use SSH private key file?" $true
    }

    if ($script:Config.UseSshKey -and -not $script:Config.SshKeyPath) {
        $key = Choose-File "Select SSH private key"
        if (-not $key) { throw "SSH key selection cancelled." }
        $script:Config.SshKeyPath = $key
    }
}

function Get-SshArgs {
    Ensure-SshTools
    Ensure-SshKeySelection

    $args = @()
    if ($script:Config.UseSshKey) {
        $args += @("-i", $script:Config.SshKeyPath)
    }
    $known = Get-KnownHostsPath
    $args += @("-o","StrictHostKeyChecking=no","-o","UserKnownHostsFile=$known")
    return $args
}

function Get-ScpArgs {
    Ensure-SshTools
    Ensure-SshKeySelection

    $args = @()
    if ($script:Config.UseSshKey) {
        $args += @("-i", $script:Config.SshKeyPath)
    }
    $known = Get-KnownHostsPath
    $args += @("-o","StrictHostKeyChecking=no","-o","UserKnownHostsFile=$known")
    return $args
}

# =================================================================
# CORRECTED: FULL SSH EXECUTION USING Start-Process (NO POPUPS)
# =================================================================

function Invoke-RemoteInteractive {
    param(
        [string]$User,
        [string]$ServerHost,
        [string]$Command
    )

    $sshArgs = Get-SshArgs
    $sshArgs += @("$User@$ServerHost", $Command)

    $debug = "ssh " + ($sshArgs -join " ")
    Write-Host "`nDEBUG SSH COMMAND:" -ForegroundColor Yellow
    Write-Host "  $debug`n" -ForegroundColor Yellow

    $exe = Get-RealSshExe
    Start-Process -FilePath $exe -ArgumentList $sshArgs -NoNewWindow -Wait
}

function Invoke-RemoteCapture {
    param(
        [string]$User,
        [string]$ServerHost,
        [string]$Command
    )

    $sshArgs = Get-SshArgs
    $sshArgs += @("$User@$ServerHost", $Command)

    $debug = "ssh " + ($sshArgs -join " ")
    Write-Host "`nDEBUG SSH CAPTURE COMMAND:" -ForegroundColor Yellow
    Write-Host "  $debug`n" -ForegroundColor Yellow

    $tmp = [System.IO.Path]::GetTempFileName()
    $exe = Get-RealSshExe

    Start-Process -FilePath $exe `
                  -ArgumentList $sshArgs `
                  -RedirectStandardOutput $tmp `
                  -RedirectStandardError $tmp `
                  -WindowStyle Hidden `
                  -Wait

    $content = Get-Content $tmp
    Remove-Item $tmp -ErrorAction SilentlyContinue
    return ($content -join "`n")
}

# =================================================================
# NORMAL WIZARD FUNCTIONS (UNCHANGED)
# =================================================================

function Run-CreateNonRootUser {
    Show-Banner
    Write-Host "CREATE NON-ROOT SUDO USER" -ForegroundColor $Cyan
    Write-Host ""

    Ensure-ConfigField -Field "NewServerHost" -Prompt "New Server IP"
    Ensure-ConfigField -Field "NewServerUser" -Prompt "User to connect as"

    if (-not (Confirm-YesNo "Run create_sudo_user.sh on server?")) { return }

    $cmd = "rm -f create_sudo_user.sh && curl -fsSL -o create_sudo_user.sh https://github.com/StardustCollective/NodeCloud/raw/main/scripts/create_sudo_user.sh && sudo bash create_sudo_user.sh"

    Invoke-RemoteInteractive -User $script:Config.NewServerUser `
                             -ServerHost $script:Config.NewServerHost `
                             -Command $cmd

    Pause-AnyKey
}

function Select-FromList {
    param(
        [string[]]$Items,
        [string]$Title
    )

    $index = 0
    while ($true) {
        Show-Banner
        Write-Host $Title -ForegroundColor $Cyan
        Write-Host ""

        for ($i=0; $i -lt $Items.Count; $i++) {
            if ($i -eq $index) {
                Write-Host "> $($Items[$i])" -ForegroundColor $Green
            } else {
                Write-Host "  $($Items[$i])"
            }
        }

        $key = [Console]::ReadKey($true).Key
        switch ($key) {
            "UpArrow"   { $index = ($index - 1 + $Items.Count) % $Items.Count }
            "DownArrow" { $index = ($index + 1) % $Items.Count }
            "Enter"     { return $Items[$index] }
            "Escape"    { return $null }
        }
    }
}

function Backup-P12FromOldServer {
    Show-Banner
    Write-Host "BACKING UP P12 FROM OLD SERVER" -ForegroundColor $Cyan
    Write-Host ""

    Ensure-ConfigField -Field "OldServerHost" -Prompt "Old Server IP"
    Ensure-ConfigField -Field "OldServerUser" -Prompt "Old Server Username"

    if (-not (Confirm-YesNo "Scan for P12 files on old server?")) { return }

    $findCmd = "find /root /home /var/tessellation /opt -maxdepth 5 \( -name hash -o -name ordinal \) -prune -o -type f -iname '*.p12' -print"

    $output = Invoke-RemoteCapture -User $script:Config.OldServerUser `
                                   -ServerHost $script:Config.OldServerHost `
                                   -Command $findCmd

    $paths = $output -split "`n" | Where-Object { $_ -and $_.Trim() }
    if (-not $paths) { Write-Host "No P12 files found."; Pause-AnyKey; return }

    $selected = Select-FromList -Items $paths -Title "Select the P12 file:"
    if (-not $selected) { return }

    Write-Host "`nSelected: $selected" -ForegroundColor Green

    Pause-AnyKey
}

function Upload-P12ToNewServer {
    Show-Banner
    Write-Host "RUNNING WINDOWS P12 UPLOADER" -ForegroundColor $Cyan
    Write-Host ""

    if (-not (Confirm-YesNo "Launch Windows P12 upload tool now?")) { return }

    iex (iwr "https://github.com/StardustCollective/NodeCloud/raw/main/scripts/uploadP12/windows/upload-p12.ps1" `
        -UseBasicParsing).Content

    Pause-AnyKey
}

function Full-Flow {
    Show-Banner

    if (Confirm-YesNo "Step 1: Create non-root user?") {
        Run-CreateNonRootUser
    }
    if (Confirm-YesNo "Step 2: Backup existing P12 from old server?" $true) {
        Backup-P12FromOldServer
    }
    if (Confirm-YesNo "Step 3: Upload P12 to new server?" $true) {
        Upload-P12ToNewServer
    }

    Show-Banner
    Write-Host "ALL STEPS COMPLETE" -ForegroundColor $Green
    Pause-AnyKey
}

function Show-Menu {
    $Items = @(
        "Full Guided Flow",
        "Create Non-Root User",
        "Backup P12 from Old Server",
        "Upload P12 to New Server",
        "Exit"
    )

    $index = 0

    while ($true) {
        Show-Banner
        Write-Host "Use ↑/↓ and Enter" -ForegroundColor $Gray
        Write-Host ""

        for ($i=0; $i -lt $Items.Count; $i++) {
            if ($i -eq $index) {
                Write-Host "> $($Items[$i])" -ForegroundColor Green
            } else {
                Write-Host "  $($Items[$i])"
            }
        }

        switch (([Console]::ReadKey($true)).Key) {
            "UpArrow"   { $index = ($index - 1 + $Items.Count) % $Items.Count }
            "DownArrow" { $index = ($index + 1) % $Items.Count }
            "Enter" {
                switch ($index) {
                    0 { Full-Flow }
                    1 { Run-CreateNonRootUser }
                    2 { Backup-P12FromOldServer }
                    3 { Upload-P12ToNewServer }
                    4 { return }
                }
            }
        }
    }
}

# start wizard
Ensure-SshTools
Show-Banner
Write-Host "This wizard runs on your WINDOWS PC." -ForegroundColor $Cyan
Read-Host "Press Enter to open menu..."
Show-Menu
