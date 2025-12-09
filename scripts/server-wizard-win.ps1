$scriptContent = @"
# ============================================================
#  STARDUST COLLECTIVE
#  SERVER SETUP WIZARD (WINDOWS - LOCAL)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms

# ----------------- VARIABLES -----------------

$Cyan   = "Cyan"
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"
$Gray   = "Gray"

$Config = @{
    OldServerHost = $null
    OldServerUser = $null
    NewServerHost = $null
    NewServerUser = $null
    UseSshKey     = $null
    SshKeyPath    = $null
    LocalP12Path  = $null
}

# ----------------- HELPERS -----------------

function Show-Banner {
    Clear-Host
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host "                    STARDUST COLLECTIVE" -ForegroundColor Cyan
    Write-Host "              SERVER SETUP WIZARD (WINDOWS)" -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Pause-AnyKey {
    Write-Host ""
    Write-Host "Press any key to return..." -ForegroundColor Gray
    [Console]::ReadKey($true) | Out-Null
}

function Read-NonEmpty {
    param([string]$Prompt, [string]$Default = $null)
    while ($true) {
        $input = if ($Default) { Read-Host "$Prompt [$Default]" } else { Read-Host $Prompt }
        if ([string]::IsNullOrWhiteSpace($input)) {
            if ($Default) { return $Default }
        } else {
            return $input.Trim()
        }
    }
}

function Confirm-YesNo {
    param([string]$Prompt, [bool]$DefaultYes = $true)
    $hint = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    $input = Read-Host "$Prompt $hint"
    if ([string]::IsNullOrWhiteSpace($input)) { return $DefaultYes }
    return $input.ToLower() -in @("y","yes")
}

function Choose-File {
    param([string]$Title = "Select File")
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $Title
    if ($dlg.ShowDialog() -eq "OK") { return $dlg.FileName }
    return $null
}

function Ensure-SshTools {
    if (-not (Get-Command ssh -ErrorAction SilentlyContinue)) {
        Show-Banner
        Write-Host "ERROR: OpenSSH Client not installed." -ForegroundColor Red
        Read-Host "Install it and press Enter."
        exit
    }
}

function Get-KnownHostsPath {
    $dir = Join-Path $env:USERPROFILE ".ssh"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return (Join-Path $dir "known_hosts")
}

function Ensure-SshKeySelection {
    if ($null -eq $Config.UseSshKey) {
        $Config.UseSshKey = Confirm-YesNo "Use SSH private key?" $true
    }

    if ($Config.UseSshKey -and -not $Config.SshKeyPath) {
        $key = Choose-File "Select SSH key"
        if (-not $key) { throw "SSH key selection cancelled" }
        $Config.SshKeyPath = $key
    }
}

function Get-SshArgs {
    Ensure-SshTools
    Ensure-SshKeySelection

    $args = @()
    if ($Config.UseSshKey) {
        $args += @("-i", $Config.SshKeyPath)
    }
    $known = Get-KnownHostsPath
    $args += @("-o","StrictHostKeyChecking=no","-o","UserKnownHostsFile=$known")
    return $args
}

# ----------------- SSH EXECUTION (FIXED) -----------------

function SSH-Interactive {
    param($User, $Host, $Command)

    $exe = (Get-Command ssh).Source
    $args = Get-SshArgs
    $args += @("$User@$Host", $Command)

    Start-Process -FilePath $exe -ArgumentList $args -Wait -NoNewWindow
}

function SSH-Capture {
    param($User, $Host, $Command)

    $exe = (Get-Command ssh).Source
    $args = Get-SshArgs
    $args += @("$User@$Host", $Command)

    $tmp = [IO.Path]::GetTempFileName()

    Start-Process -FilePath $exe `
        -ArgumentList $args `
        -RedirectStandardOutput $tmp `
        -RedirectStandardError $tmp `
        -WindowStyle Hidden `
        -Wait

    $out = Get-Content $tmp
    Remove-Item $tmp
    return ($out -join "`n")
}

# ----------------- LOGIC -----------------

function Ensure-ConfigField {
    param([string]$Field, [string]$Prompt)
    if (-not $Config[$Field]) {
        $Config[$Field] = Read-NonEmpty $Prompt
    }
}

function Run-CreateUser {
    Show-Banner
    Write-Host "CREATE NON-ROOT USER" -ForegroundColor Cyan

    Ensure-ConfigField "NewServerHost" "New server IP"
    Ensure-ConfigField "NewServerUser" "User to connect as"

    if (-not (Confirm-YesNo "Run script on server?")) { return }

    $cmd = "rm -f create_sudo_user.sh && curl -fsSL -o create_sudo_user.sh https://github.com/StardustCollective/NodeCloud/raw/main/scripts/create_sudo_user.sh && sudo bash create_sudo_user.sh"

    SSH-Interactive $Config.NewServerUser $Config.NewServerHost $cmd

    Pause-AnyKey
}

function Backup-P12 {
    Show-Banner
    Write-Host "Searching for P12..." -ForegroundColor Cyan

    Ensure-ConfigField "OldServerHost" "Old server IP"
    Ensure-ConfigField "OldServerUser" "Old server user"

    $cmd = "find /root /home /var/tessellation /opt -maxdepth 5 -iname '*.p12'"
    $out = SSH-Capture $Config.OldServerUser $Config.OldServerHost $cmd

    Write-Host $out
    Pause-AnyKey
}

function Upload-P12 {
    Show-Banner
    Write-Host "Launching WINDOWS P12 upload tool..." -ForegroundColor Cyan
    iex (iwr "https://github.com/StardustCollective/NodeCloud/raw/main/scripts/uploadP12/windows/upload-p12.ps1").Content
    Pause-AnyKey
}

function Full-Flow {
    if (Confirm-YesNo "1) Create user?") { Run-CreateUser }
    if (Confirm-YesNo "2) Backup P12?") { Backup-P12 }
    if (Confirm-YesNo "3) Upload P12?") { Upload-P12 }
}

function Show-Menu {
    $items = @(
        "Full Guided Flow",
        "Create Non-Root User",
        "Backup P12",
        "Upload P12",
        "Exit"
    )

    $i = 0

    while ($true) {
        Show-Banner
        Write-Host "Use ↑ ↓ and Enter:`n"

        for ($x=0; $x -lt $items.Count; $x++) {
            if ($x -eq $i) { Write-Host "> $($items[$x])" -ForegroundColor Green }
            else { Write-Host "  $($items[$x])" }
        }

        switch (([Console]::ReadKey($true)).Key) {
            "UpArrow"   { $i = ($i - 1 + $items.Count) % $items.Count }
            "DownArrow" { $i = ($i + 1) % $items.Count }
            "Enter" {
                switch ($i) {
                    0 { Full-Flow }
                    1 { Run-CreateUser }
                    2 { Backup-P12 }
                    3 { Upload-P12 }
                    4 { return }
                }
            }
        }
    }
}

# ----------------- START -----------------

Ensure-SshTools
Show-Banner
Read-Host "Press Enter to continue..."
Show-Menu

"@

iex $scriptContent
