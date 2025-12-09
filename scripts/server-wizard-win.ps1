& {

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
    param([string]$Prompt, [string]$Default = $null)
    while ($true) {
        if ($Default) {
            $input = Read-Host "$Prompt [$Default]"
            if ([string]::IsNullOrWhiteSpace($input)) { return $Default }
        } else {
            $input = Read-Host $Prompt
        }
        if ($input) { return $input.Trim() }
    }
}

function Confirm-YesNo {
    param([string]$Prompt, [bool]$DefaultYes = $true)
    $hint = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    $answer = Read-Host "$Prompt $hint"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $DefaultYes }
    return $answer.ToLower() -in @("y","yes")
}

function Choose-File {
    param([string]$Title = "Select a file")
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Title = $Title
    if ($d.ShowDialog() -eq "OK") { return $d.FileName }
    return $null
}

function Ensure-SshTools {
    if (-not (Get-Command ssh -ErrorAction SilentlyContinue)) {
        Show-Banner
        Write-Host "ERROR: SSH not installed." -ForegroundColor $Red
        Read-Host "Install OpenSSH client and press Enter"
        exit
    }
}

function Get-KnownHostsPath {
    $dir = Join-Path $env:USERPROFILE ".ssh"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    return (Join-Path $dir "known_hosts")
}

function Ensure-SshKeySelection {
    if ($null -eq $script:Config.UseSshKey) {
        $script:Config.UseSshKey = Confirm-YesNo "Use SSH private key?" $true
    }
    if ($script:Config.UseSshKey -and -not $script:Config.SshKeyPath) {
        $key = Choose-File "Select SSH private key"
        if (-not $key) { throw "Cancelled." }
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
    $args = Get-SshArgs
    return $args
}

# ====================== FIXED SSH EXECUTION ======================

function Invoke-RemoteInteractive {
    param($User, $ServerHost, $Command)

    $ssh = Get-RealSshExe
    $args = Get-SshArgs
    $args += @("$User@$ServerHost", $Command)

    Write-Host "`nRunning interactive SSH on $User@$ServerHost" -ForegroundColor Gray
    Write-Host "COMMAND: $Command"

    Start-Process -FilePath $ssh -ArgumentList $args -NoNewWindow -Wait
}

function Invoke-RemoteCapture {
    param($User, $ServerHost, $Command)

    $ssh = Get-RealSshExe
    $args = Get-SshArgs
    $args += @("$User@$ServerHost", $Command)

    $tmp = [System.IO.Path]::GetTempFileName()

    Start-Process -FilePath $ssh `
        -ArgumentList $args `
        -RedirectStandardOutput $tmp `
        -RedirectStandardError $tmp `
        -WindowStyle Hidden `
        -Wait

    $out = Get-Content $tmp
    Remove-Item $tmp -ErrorAction SilentlyContinue
    return ($out -join "`n")
}

# ======================= MAIN WIZARD =======================

function Run-CreateNonRootUser {
    Show-Banner
    Write-Host "CREATE NON-ROOT USER" -ForegroundColor $Cyan
    Write-Host ""

    Ensure-ConfigField -Field "NewServerHost" -Prompt "New server IP"
    Ensure-ConfigField -Field "NewServerUser" -Prompt "User to connect as"

    if (-not (Confirm-YesNo "Run setup on remote server?")) { return }

    $cmd = "rm -f create_sudo_user.sh && curl -fsSL -o create_sudo_user.sh https://github.com/StardustCollective/NodeCloud/raw/main/scripts/create_sudo_user.sh && sudo bash create_sudo_user.sh"

    Invoke-RemoteInteractive $script:Config.NewServerUser $script:Config.NewServerHost $cmd

    Pause-AnyKey
}

function Backup-P12FromOldServer {
    Show-Banner
    Write-Host "P12 BACKUP" -ForegroundColor $Cyan

    Ensure-ConfigField -Field "OldServerHost" -Prompt "Old server IP"
    Ensure-ConfigField -Field "OldServerUser" -Prompt "Old server user"

    $find = "find /root /home /var/tessellation /opt -maxdepth 5 -iname '*.p12'"
    $out = Invoke-RemoteCapture $script:Config.OldServerUser $script:Config.OldServerHost $find

    Write-Host "`nFOUND:`n$out"
    Pause-AnyKey
}

function Upload-P12ToNewServer {
    Show-Banner
    Write-Host "Launching Windows P12 uploader..." -ForegroundColor $Cyan
    iex (iwr "https://github.com/StardustCollective/NodeCloud/raw/main/scripts/uploadP12/windows/upload-p12.ps1").Content
    Pause-AnyKey
}

function Full-Flow {
    if (Confirm-YesNo "1) Create non-root user?") { Run-CreateNonRootUser }
    if (Confirm-YesNo "2) Backup P12 from old server?") { Backup-P12FromOldServer }
    if (Confirm-YesNo "3) Upload P12 to new server?") { Upload-P12ToNewServer }
    Pause-AnyKey
}

function Show-Menu {
    $items = @(
        "Full Guided Flow",
        "Create Non-Root User",
        "Backup P12 from Old Server",
        "Upload P12 to New Server",
        "Exit"
    )
    $i = 0
    while ($true) {
        Show-Banner
        Write-Host "Use Up/Down + Enter" -ForegroundColor Gray
        Write-Host ""
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
                    1 { Run-CreateNonRootUser }
                    2 { Backup-P12FromOldServer }
                    3 { Upload-P12ToNewServer }
                    4 { return }
                }
            }
        }
    }
}

# ======================= START =======================

Ensure-SshTools
Show-Banner
Write-Host "This wizard runs on your Windows PC." -ForegroundColor $Cyan
Read-Host "Press Enter to continue..."
Show-Menu

}  # END WRAPPING BLOCK
