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
    SshKeyPath    = $null
    LocalP12Path  = $null
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
        if ($null -ne $Default -and $Default -ne "") {
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
    $answer = $answer.ToLower()
    return ($answer -eq "y" -or $answer -eq "yes")
}

function Choose-File {
    param(
        [string]$Title = "Select a file"
    )
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = $Title
    $ofd.Filter = "All files (*.*)|*.*"
    $ofd.Multiselect = $false

    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $ofd.FileName
    }
    return $null
}

function Ensure-ConfigField {
    param(
        [string]$Field,
        [string]$Prompt
    )
    if (-not $script:Config[$Field]) {
        $script:Config[$Field] = Read-NonEmpty $Prompt
    }
}

function Ensure-SshKey {
    if (-not $script:Config.SshKeyPath) {
        Write-Host ""
        Write-Host "Select your SSH private key used to connect to your server." -ForegroundColor $Cyan
        $path = Choose-File -Title "Select SSH private key"
        if (-not $path) {
            throw "SSH key selection cancelled."
        }
        $script:Config.SshKeyPath = $path
    }
}

function Get-SshBase {
    Ensure-SshKey
    return @("ssh", "-i", $script:Config.SshKeyPath, "-o", "StrictHostKeyChecking=no")
}

function Get-ScpBase {
    Ensure-SshKey
    return @("scp", "-i", $script:Config.SshKeyPath, "-o", "StrictHostKeyChecking=no")
}

function Invoke-Remote {
    param(
        [string]$User,
        [string]$ServerHost,
        [string]$Command
    )

    $target = "$($User)@$($ServerHost)"
    $args   = (Get-SshBase) + @($target, $Command)

    Write-Host ""
    Write-Host ("Running on {0}@{1}:" -f $User, $ServerHost) -ForegroundColor $Gray
    Write-Host "  $Command" -ForegroundColor $Gray

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $args[0]
    $psi.Arguments              = ($args[1..($args.Count-1)] -join " ")
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null   = $proc.Start()
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($proc.ExitCode -ne 0) {
        if ($stderr) { Write-Host $stderr -ForegroundColor $Red }
        throw "Remote command failed with exit code $($proc.ExitCode)."
    }
    return $stdout
}

function Run-CreateNonRootUser {
    Show-Banner
    Write-Host "CREATE NON-ROOT SUDO USER ON NEW SERVER" -ForegroundColor $Cyan
    Write-Host ""

    Ensure-ConfigField -Field "NewServerHost" -Prompt "New server IP / hostname"
    Ensure-ConfigField -Field "NewServerUser" -Prompt "User to connect as (usually root on a fresh server)"

    if (-not (Confirm-YesNo "Run create_sudo_user.sh on $($script:Config.NewServerUser)@$($script:Config.NewServerHost)?")) {
        Write-Host "Skipped." -ForegroundColor $Yellow
        Pause-AnyKey
        return
    }

    $cmd = "rm -f create_sudo_user.sh && curl -fsSL -o create_sudo_user.sh https://github.com/StardustCollective/NodeCloud/raw/main/scripts/create_sudo_user.sh && sudo bash create_sudo_user.sh"
    try {
        $null = Invoke-Remote -User $script:Config.NewServerUser -ServerHost $script:Config.NewServerHost -Command $cmd
        Write-Host ""
        Write-Host "create_sudo_user.sh completed. Reconnect as your new sudo user (e.g. nodeadmin) before continuing." -ForegroundColor $Green
    } catch {
        Write-Host ""
        Write-Host "Failed to run create_sudo_user.sh:" -ForegroundColor $Red
        Write-Host $_ -ForegroundColor $Red
    }
    Pause-AnyKey
}

function Select-FromList {
    param(
        [string[]]$Items,
        [string]$Title = "Select an item"
    )

    if (-not $Items -or $Items.Count -eq 0) { return $null }

    $index = 0
    while ($true) {
        Show-Banner
        Write-Host $Title -ForegroundColor $Cyan
        Write-Host ""
        Write-Host "Use ↑ / ↓ arrows and Enter to select. Press Esc to cancel." -ForegroundColor $Gray
        Write-Host ""

        for ($i = 0; $i -lt $Items.Count; $i++) {
            if ($i -eq $index) {
                Write-Host ("> " + $Items[$i]) -ForegroundColor $Green
            } else {
                Write-Host ("  " + $Items[$i])
            }
        }

        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            "UpArrow"   { $index = ($index - 1); if ($index -lt 0) { $index = $Items.Count - 1 } }
            "DownArrow" { $index = ($index + 1); if ($index -ge $Items.Count) { $index = 0 } }
            "Enter"     { return $Items[$index] }
            "Escape"    { return $null }
        }
    }
}

function Backup-P12FromOldServer {
    Show-Banner
    Write-Host "BACKUP P12 FROM OLD SERVER" -ForegroundColor $Cyan
    Write-Host ""
    Write-Host "This will:" -ForegroundColor $Cyan
    Write-Host "  1) SSH to your OLD server"
    Write-Host "  2) Search common paths for *.p12 (depth 5, skipping hash/ordinal)"
    Write-Host "  3) Let you choose which P12 to back up"
    Write-Host "  4) Copy it into ~/p12-backups on the old server"
    Write-Host "  5) Download a copy to your Windows machine"
    Write-Host ""

    Ensure-ConfigField -Field "OldServerHost" -Prompt "Old server IP / hostname"
    Ensure-ConfigField -Field "OldServerUser" -Prompt "User to connect as (e.g. nodeadmin or root)"

    if (-not (Confirm-YesNo "Start the scan on $($script:Config.OldServerUser)@$($script:Config.OldServerHost)?")) {
        Write-Host "Cancelled." -ForegroundColor $Yellow
        Pause-AnyKey
        return
    }

    $findCmd = @"
find /root /home /var/tessellation /opt -maxdepth 5 \( -name hash -o -name ordinal \) -prune -o -type f -iname '*.p12' -print 2>/dev/null
"@.Trim()

    try {
        $output = Invoke-Remote -User $script:Config.OldServerUser -ServerHost $script:Config.OldServerHost -Command $findCmd
    } catch {
        Write-Host ""
        Write-Host "Failed to search for P12 files:" -ForegroundColor $Red
        Write-Host $_ -ForegroundColor $Red
        Pause-AnyKey
        return
    }

    $paths = $output -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if (-not $paths -or $paths.Count -eq 0) {
        Write-Host ""
        Write-Host "No .p12 files found in the common search locations." -ForegroundColor $Red
        Pause-AnyKey
        return
    }

    $selected = Select-FromList -Items $paths -Title "Select the P12 to back up:"
    if (-not $selected) {
        Write-Host "Selection cancelled." -ForegroundColor $Yellow
        Pause-AnyKey
        return
    }

    $fileName         = [System.IO.Path]::GetFileName($selected)
    $remoteBackupDir  = "~/p12-backups"
    $remoteBackupPath = "$remoteBackupDir/$fileName"

    $backupCmd = "mkdir -p $remoteBackupDir && if [ -e '$remoteBackupPath' ]; then echo 'EXISTS'; else echo 'OK'; fi"
    $status   = Invoke-Remote -User $script:Config.OldServerUser -ServerHost $script:Config.OldServerHost -Command $backupCmd
    $status   = $status.Trim()

    if ($status -eq "EXISTS") {
        if (-not (Confirm-YesNo "Remote backup $remoteBackupPath already exists. Overwrite?" $false)) {
            Write-Host "Remote backup skipped." -ForegroundColor $Yellow
            Pause-AnyKey
            return
        }
    }

    $copyCmd = "mkdir -p $remoteBackupDir && cp -f '$selected' '$remoteBackupPath'"
    try {
        $null = Invoke-Remote -User $script:Config.OldServerUser -ServerHost $script:Config.OldServerHost -Command $copyCmd
    } catch {
        Write-Host ""
        Write-Host "Failed to copy P12 into backup folder on old server:" -ForegroundColor $Red
        Write-Host $_ -ForegroundColor $Red
        Pause-AnyKey
        return
    }

    Write-Host ""
    Write-Host "Backed up on old server at: $remoteBackupPath" -ForegroundColor $Green

    $localDir = Join-Path $PWD "p12-backups"
    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir | Out-Null
    }

    $localPath = Join-Path $localDir $fileName
    if (Test-Path $localPath) {
        if (-not (Confirm-YesNo "Local file $localPath already exists. Overwrite?" $false)) {
            Write-Host "Download skipped; existing local file preserved." -ForegroundColor $Yellow
            Pause-AnyKey
            return
        }
    }

    $scpBase = Get-ScpBase
    $scpArgs = $scpBase + @("$($script:Config.OldServerUser)@$($script:Config.OldServerHost):p12-backups/$fileName", $localPath)

    Write-Host ""
    Write-Host "Downloading backup to: $localPath" -ForegroundColor $Cyan
    & $scpArgs[0] $scpArgs[1..($scpArgs.Count-1)]
    if ($LASTEXITCODE -ne 0) {
        Write-Host "scp failed with exit code $LASTEXITCODE" -ForegroundColor $Red
    } else {
        Write-Host "P12 backup downloaded to: $localPath" -ForegroundColor $Green
        $script:Config.LocalP12Path = $localPath
    }

    Pause-AnyKey
}

function Upload-P12ToNewServer {
    Show-Banner
    Write-Host "UPLOAD P12 TO NEW SERVER (CALLS EXISTING WINDOWS TOOL)" -ForegroundColor $Cyan
    Write-Host ""
    Write-Host "This will download and run your existing Windows P12 upload wizard." -ForegroundColor $Cyan
    Write-Host "That tool will:" -ForegroundColor $Cyan
    Write-Host "  1) Let you browse for your .p12"
    Write-Host "  2) Verify the password"
    Write-Host "  3) Ask for server info"
    Write-Host "  4) Upload the P12 to the server"
    Write-Host ""

    if (-not (Confirm-YesNo "Continue and run the Windows P12 upload tool now?")) {
        Write-Host "Cancelled." -ForegroundColor $Yellow
        Pause-AnyKey
        return
    }

    try {
        iex (iwr "https://github.com/StardustCollective/NodeCloud/raw/main/scripts/uploadP12/windows/upload-p12.ps1" -UseBasicParsing).Content
    } catch {
        Write-Host ""
        Write-Host "Failed to run upload-p12.ps1 from GitHub:" -ForegroundColor $Red
        Write-Host $_ -ForegroundColor $Red
    }

    Pause-AnyKey
}

function Full-Flow {
    Show-Banner
    Write-Host "FULL GUIDED FLOW" -ForegroundColor $Cyan
    Write-Host ""
    Write-Host "This will walk you through the main steps in order:" -ForegroundColor $Cyan
    Write-Host "  1) (Optional) Create non-root sudo user on NEW server"
    Write-Host "  2) (Optional) Backup P12 from OLD server"
    Write-Host "  3) Upload P12 to NEW server using the Windows upload tool"
    Write-Host ""

    if (Confirm-YesNo "Step 1: Run create_sudo_user.sh on NEW server now?") {
        Run-CreateNonRootUser
    }

    if (Confirm-YesNo "Step 2: Do you still need to back up the P12 from an OLD server?" $true) {
        Backup-P12FromOldServer
    }

    if (Confirm-YesNo "Step 3: Run the Windows P12 upload wizard to send the P12 to the NEW server?" $true) {
        Upload-P12ToNewServer
    }

    Show-Banner
    Write-Host "Full flow complete." -ForegroundColor $Green
    Write-Host ""
    Write-Host "Once your P12 is on the new server, you can continue with node setup." -ForegroundColor $Cyan
    Pause-AnyKey
}

function Show-Menu {
    $items = @(
        "Full Guided Flow",
        "Create Non-Root User on New Server",
        "Backup P12 from Old Server",
        "Upload P12 to New Server (Windows tool)",
        "Exit"
    )
    $index = 0

    while ($true) {
        Show-Banner
        Write-Host "Use ↑ / ↓ and Enter to choose an option:" -ForegroundColor $Cyan
        Write-Host ""

        for ($i = 0; $i -lt $items.Count; $i++) {
            if ($i -eq $index) {
                Write-Host ("> " + $items[$i]) -ForegroundColor $Green
            } else {
                Write-Host ("  " + $items[$i])
            }
        }

        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            "UpArrow"   { $index = ($index - 1); if ($index -lt 0) { $index = $items.Count - 1 } }
            "DownArrow" { $index = ($index + 1); if ($index -ge $items.Count) { $index = 0 } }
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

Show-Banner
Write-Host "This wizard runs on your WINDOWS PC." -ForegroundColor $Cyan
Write-Host "It will SSH into your servers and run the needed scripts for you." -ForegroundColor $Cyan
Write-Host ""
Read-Host "Press Enter to open the menu..."
Show-Menu
