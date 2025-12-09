<#
    server-toolkit.ps1
    Stardust Collective - Server Setup Toolkit (Windows Edition)

    Requirements:
      - PowerShell 5+ (Windows)
      - OpenSSH client (ssh, scp, ssh-keygen) in PATH
      - Optional: openssl for P12 verification
      - Optional: puttygen.exe for PuTTY key export

    Notes:
      - All server actions happen over SSH.
      - Profiles are stored in: $HOME\.ssh\{ProfileName}_ssh_config.txt
      - No passwords are ever written to disk.
#>

# ====================
# BASIC CONFIG
# ====================
$Host.UI.RawUI.WindowTitle = "Stardust Collective - Server Toolkit"

$Script:ProfilesDir   = Join-Path $HOME ".ssh"
$Script:KnownHosts    = Join-Path $HOME ".ssh\known_hosts"
$Script:CachedPass    = @{}        # per-session password cache (not persisted)
$Script:NewServerConn = $null      # per-session connection object for "new server"
$Script:OldServerConn = $null      # per-session connection object for "old server"

# Session-scoped convenience caches (not persisted)
$Script:LastP12Path       = $null  # last selected local .p12 path
$Script:LastIdentityPath  = $null  # last entered SSH private key path

# Ensure .ssh directory exists
if (-not (Test-Path $Script:ProfilesDir)) {
    New-Item -ItemType Directory -Path $Script:ProfilesDir | Out-Null
}

# ====================
# UTILS: COLORS AND PROMPTS
# ====================
function Write-Info($Text)  { Write-Host $Text -ForegroundColor Cyan }
function Write-Ok($Text)    { Write-Host $Text -ForegroundColor Green }
function Write-Warn($Text)  { Write-Host $Text -ForegroundColor Yellow }
function Write-Err($Text)   { Write-Host $Text -ForegroundColor Red }

function Show-Banner {
    Clear-Host
    Write-Info "STARDUST COLLECTIVE - SERVER TOOLKIT"
    Write-Info "----------------------------------------------"
    Write-Host
}

function Read-Text([string]$Prompt) {
    Write-Host ("{0}: " -f $Prompt) -NoNewline
    Read-Host
}

function Read-TextWithDefault([string]$Prompt, [string]$Default) {
    if ($Default) {
        Write-Host ("{0} [{1}]: " -f $Prompt, $Default) -NoNewline
    } else {
        Write-Host ("{0}: " -f $Prompt) -NoNewline
    }
    $resp = Read-Host
    if ([string]::IsNullOrWhiteSpace($resp)) {
        return $Default
    }
    return $resp
}

function Read-PasswordText([string]$Prompt) {
    Write-Host ("{0}: " -f $Prompt) -NoNewline
    $sec = Read-Host -AsSecureString
    if (-not $sec) { return "" }
    return [Runtime.InteropServices.Marshal]::PtrToStringUni(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    )
}

function Confirm([string]$Prompt, [bool]$DefaultNo = $true) {
    $suffix = if ($DefaultNo) { "[y/N]" } else { "[Y/n]" }
    Write-Host "$Prompt $suffix " -NoNewline
    $resp = Read-Host
    if ([string]::IsNullOrWhiteSpace($resp)) {
        return -not $DefaultNo
    }
    return $resp -match '^[Yy]'
}

function Pause() {
    Write-Host
    Write-Host "Press Enter to continue..." -NoNewline
    [void][System.Console]::ReadLine()
}

# ====================
# INTRO SCREEN
# ====================
function Show-Intro {
    Show-Banner
    Write-Host "This tool runs from your local Windows computer and connects to your Linux server"
    Write-Host "over SSH to help with:"
    Write-Host "  - New server setup"
    Write-Host "  - Creating non-root (sudo) users"
    Write-Host "  - Backing up and uploading .p12 files"
    Write-Host "  - Exporting SSH profiles"
    Write-Host "  - Exporting PuTTY keys (Windows only)"
    Write-Host
    Write-Host "Connection profiles are saved to:"
    Write-Host "  $Script:ProfilesDir\{ProfileName}_ssh_config.txt"
    Write-Host
    Write-Host "Press any key to continue to the menu..." -NoNewline
    [void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ====================
# MENU ENGINE
# ====================
function Show-MainMenu {
    $items = @(
        "New Server Setup"
        "Create Non-Root User"
        "Backup P12 File (From old server)"
        "Upload P12 File"
        "Export SSH-Config File"
        "Export to PuTTY Private Key"
        "Server Login (Launches in new window)"
        "Exit"
    )
    $index = 0
    while ($true) {
        Show-Banner
        for ($i = 0; $i -lt $items.Count; $i++) {
            if ($i -eq $index) {
                Write-Host ("> " + $items[$i]) -ForegroundColor Green
            } else {
                Write-Host ("  " + $items[$i])
            }
        }

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        switch ($key.VirtualKeyCode) {
            38 { if ($index -gt 0) { $index-- } }               # Up
            40 { if ($index -lt $items.Count - 1) { $index++ } } # Down
            13 { return $items[$index] }                        # Enter
        }
    }
}

# ====================
# PROFILE HANDLING
# ====================
function Test-ProfileNameValid([string]$Name) {
    return $Name -match '^[A-Za-z0-9-]+$'
}

function Get-ProfilePath([string]$ProfileName) {
    Join-Path $Script:ProfilesDir ("{0}_ssh_config.txt" -f $ProfileName)
}

function Save-SSHProfile($Profile) {
    $name = $Profile.Name
    if (-not (Test-ProfileNameValid $name)) {
        Write-Err "Profile name must be alphanumeric with dashes only."
        return
    }
    $path = Get-ProfilePath $name
    if (Test-Path $path) {
        if (-not (Confirm "Profile '$name' exists. Overwrite?")) {
            Write-Warn "Profile not saved."
            return
        }
    }

    $lines = @()
    $lines += "### This ssh_config file can also be used to import this server's settings into Termius. ###"
    $lines += ""
    $lines += "Host $($Profile.Name)"
    $lines += "    HostName $($Profile.Host)"
    $lines += "    User $($Profile.User)"
    $lines += "    Port $($Profile.Port)"
    if ($Profile.IdentityFile) {
        $lines += "    IdentityFile $($Profile.IdentityFile)"
    }
    $lines += ""
    $lines += "# CreatedOn: $(Get-Date -Format yyyy-MM-dd)"
    $lines += ("# SSH login: ssh{0} {1}@{2}" -f `
        ($(if ($Profile.IdentityFile) { " -i $($Profile.IdentityFile)" } else { "" })),
        $Profile.User, $Profile.Host)
    $lines += ("# SFTP: sftp{0} {1}@{2}" -f `
        ($(if ($Profile.IdentityFile) { " -i $($Profile.IdentityFile)" } else { "" })),
        $Profile.User, $Profile.Host)

    Set-Content -Path $path -Value $lines -Encoding UTF8
    Write-Ok "Profile saved to: $path"
}

function Get-StoredProfiles {
    $files = Get-ChildItem -Path $Script:ProfilesDir -Filter "*_ssh_config.txt" -ErrorAction SilentlyContinue
    if (-not $files) { return @() }
    $profiles = @()
    foreach ($f in $files) {
        $content = Get-Content $f -ErrorAction SilentlyContinue
        $hostLine = $content | Where-Object { $_ -match '^\s*Host\s+' } | Select-Object -First 1
        $hostNameLine = $content | Where-Object { $_ -match '^\s*HostName\s+' } | Select-Object -First 1
        $userLine = $content | Where-Object { $_ -match '^\s*User\s+' } | Select-Object -First 1
        $portLine = $content | Where-Object { $_ -match '^\s*Port\s+' } | Select-Object -First 1
        $identLine = $content | Where-Object { $_ -match '^\s*IdentityFile\s+' } | Select-Object -First 1

        if (-not $hostLine -or -not $hostNameLine -or -not $userLine) { continue }

        $p = [pscustomobject]@{
            Name         = ($hostLine -replace '^\s*Host\s+', '').Trim()
            Host         = ($hostNameLine -replace '^\s*HostName\s+', '').Trim()
            User         = ($userLine -replace '^\s*User\s+', '').Trim()
            Port         = if ($portLine) { ($portLine -replace '^\s*Port\s+', '').Trim() } else { "22" }
            IdentityFile = if ($identLine) { ($identLine -replace '^\s*IdentityFile\s+', '').Trim() } else { "" }
        }
        $profiles += $p
    }
    $profiles | Sort-Object Name
}

function Select-ProfileFromDisk([string]$Purpose) {
    $profiles = Get-StoredProfiles
    if (-not $profiles -or $profiles.Count -eq 0) {
        Write-Warn "No stored profiles found."
        return $null
    }

    $index = 0
    while ($true) {
        Show-Banner
        Write-Host "$Purpose - Select a stored connection profile:"
        Write-Host
        for ($i = 0; $i -lt $profiles.Count; $i++) {
            $line = "{0} ({1}@{2}:{3})" -f $profiles[$i].Name, $profiles[$i].User, $profiles[$i].Host, $profiles[$i].Port
            if ($i -eq $index) {
                Write-Host ("> " + $line) -ForegroundColor Green
            } else {
                Write-Host ("  " + $line)
            }
        }
        Write-Host
        Write-Host "Use Up/Down arrows and Enter. Press Esc to cancel."

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if ($key.VirtualKeyCode -eq 27) { return $null } # Esc

        switch ($key.VirtualKeyCode) {
            38 { if ($index -gt 0) { $index-- } }
            40 { if ($index -lt $profiles.Count - 1) { $index++ } }
            13 { return $profiles[$index] }
        }
    }
}

function Prompt-Connection([string]$Purpose, [ref]$CachedConn) {
    if ($CachedConn.Value) {
        $c = $CachedConn.Value
        if (Confirm "Reuse $Purpose connection: $($c.User)@$($c.Host):$($c.Port)?") {
            return $c
        }
    }

    $useStored = $false
    if (Get-StoredProfiles) {
        $useStored = Confirm "Load a stored connection profile for $Purpose?"
    }

    if ($useStored) {
        $p = Select-ProfileFromDisk "$Purpose"
        if ($p) {
            $CachedConn.Value = $p
            return $p
        }
    }

    Show-Banner
    Write-Host "$Purpose - manual connection entry"
    $serverHost = Read-Text "Server IP or Hostname"
    $user       = Read-Text "SSH Username"
    $port       = Read-Text "SSH Port (default 22)"
    if (-not $port) { $port = "22" }

    $identPrompt = "Path to SSH private key (blank for password auth)"
    $ident = Read-TextWithDefault $identPrompt $Script:LastIdentityPath

    $conn = [pscustomobject]@{
        Name         = ""
        Host         = $serverHost
        User         = $user
        Port         = $port
        IdentityFile = $ident
    }

    if ($ident) { $Script:LastIdentityPath = $ident }

    if (Confirm "Save this connection as a profile?") {
        $pname = Read-Text "Profile name (letters, numbers, dashes only)"
        if (-not (Test-ProfileNameValid $pname)) {
            Write-Err "Invalid profile name. Not saving."
        } else {
            $conn.Name = $pname
            Save-SSHProfile $conn
        }
    }

    $CachedConn.Value = $conn
    return $conn
}

# ====================
# SSH HELPERS
# ====================
function Build-SshBaseArgs($Conn) {
    $args = @("-p", $Conn.Port)
    if ($Conn.IdentityFile) {
        $args += @("-i", $Conn.IdentityFile)
    }
    return $args
}

function Invoke-OpenSSHWithKnownHostsFix {
    param(
        [Parameter(Mandatory=$true)][string]$Exe,    # "ssh" or "scp"
        [Parameter(Mandatory=$true)][string[]]$Args,
        [Parameter(Mandatory=$true)][string]$HostName,
        [int]$MaxRetries = 1
    )

    $retry = 0
    while ($true) {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Exe
        $psi.Arguments = [string]::Join(" ", $Args)
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $p = [System.Diagnostics.Process]::Start($psi)
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()

        if ($stdout) { Write-Host $stdout }
        if ($stderr) { Write-Host $stderr }

        if ($p.ExitCode -eq 0) {
            return 0
        }

        if ($stderr -like "*REMOTE HOST IDENTIFICATION HAS CHANGED!*" -and $retry -lt $MaxRetries) {
            Write-Warn "Host key mismatch detected for $HostName. Cleaning known_hosts and retrying..."
            & ssh-keygen -R $HostName | Out-Null
            $retry++
            continue
        }

        return $p.ExitCode
    }
}

function Invoke-SshCommand($Conn, [string]$RemoteCmd) {
    $args = Build-SshBaseArgs $Conn
    $args += ("{0}@{1}" -f $Conn.User, $Conn.Host)
    if ($RemoteCmd) { $args += $RemoteCmd }

    [void](Invoke-OpenSSHWithKnownHostsFix -Exe "ssh" -Args $args -HostName $Conn.Host)
}

function Invoke-ScpDownload($Conn, [string]$RemotePath, [string]$LocalPath) {
    $args = Build-SshBaseArgs $Conn
    $source = "{0}@{1}:{2}" -f $Conn.User, $Conn.Host, $RemotePath
    $args += @($source, $LocalPath)

    [void](Invoke-OpenSSHWithKnownHostsFix -Exe "scp" -Args $args -HostName $Conn.Host)
}

function Invoke-ScpUpload($Conn, [string]$LocalPath, [string]$RemotePath) {
    $args = Build-SshBaseArgs $Conn
    $dest = "{0}@{1}:{2}" -f $Conn.User, $Conn.Host, $RemotePath
    $args += @($LocalPath, $dest)

    [void](Invoke-OpenSSHWithKnownHostsFix -Exe "scp" -Args $args -HostName $Conn.Host)
}

# ====================
# P12 HELPERS
# ====================
function Select-LocalP12File {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "P12 Files (*.p12)|*.p12|All Files (*.*)|*.*"
    $ofd.Title  = "Select a .p12 file"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $ofd.FileName
    }
    return $null
}

function Test-P12Password([string]$FilePath, [string]$Password) {
    $openssl = Get-Command "openssl.exe" -ErrorAction SilentlyContinue
    if (-not $openssl) {
        Write-Warn "openssl not found; skipping P12 password verification."
        return $true
    }
    & $openssl.Source "pkcs12" "-in" $FilePath "-nokeys" "-passin" "pass:$Password" "-passout" "pass:dummy" 1>$null 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Get-P12Alias([string]$FilePath, [string]$Password) {
    $openssl = Get-Command "openssl.exe" -ErrorAction SilentlyContinue
    if (-not $openssl) {
        return $null
    }

    $info = & $openssl.Source "pkcs12" "-in" $FilePath "-nokeys" "-info" "-passin" "pass:$Password" 2>&1
    if ($LASTEXITCODE -ne 0 -or -not $info) {
        $info = & $openssl.Source "pkcs12" "-legacy" "-in" $FilePath "-nokeys" "-info" "-passin" "pass:$Password" 2>&1
    }

    if (-not $info) { return $null }

    foreach ($line in $info) {
        if ($line -match "friendlyName\s*:\s*(.+)$") {
            return $Matches[1].Trim()
        }
    }
    return $null
}

# ====================
# FEATURES
# ====================

# 1) EXPORT SSH CONFIG MANUALLY
function Run-ExportSSHProfile {
    Show-Banner
    Write-Host "Export SSH-Config File"
    $name = Read-Text "Profile name (letters, numbers, dashes only)"
    if (-not (Test-ProfileNameValid $name)) {
        Write-Err "Invalid profile name."
        Pause
        return
    }
    $serverHost = Read-Text "Server IP or Hostname"
    $user       = Read-Text "SSH Username"
    $port       = Read-Text "SSH Port (default 22)"
    if (-not $port) { $port = "22" }
    $ident      = Read-Text "Path to SSH private key (blank for password auth)"

    $profile = [pscustomobject]@{
        Name         = $name
        Host         = $serverHost
        User         = $user
        Port         = $port
        IdentityFile = $ident
    }

    Save-SSHProfile $profile
    Pause
}

# 2) BACKUP P12 FROM OLD SERVER
function Run-BackupP12 {
    $conn = Prompt-Connection "Old Server (P12 backup source)" ([ref]$Script:OldServerConn)
    if (-not $conn) { Pause; return }

    Show-Banner
    Write-Host "Scanning for .p12 files on old server (up to 5 levels deep)..."
    Write-Host

    $remoteCmd = "cd ~; find . -maxdepth 5 -type f -name '*.p12' ! -path '*\/hash\/*' ! -path '*\/ordinal\/*'"
    $paths = Invoke-SshCommand $conn $remoteCmd 2>$null

    if (-not $paths) {
        Write-Warn "No .p12 files found on remote server."
        Pause
        return
    }

    $list = $paths -split "`n" | Where-Object { $_ -and ($_ -ne ".") }
    $list = $list | Sort-Object
    $list = $list | ForEach-Object { $_.Trim() }

    if (-not $list -or $list.Count -eq 0) {
        Write-Warn "No .p12 files found."
        Pause
        return
    }

    $index = 0
    while ($true) {
        Show-Banner
        Write-Host "Select a .p12 file to download from old server:"
        Write-Host
        for ($i = 0; $i -lt $list.Count; $i++) {
            $line = $list[$i]
            if ($i -eq $index) {
                Write-Host ("> " + $line) -ForegroundColor Green
            } else {
                Write-Host ("  " + $line)
            }
        }
        Write-Host
        Write-Host "Use Up/Down and Enter. Esc to cancel."

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        if ($key.VirtualKeyCode -eq 27) { return }

        switch ($key.VirtualKeyCode) {
            38 { if ($index -gt 0) { $index-- } }
            40 { if ($index -lt $list.Count - 1) { $index++ } }
            13 {
                $chosen = $list[$index]
                $localDir = Join-Path $HOME "Downloads"
                if (-not (Test-Path $localDir)) {
                    New-Item -ItemType Directory -Path $localDir | Out-Null
                }
                $localPath = Join-Path $localDir ([IO.Path]::GetFileName($chosen))
                Invoke-ScpDownload $conn $chosen $localPath
                Write-Ok "P12 file downloaded to: $localPath"
                Pause
                return
            }
        }
    }
}

# 3) UPLOAD P12 TO NEW SERVER
function Run-UploadP12 {
    $file = $null

    # Offer to reuse last P12 path if known and exists
    if ($Script:LastP12Path -and (Test-Path $Script:LastP12Path)) {
        if (Confirm ("Reuse last selected .p12 file? `"$($Script:LastP12Path)`"")) {
            $file = $Script:LastP12Path
        }
    }

    if (-not $file) {
        $file = Select-LocalP12File
    }

    if (-not $file) {
        Write-Warn "No file selected."
        Pause
        return
    }
    Write-Ok "Selected: $file"
    $Script:LastP12Path = $file

    $maxAttempts = 5
    $pw = $null
    $alias = $null
    $ok = $false

    for ($i = 1; $i -le $maxAttempts; $i++) {
        $pw = Read-PasswordText "Enter .p12 password (attempt $i of $maxAttempts)"
        if (-not $pw) {
            Write-Warn "Empty password not allowed."
            continue
        }

        if (Test-P12Password $file $pw) {
            Write-Ok "P12 password verified locally."
            $alias = Get-P12Alias $file $pw
            if ($alias) {
                Write-Ok "P12 alias (friendlyName): $alias"
                Write-Warn "Make sure to write this alias down and document it."
            } else {
                Write-Warn "No friendlyName/alias found in this P12."
            }
            $ok = $true
            break
        } else {
            Write-Err "Incorrect P12 password."
        }
    }

    if (-not $ok) {
        Write-Err "Too many incorrect attempts. Aborting upload."
        Pause
        return
    }

    $conn = Prompt-Connection "New Server (P12 upload target)" ([ref]$Script:NewServerConn)
    if (-not $conn) { Pause; return }

    $remoteDest = "~/"
    Invoke-ScpUpload $conn $file $remoteDest
    Write-Ok "P12 file uploaded to home directory on remote server (~/)."
    Write-Host "Remote path will resolve to:"
    Write-Host "  /root           if logged in as root"
    Write-Host "  /home/<user>    if logged in as a normal user"

    if ($alias) {
        Write-Host
        Write-Info "Reminder: alias for this P12 is: $alias"
    }

    Pause
}

# 4) NEW SERVER SETUP (REMOTE SCAFFOLD)
function Run-NewServerSetup {
    $conn = Prompt-Connection "New Server (initial root or admin login)" ([ref]$Script:NewServerConn)
    if (-not $conn) { Pause; return }

    Show-Banner
    Write-Host "New Server Setup"
    Write-Host
    Write-Host "This will:"
    Write-Host "  - Create a new non-root sudo user on the remote Linux server"
    Write-Host "  - Copy SSH authorized_keys from root or SUDO_USER if available"
    Write-Host "  - Test SSH login as the new user"
    Write-Host "  - Optionally harden sshd to disable root SSH and lock root password"
    Write-Host
    if (-not (Confirm "Continue with New Server Setup?")) {
        return
    }

    # 1) Ask for new user credentials
    $newUser = Read-Text "Enter new username (e.g. nodeadmin)"
    if (-not $newUser) {
        Write-Err "Username cannot be empty."
        Pause
        return
    }

    $pw1 = Read-PasswordText "Enter password for $newUser"
    $pw2 = Read-PasswordText "Confirm password"
    if ($pw1 -ne $pw2) {
        Write-Err "Passwords do not match."
        Pause
        return
    }

    # 2) Remote script: create user + add to sudo/wheel + copy authorized_keys
    Write-Info "Creating non-root user '$newUser' on remote server..."
    $remoteUserScript = @"
set -e
NEWUSER='$newUser'

if id "\$NEWUSER" >/dev/null 2>&1; then
  echo "User \$NEWUSER already exists. Skipping creation."
  exit 0
fi

if command -v adduser >/dev/null 2>&1; then
  adduser --disabled-password --gecos "" "\$NEWUSER"
else
  useradd -m -s /bin/bash "\$NEWUSER"
fi

echo "`$NEWUSER:$pw1" | chpasswd

if getent group sudo >/dev/null 2>&1; then
  usermod -aG sudo "\$NEWUSER"
elif getent group wheel >/dev/null 2>&1; then
  usermod -aG wheel "\$NEWUSER"
fi

SRC_KEYS=""
if [ -f "/root/.ssh/authorized_keys" ]; then
  SRC_KEYS="/root/.ssh/authorized_keys"
elif [ -n "\$SUDO_USER" ] && [ -f "/home/\$SUDO_USER/.ssh/authorized_keys" ]; then
  SRC_KEYS="/home/\$SUDO_USER/.ssh/authorized_keys"
fi

if [ -n "\$SRC_KEYS" ]; then
  HOME_DIR=\$(eval echo "~\$NEWUSER")
  mkdir -p "\$HOME_DIR/.ssh"
  cp "\$SRC_KEYS" "\$HOME_DIR/.ssh/authorized_keys"
  chown -R "`$NEWUSER:`$NEWUSER" "$HOME_DIR/.ssh"
  chmod 700 "\$HOME_DIR/.ssh"
  chmod 600 "\$HOME_DIR/.ssh/authorized_keys"
fi

echo "User \$NEWUSER created and configured."
"@

    $args = Build-SshBaseArgs $conn
    $args += ("{0}@{1}" -f $conn.User, $conn.Host)
    $args += "bash -s"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "ssh"
    $psi.Arguments = [string]::Join(" ", $args)
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $false

    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.WriteLine($remoteUserScript)
    $proc.StandardInput.Close()
    $out = $proc.StandardOutput.ReadToEnd()
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($out) { Write-Host $out }
    if ($err) { Write-Host $err }
    if ($proc.ExitCode -ne 0) {
        Write-Err "Remote user creation script failed (exit $($proc.ExitCode))."
        Pause
        return
    }

    Write-Ok "Remote user '$newUser' created (or already present)."

    # 3) Test SSH login as new user (using same IdentityFile and port)
    Write-Info "Testing SSH login as $newUser..."
    $testConn = [pscustomobject]@{
        Host         = $conn.Host
        User         = $newUser
        Port         = $conn.Port
        IdentityFile = $conn.IdentityFile
    }

    $testArgs = Build-SshBaseArgs $testConn
    $testArgs += ("{0}@{1}" -f $testConn.User, $testConn.Host)
    $testArgs += "echo ok"

    $testExit = Invoke-OpenSSHWithKnownHostsFix -Exe "ssh" -Args $testArgs -HostName $testConn.Host -MaxRetries 1
    if ($testExit -eq 0) {
        Write-Ok "SSH login as '$newUser' succeeded."
        $newLoginWorks = $true
    } else {
        Write-Warn "SSH login as '$newUser' failed."
        $newLoginWorks = $false
    }

    # 4) Optionally harden sshd
    if (-not $newLoginWorks) {
        Write-Warn "New user login did not succeed. Root SSH hardening is NOT recommended."
        if (-not (Confirm "Proceed to harden sshd anyway (NOT recommended)?")) {
            Pause
            return
        }
    } else {
        if (-not (Confirm "Disable root SSH login and lock root password now?")) {
            Pause
            return
        }
    }

    Write-Info "Hardening sshd on remote server..."
    $remoteHardenScript = @"
set -e

SSHD_CFG="/etc/ssh/sshd_config"
BACKUP="/etc/ssh/sshd_config.$(date +%Y%m%d-%H%M%S).bak"

cp -a "\$SSHD_CFG" "\$BACKUP"

# Comment existing PermitRootLogin lines
sed -i 's/^[[:space:]]*PermitRootLogin[[:space:]].*/# &/I' "\$SSHD_CFG"

# In any sshd_config.d files, comment PermitRootLogin yes
if ls /etc/ssh/sshd_config.d/*.conf >/dev/null 2>&1; then
  sed -i 's/^[[:space:]]*PermitRootLogin[[:space:]]*yes/# &/I' /etc/ssh/sshd_config.d/*.conf || true
fi

# Remove any existing 'Match User root' block at end (simple approach)
if grep -qi '^[[:space:]]*Match[[:space:]]\+User[[:space:]]\+root' "\$SSHD_CFG"; then
  # Delete from Match User root to EOF and re-add below
  awk '
  BEGIN{del=0}
  /^Match[[:space:]]+User[[:space:]]+root/{del=1}
  !del{print}
  ' "\$SSHD_CFG" > "\$SSHD_CFG.tmp"
  mv "\$SSHD_CFG.tmp" "\$SSHD_CFG"
fi

{
  echo ""
  echo "PermitRootLogin no"
  echo ""
  echo "Match User root"
  echo "  PasswordAuthentication no"
  echo "  PermitRootLogin no"
} >> "\$SSHD_CFG"

if ! sshd -t 2>/tmp/sshd_test.err; then
  echo "sshd config test FAILED. Restoring backup."
  cat /tmp/sshd_test.err
  mv -f "\$BACKUP" "\$SSHD_CFG"
  exit 1
fi

if command -v systemctl >/dev/null 2>&1; then
  systemctl restart ssh || systemctl restart sshd
else
  service ssh restart 2>/dev/null || service sshd restart 2>/dev/null
fi

passwd -l root >/dev/null 2>&1 || true
if [ -f /root/.ssh/authorized_keys ]; then
  mv /root/.ssh/authorized_keys /root/.ssh/authorized_keys.disabled 2>/dev/null || true
fi

echo "Root SSH login disabled, root password locked, sshd restarted."
"@

    $hArgs = Build-SshBaseArgs $conn
    $hArgs += ("{0}@{1}" -f $conn.User, $conn.Host)
    $hArgs += "bash -s"

    $hPsi = New-Object System.Diagnostics.ProcessStartInfo
    $hPsi.FileName = "ssh"
    $hPsi.Arguments = [string]::Join(" ", $hArgs)
    $hPsi.RedirectStandardInput = $true
    $hPsi.RedirectStandardOutput = $true
    $hPsi.RedirectStandardError  = $true
    $hPsi.UseShellExecute = $false
    $hPsi.CreateNoWindow = $false

    $hProc = [System.Diagnostics.Process]::Start($hPsi)
    $hProc.StandardInput.WriteLine($remoteHardenScript)
    $hProc.StandardInput.Close()
    $hOut = $hProc.StandardOutput.ReadToEnd()
    $hErr = $hProc.StandardError.ReadToEnd()
    $hProc.WaitForExit()

    if ($hOut) { Write-Host $hOut }
    if ($hErr) { Write-Host $hErr }
    if ($hProc.ExitCode -ne 0) {
        Write-Err "Remote sshd hardening script failed (exit $($hProc.ExitCode))."
        Pause
        return
    }

    Write-Ok "Root SSH login disabled and root password locked on remote server."

    # 5) Export profile for the new user and offer to launch
    if (Confirm "Export SSH connection profile for new user '$newUser'?") {
        $serverHost = $conn.Host
        $port       = $conn.Port
        $ident      = $conn.IdentityFile
        $profileName = Read-Text "Profile name for this new user (default: $newUser)"
        if (-not $profileName) { $profileName = $newUser }

        $profile = [pscustomobject]@{
            Name         = $profileName
            Host         = $serverHost
            User         = $newUser
            Port         = $port
            IdentityFile = $ident
        }

        Save-SSHProfile $profile

        if (Confirm "Launch SSH for this new profile in a new window now?") {
            if (Get-Command wt.exe -ErrorAction SilentlyContinue) {
                $argsList = @("new-window", "ssh")
                if ($ident) { $argsList += @("-i", $ident) }
                $argsList += @("-p", $port, ("{0}@{1}" -f $newUser, $serverHost))
                Start-Process wt.exe -ArgumentList $argsList
            } else {
                $cmdPieces = @("ssh")
                if ($ident) { $cmdPieces += @("-i", $ident) }
                $cmdPieces += @("-p", $port, ("{0}@{1}" -f $newUser, $serverHost))
                $cmdString = $cmdPieces -join " "
                Start-Process "powershell.exe" -ArgumentList "-NoExit", "-Command", $cmdString
            }
            Write-Ok "SSH session launched for $newUser."
        }
    }
    Pause
}

# 5) CREATE NON-ROOT USER (REMOTE)
function Run-CreateNonRootUser {
    $conn = Prompt-Connection "Server (for non-root user creation)" ([ref]$Script:NewServerConn)
    if (-not $conn) { Pause; return }

    Show-Banner
    Write-Host "Remote Non-Root User Creation"
    Write-Host
    $newUser = Read-Text "Enter new username (e.g. nodeadmin)"
    if (-not $newUser) {
        Write-Err "Username cannot be empty."
        Pause
        return
    }
    $pw1 = Read-PasswordText "Enter password for $newUser"
    $pw2 = Read-PasswordText "Confirm password"
    if ($pw1 -ne $pw2) {
        Write-Err "Passwords do not match."
        Pause
        return
    }

    Write-Host
    Write-Host "The following commands will run on the remote server:"
    Write-Host "  - create user"
    Write-Host "  - add to sudo or wheel group (if exists)"
    Write-Host "  - copy authorized_keys from root or SUDO_USER if available"
    Write-Host
    if (-not (Confirm "Proceed with remote user creation?")) { return }

    $remoteUserScript = @"
set -e
NEWUSER='$newUser'

if id "\$NEWUSER" >/dev/null 2>&1; then
  echo "User \$NEWUSER already exists. Skipping creation."
  exit 0
fi

if command -v adduser >/dev/null 2>&1; then
  adduser --disabled-password --gecos "" "\$NEWUSER"
else
  useradd -m -s /bin/bash "\$NEWUSER"
fi

echo "`$NEWUSER:$pw1" | chpasswd

if getent group sudo >/dev/null 2>&1; then
  usermod -aG sudo "\$NEWUSER"
elif getent group wheel >/dev/null 2>&1; then
  usermod -aG wheel "\$NEWUSER"
fi

SRC_KEYS=""
if [ -f "/root/.ssh/authorized_keys" ]; then
  SRC_KEYS="/root/.ssh/authorized_keys"
elif [ -n "\$SUDO_USER" ] && [ -f "/home/\$SUDO_USER/.ssh/authorized_keys" ]; then
  SRC_KEYS="/home/\$SUDO_USER/.ssh/authorized_keys"
fi

if [ -n "\$SRC_KEYS" ]; then
  HOME_DIR=\$(eval echo "~\$NEWUSER")
  mkdir -p "\$HOME_DIR/.ssh"
  cp "\$SRC_KEYS" "\$HOME_DIR/.ssh/authorized_keys"
  chown -R "`$NEWUSER:`$NEWUSER" "$HOME_DIR/.ssh"
  chmod 700 "\$HOME_DIR/.ssh"
  chmod 600 "\$HOME_DIR/.ssh/authorized_keys"
fi

echo "User \$NEWUSER created and configured."
"@

    $args = Build-SshBaseArgs $conn
    $args += ("{0}@{1}" -f $conn.User, $conn.Host)
    $args += "bash -s"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "ssh"
    $psi.Arguments = [string]::Join(" ", $args)
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $false

    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.WriteLine($remoteUserScript)
    $proc.StandardInput.Close()
    $out = $proc.StandardOutput.ReadToEnd()
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($out) { Write-Host $out }
    if ($err) { Write-Host $err }

    if ($proc.ExitCode -ne 0) {
        Write-Err "Remote user creation script failed (exit $($proc.ExitCode))."
    } else {
        Write-Ok "Remote user creation script executed."
    }
    Pause
}

# 6) EXPORT TO PUTTY PRIVATE KEY
function Run-ExportPuTTYKey {
    Show-Banner
    Write-Host "Export to PuTTY Private Key (.ppk)"
    $source = Read-Text "Path to OpenSSH private key (e.g. id_ed25519)"
    if (-not (Test-Path $source)) {
        Write-Err "File not found: $source"
        Pause
        return
    }
    $defaultOut = [IO.Path]::ChangeExtension($source, ".ppk")
    $dest = Read-Text "Output PuTTY key path (default: $defaultOut)"
    if (-not $dest) { $dest = $defaultOut }

    $puttygen = Get-Command "puttygen.exe" -ErrorAction SilentlyContinue
    if (-not $puttygen) {
        Write-Err "puttygen.exe not found in PATH. Install PuTTY or PuTTYgen."
        Pause
        return
    }

    & $puttygen.Source $source "-O" "private" "-o" $dest
    if ($LASTEXITCODE -eq 0) {
        Write-Ok "PuTTY key saved to: $dest"
    } else {
        Write-Err "PuTTY key export failed."
    }
    Pause
}

# 7) SERVER LOGIN (NEW WINDOW)
function Run-ServerLogin {
    $conn = Prompt-Connection "Server Login" ([ref]$Script:NewServerConn)
    if (-not $conn) { Pause; return }

    if (Confirm "Launch SSH in new Windows Terminal window now?") {
        if (Get-Command wt.exe -ErrorAction SilentlyContinue) {
            $args = @("new-window", "ssh")
            if ($conn.IdentityFile) {
                $args += @("-i", $conn.IdentityFile)
            }
            $args += @("-p", $conn.Port, ("{0}@{1}" -f $conn.User, $conn.Host))
            Start-Process wt.exe -ArgumentList $args
        } else {
            $cmdPieces = @("ssh")
            if ($conn.IdentityFile) { $cmdPieces += @("-i", $conn.IdentityFile) }
            $cmdPieces += @("-p", $conn.Port, ("{0}@{1}" -f $conn.User, $conn.Host))
            $cmdString = $cmdPieces -join " "
            Start-Process "powershell.exe" -ArgumentList "-NoExit", "-Command", $cmdString
        }
        Write-Ok "SSH session launched in a new window."
    }
    Pause
}

# ====================
# MAIN LOOP
# ====================
Show-Intro

while ($true) {
    $choice = Show-MainMenu
    switch ($choice) {
        "New Server Setup" {
            Run-NewServerSetup
        }
        "Create Non-Root User" {
            Run-CreateNonRootUser
        }
        "Backup P12 File (From old server)" {
            Run-BackupP12
        }
        "Upload P12 File" {
            Run-UploadP12
        }
        "Export SSH-Config File" {
            Run-ExportSSHProfile
        }
        "Export to PuTTY Private Key" {
            Run-ExportPuTTYKey
        }
        "Server Login (Launches in new window)" {
            Run-ServerLogin
        }
        "Exit" {
            break
        }
    }
}
