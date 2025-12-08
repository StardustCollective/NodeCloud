Add-Type -AssemblyName System.Windows.Forms

function Test-IsSshPrivateKey {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    $lines = Get-Content -LiteralPath $Path -TotalCount 5
    foreach ($line in $lines) {
        if ($line -match '^(-----BEGIN [A-Z0-9 ]*PRIVATE KEY-----|PuTTY-User-Key-File-|\s*ssh-|sk-ssh-)') {
            return $true
        }
    }
    return $false
}

function Select-SshKeyFile {
    param(
        [string]$PromptTitle = "Select your SSH private key"
    )

    $sshDialog = New-Object System.Windows.Forms.OpenFileDialog
    $sshDialog.Multiselect = $false
    $sshDialog.Title = $PromptTitle
    $sshDialog.Filter = "All Files (*.*)|*.*"

    $sshDir = Join-Path $HOME ".ssh"
    if (Test-Path -LiteralPath $sshDir) {
        $sshDialog.InitialDirectory = $sshDir
    }

    while ($true) {
        $result = $sshDialog.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            return $null
        }

        $candidate = $sshDialog.FileName
        if (Test-IsSshPrivateKey -Path $candidate) {
            return $candidate
        } else {
            Write-Host "The selected file does not appear to be an SSH private key. Please choose a valid key file." -ForegroundColor Yellow
        }
    }
}

function Ensure-SshTools {
    $ssh = Get-Command ssh -ErrorAction SilentlyContinue
    $sshKeygen = Get-Command ssh-keygen -ErrorAction SilentlyContinue

    if (-not $ssh -or -not $sshKeygen) {
        Write-Host "ssh and/or ssh-keygen commands were not found on this system." -ForegroundColor Red
        Write-Host "Please install the OpenSSH Client feature in Windows (Settings → Apps → Optional Features)." -ForegroundColor Red
        return $false
    }
    return $true
}

function Get-KnownHostsPath {
    $userProfile = $env:USERPROFILE
    if (-not $userProfile) {
        $userProfile = $HOME
    }

    $sshDir = Join-Path $userProfile ".ssh"
    if (-not (Test-Path -LiteralPath $sshDir)) {
        New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
    }

    return (Join-Path $sshDir "known_hosts")
}

function Clean-HostFromKnownHosts {
    param(
        [string]$Server
    )

    if (Get-Command ssh-keygen -ErrorAction SilentlyContinue) {
        Write-Host "Cleaning host key for $Server from known_hosts..." -ForegroundColor Yellow
        ssh-keygen -R $Server | Out-Null
    }
}

function Invoke-ScpWithHostHandling {
    param(
        [string]$Username,
        [string]$Server,
        [string]$PrivateKeyPath,
        [string]$LocalPath
    )

    $knownHosts = Get-KnownHostsPath
    $remoteTarget = "$Username@$Server`:~/"

    $args = @()

    if ($PrivateKeyPath) {
        $args += "-i"
        $args += $PrivateKeyPath
    }

    $args += @(
        "-o", "StrictHostKeyChecking=no",
        "-o", "UserKnownHostsFile=$knownHosts",
        $LocalPath,
        $remoteTarget
    )

    $attempt = 1
    while ($attempt -le 2) {
        Write-Host "Running scp (attempt $attempt)..." -ForegroundColor Cyan
        $scpOutput = & scp @args 2>&1
        $code = $LASTEXITCODE

        if ($code -eq 0) {
            return 0
        }

        if ($attempt -eq 1 -and ($scpOutput -match "REMOTE HOST IDENTIFICATION HAS CHANGED" -or $scpOutput -match "Offending .*known_hosts")) {
            Write-Host "Host key mismatch detected for $Server." -ForegroundColor Yellow
            Clean-HostFromKnownHosts -Server $Server
            $attempt++
            continue
        }

        Write-Host $scpOutput -ForegroundColor DarkYellow
        return $code
    }

    return $code
}

function Generate-NewSshKeyPair {
    param(
        [string]$Username,
        [string]$Server
    )

    $sshDir = Join-Path $HOME ".ssh"
    if (-not (Test-Path -LiteralPath $sshDir)) {
        New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
    }

    $baseName = "nodecloud_${Username}_$Server_ed25519"
    $keyPath = Join-Path $sshDir $baseName
    $counter = 0
    while (Test-Path -LiteralPath $keyPath) {
        $counter++
        $keyPath = Join-Path $sshDir ("{0}_{1}" -f $baseName, $counter)
    }

    Write-Host ""
    Write-Host "Generating a new SSH key pair:" -ForegroundColor Cyan
    Write-Host "  Private key: $keyPath"
    Write-Host "  Public key:  $keyPath.pub"
    Write-Host ""

    $usePass = Read-Host "Do you want to set a passphrase on the SSH key? [y/N]"
    $passphrase = ""
    if ($usePass -match '^[Yy]') {
        $secure = Read-Host "Enter passphrase (will not be shown)" -AsSecureString
        $passphrase = [System.Runtime.InteropServices.Marshal]::PtrToStringUni(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        )
    }

    $comment = "nodecloud-$Username@$Server"
    $args = @("-t", "ed25519", "-f", $keyPath, "-C", $comment)

    if ($passphrase -ne "") {
        Write-Host "ssh-keygen will now ask you to confirm the passphrase." -ForegroundColor Yellow
    } else {
        $args += @("-N", "")
    }

    & ssh-keygen @args

    if ($LASTEXITCODE -ne 0) {
        Write-Host "ssh-keygen failed. Aborting key setup." -ForegroundColor Red
        return $null
    }

    return $keyPath
}

function Import-ExistingSshKey {
    Write-Host ""
    Write-Host "Select an existing SSH private key to import (it will NOT be modified)." -ForegroundColor Cyan
    $keyPath = Select-SshKeyFile -PromptTitle "Select an existing SSH private key to import"
    if (-not $keyPath) {
        Write-Host "No SSH key selected. Aborting key import." -ForegroundColor Yellow
        return $null
    }

    $hasPass = Read-Host "Does this SSH key have a passphrase? [y/N]"
    if ($hasPass -match '^[Yy]') {
        Write-Host "Reminder: This script does not verify your SSH key passphrase directly for security reasons." -ForegroundColor Yellow
        Write-Host "It can use ssh itself to test the key against the server (ssh will prompt you for the passphrase)." -ForegroundColor Yellow
        Write-Host ""
    }

    return $keyPath
}

function Install-PubKeyToServer {
    param(
        [string]$Username,
        [string]$Server,
        [string]$PrivateKeyPath
    )

    if (-not (Ensure-SshTools)) {
        return $false
    }

    $tempPub = [System.IO.Path]::GetTempFileName()
    $pubPath = "$tempPub.pub"
    Move-Item -LiteralPath $tempPub -Destination $pubPath -Force

    Write-Host ""
    Write-Host "Deriving public key from private key..." -ForegroundColor Cyan
    & ssh-keygen -y -f $PrivateKeyPath > $pubPath

    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to derive public key from private key. Aborting key installation." -ForegroundColor Red
        Remove-Item -LiteralPath $pubPath -ErrorAction SilentlyContinue
        return $false
    }

    Write-Host "Installing public key into ~/.ssh/authorized_keys on the server..." -ForegroundColor Cyan

    $remoteUserHost = "$Username@$Server"
    $remoteCmd = "mkdir -p ~/.ssh && chmod 700 ~/.ssh && cat >> ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys"
    $knownHosts = Get-KnownHostsPath

    try {
        $pubContent = Get-Content -LiteralPath $pubPath -Raw
        $sshArgs = @(
            "-o", "StrictHostKeyChecking=no",
            "-o", "UserKnownHostsFile=$knownHosts",
            $remoteUserHost,
            $remoteCmd
        )

        $sshOutput = $pubContent | & ssh @sshArgs 2>&1
        $code = $LASTEXITCODE

        if ($code -ne 0) {
            if ($sshOutput -match "REMOTE HOST IDENTIFICATION HAS CHANGED" -or $sshOutput -match "Offending .*known_hosts") {
                Write-Host "Host key mismatch detected while installing pubkey. Cleaning and retrying..." -ForegroundColor Yellow
                Clean-HostFromKnownHosts -Server $Server
                $sshOutput = $pubContent | & ssh @sshArgs 2>&1
                $code = $LASTEXITCODE
            }
        }

        if ($code -ne 0) {
            Write-Host "ssh command to install authorized key failed." -ForegroundColor Red
            Write-Host $sshOutput -ForegroundColor DarkYellow
            Remove-Item -LiteralPath $pubPath -ErrorAction SilentlyContinue
            return $false
        }
    } catch {
        Write-Host "Error running ssh to install authorized key: $_" -ForegroundColor Red
        Remove-Item -LiteralPath $pubPath -ErrorAction SilentlyContinue
        return $false
    }

    Remove-Item -LiteralPath $pubPath -ErrorAction SilentlyContinue
    Write-Host "SSH public key successfully installed on the server." -ForegroundColor Green
    return $true
}

function Test-P12Password {
    param(
        [string]$Path
    )

    for ($i = 1; $i -le 12; $i++) {
        $prompt = "Enter the password for this .p12 file (attempt $i of 12): "
        $secure = Read-Host $prompt -AsSecureString

        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

        try {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($Path, $plain, 'Exportable')
            Write-Host "P12 password verified successfully." -ForegroundColor Green
            return $secure
        } catch {
            Write-Host "The password did not work for this .p12 file." -ForegroundColor Yellow
        }
    }

    Write-Host "Too many failed attempts. Aborting." -ForegroundColor Red
    return $null
}

function Test-SshKeyAgainstServer {
    param(
        [string]$Username,
        [string]$Server,
        [string]$PrivateKeyPath
    )

    if (-not (Ensure-SshTools)) {
        return $false
    }

    $knownHosts = Get-KnownHostsPath
    $remoteUserHost = "$Username@$Server"
    $args = @(
        "-i", $PrivateKeyPath,
        "-o", "StrictHostKeyChecking=no",
        "-o", "UserKnownHostsFile=$knownHosts",
        $remoteUserHost,
        "echo 'SSH key test ok'"
    )

    $attempt = 1
    while ($attempt -le 2) {
        Write-Host ""
        Write-Host "Testing SSH key against the server (attempt $attempt)..." -ForegroundColor Cyan
        Write-Host "You may be prompted for the SSH key passphrase by ssh itself." -ForegroundColor Yellow
        Write-Host ""

        $output = & ssh @args 2>&1
        $code = $LASTEXITCODE

        if ($code -eq 0) {
            Write-Host "SSH key successfully authenticated with the server." -ForegroundColor Green
            return $true
        }

        if ($attempt -eq 1 -and ($output -match "REMOTE HOST IDENTIFICATION HAS CHANGED" -or $output -match "Offending .*known_hosts")) {
            Write-Host "Host key mismatch detected for $Server during key test. Cleaning and retrying..." -ForegroundColor Yellow
            Clean-HostFromKnownHosts -Server $Server
            $attempt++
            continue
        }

        Write-Host "SSH key authentication test failed (exit code $code)." -ForegroundColor Red
        Write-Host "Possible reasons: wrong passphrase, wrong user, or key not authorized." -ForegroundColor Yellow
        Write-Host $output -ForegroundColor DarkYellow
        return $false
    }

    return $false
}

Write-Host "This script will:" -ForegroundColor Cyan
Write-Host "  1) Let you select a .p12 file" 
Write-Host "  2) Verify you know the .p12 password before uploading"
Write-Host "  3) Optionally let you select an SSH private key (defaulting to your .ssh folder)"
Write-Host "  4) Upload the .p12 to your Ubuntu server user's HOME directory using scp"
Write-Host "  5) Optionally set up and test SSH key-based login" 
Write-Host ""

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "P12 Files (*.p12)|*.p12|All Files (*.*)|*.*"
$dialog.Multiselect = $false
$dialog.Title = "Select your .p12 file"

$result = $dialog.ShowDialog()
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No file selected. Exiting."
    exit 1
}

$p12File = $dialog.FileName
Write-Host "Selected .p12 file:"
Write-Host "  $p12File"
Write-Host ""

$p12PasswordSecure = Test-P12Password -Path $p12File
if (-not $p12PasswordSecure) {
    exit 1
}

$useKey = Read-Host "Do you want to use an SSH key file for authentication? [Y/n]"
if ([string]::IsNullOrWhiteSpace($useKey)) { $useKey = "Y" }

$sshKey = $null

if ($useKey -match '^[Yy]') {
    Write-Host ""
    Write-Host "A file dialog will open in your .ssh folder if it exists." 
    Write-Host "Pick your SSH private key (for example: id_rsa, id_ed25519, etc.)."
    Write-Host ""

    $sshKey = Select-SshKeyFile -PromptTitle "Select your SSH private key"
    if ($sshKey) {
        Write-Host "Selected SSH key:"
        Write-Host "  $sshKey"
        Write-Host ""
    } else {
        Write-Host "No SSH key selected. Will fall back to password authentication." -ForegroundColor Yellow
        Write-Host ""
    }
}

$server = Read-Host "Enter server IP or hostname (Ubuntu server)"
$username = Read-Host "Enter SSH username (on the Ubuntu server)"

if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($username)) {
    Write-Host "Username and server are required. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Preparing to upload:"
Write-Host "  Local .p12 file: $p12File"
Write-Host "  Remote user:     $username"
Write-Host "  Remote host:     $server"
Write-Host "  Remote path:     (user's HOME directory: ~ )"
if ($sshKey) {
    Write-Host "  SSH key:         $sshKey"
} else {
    Write-Host "  SSH auth:        Password (no key selected)"
}
Write-Host ""

$uploadExit = Invoke-ScpWithHostHandling -Username $username -Server $server -PrivateKeyPath $sshKey -LocalPath $p12File

if ($uploadExit -eq 0) {
    Write-Host ""
    Write-Host "+ Upload completed successfully."
    Write-Host "  Your .p12 file should now be in the HOME directory of user '$username' on $server."
} else {
    Write-Host ""
    Write-Host "- Upload failed. Exit code: $uploadExit" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to close..."
    exit $uploadExit
}

if (-not $sshKey) {
    Write-Host ""
    $setup = Read-Host "You logged in using password authentication. Do you want to set up SSH key-based login for this server now? [Y/n]"
    if ([string]::IsNullOrWhiteSpace($setup)) { $setup = "Y" }

    if ($setup -match '^[Yy]') {
        Write-Host ""
        Write-Host "How would you like to proceed?" -ForegroundColor Cyan
        Write-Host "  1) Generate a NEW SSH key pair and install it on the server"
        Write-Host "  2) Import an EXISTING SSH private key and install its public key on the server"
        Write-Host "  3) Skip SSH key setup for now"
        Write-Host ""

        $choice = Read-Host "Enter 1, 2, or 3"
        switch ($choice) {
            '1' {
                if (-not (Ensure-SshTools)) { break }
                $newKeyPath = Generate-NewSshKeyPair -Username $username -Server $server
                if ($newKeyPath) {
                    if (Install-PubKeyToServer -Username $username -Server $server -PrivateKeyPath $newKeyPath) {
                        Write-Host ""
                        Write-Host "You can now use this key for future logins:" -ForegroundColor Green
                        Write-Host "  $newKeyPath"

                        $doTest = Read-Host "Do you want to test this key against the server now? [Y/n]"
                        if ([string]::IsNullOrWhiteSpace($doTest)) { $doTest = "Y" }
                        if ($doTest -match '^[Yy]') {
                            Test-SshKeyAgainstServer -Username $username -Server $server -PrivateKeyPath $newKeyPath | Out-Null
                        }
                    }
                }
            }
            '2' {
                if (-not (Ensure-SshTools)) { break }
                $importedKey = Import-ExistingSshKey
                if ($importedKey) {
                    if (Install-PubKeyToServer -Username $username -Server $server -PrivateKeyPath $importedKey) {
                        Write-Host ""
                        Write-Host "You can continue using this imported key for future logins:" -ForegroundColor Green
                        Write-Host "  $importedKey"

                        $doTest = Read-Host "Do you want to test this imported key against the server now? [Y/n]"
                        if ([string]::IsNullOrWhiteSpace($doTest)) { $doTest = "Y" }
                        if ($doTest -match '^[Yy]') {
                            Test-SshKeyAgainstServer -Username $username -Server $server -PrivateKeyPath $importedKey | Out-Null
                        }
                    }
                }
            }
            default {
                Write-Host "Skipping SSH key setup." -ForegroundColor Yellow
            }
        }
    }
}

Write-Host ""
Read-Host "Press Enter to close..."
