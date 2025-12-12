# ============================================================
#  STARDUST COLLECTIVE
#  SECURE P12 UPLOAD & SSH SETUP TOOL (WINDOWS)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms

$Cyan   = "Cyan"
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"

$script:P12FriendlyName   = $null
$script:SelectedP12File   = $null
$script:SelectedSshKeyFile = $null

function Select-P12File {
    param()

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "P12 Files (*.p12)|*.p12|All Files (*.*)|*.*"
        $dlg.Title  = "Select .p12 file"

        try {
            $docs = [Environment]::GetFolderPath('MyDocuments')
            if ($docs -and (Test-Path $docs)) {
                $dlg.InitialDirectory = $docs
            }
        } catch {}

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return (Resolve-Path $dlg.FileName).Path
        }
    } catch {
        Write-Host "! GUI file picker error: $($_.Exception.Message)" -ForegroundColor $Yellow
    }

    while ($true) {
        $path = Read-Host "Enter full path to your .p12 file (or press Enter to cancel)"
        if (-not $path) { return $null }

        if (Test-Path $path) {
            return (Resolve-Path $path).Path
        }

        Write-Host "- File not found at that path. Try again." -ForegroundColor $Red
    }
}

Clear-Host
Write-Host "==============================================================" -ForegroundColor $Cyan
Write-Host "                    STARDUST COLLECTIVE" -ForegroundColor $Cyan
Write-Host "           SECURE P12 UPLOAD & SSH SETUP TOOL" -ForegroundColor $Cyan
Write-Host "==============================================================" -ForegroundColor $Cyan
Write-Host ""
Write-Host "This guided tool will help you:" -ForegroundColor $Cyan
Write-Host "  1) Select your .p12 file"
Write-Host "  2) Verify its password locally"
Write-Host "  3) Enter server connection info"
Write-Host "  4) Choose SSH authentication (key or password)"
Write-Host "  5) Upload your .p12 securely"
Write-Host "  6) (Optional) Set up SSH key-based login" -ForegroundColor $Cyan
Write-Host ""
Read-Host "Press Enter to begin..."

function Test-IsSshPrivateKey {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $false }

    $lines = Get-Content -LiteralPath $Path -TotalCount 5
    foreach ($line in $lines) {
        if ($line -match '^-----BEGIN .*PRIVATE KEY-----') {
            return $true
        }
    }
    return $false
}

function Detect-PuTTYKey {
    param([string]$Path)

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1
    return ($firstLine -match '^PuTTY-User-Key-File-')
}

function Select-SshKeyFile {
    param([string]$PromptTitle = "Select your SSH private key")

    while ($true) {
        $file = $null

        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title  = $PromptTitle
            $dlg.Filter = "All Files (*.*)|*.*"

            $sshDir = Join-Path $HOME ".ssh"
            if (Test-Path $sshDir) {
                $dlg.InitialDirectory = $sshDir
            }

            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $file = $dlg.FileName
            } else {
                return $null
            }
        } catch {
            Write-Host "! SSH key file picker error: $($_.Exception.Message)" -ForegroundColor $Yellow
            $file = Read-Host "Enter full path to your SSH private key (or press Enter to cancel)"
            if (-not $file) { return $null }
        }

        if (-not (Test-Path $file)) {
            Write-Host "- File not found. Try again." -ForegroundColor $Red
            continue
        }

        if (Detect-PuTTYKey $file) {
            Write-Host "" 
            Write-Host "==============================================================" -ForegroundColor $Yellow
            Write-Host "  PuTTY Private Key Detected (.ppk)" -ForegroundColor $Yellow
            Write-Host "==============================================================" -ForegroundColor $Yellow
            Write-Host ""
            Write-Host "This tool requires a standard OpenSSH private key file." -ForegroundColor $Cyan
            Write-Host "Your selected file is a PuTTY Private Key (.ppk)." -ForegroundColor $Cyan
            Write-Host ""
            Write-Host "Follow these steps to export a normal SSH key from your .ppk:" -ForegroundColor $Green
            Write-Host ""
            Write-Host "  1) " -NoNewline; Write-Host "Open " -ForegroundColor $Green -NoNewline; Write-Host "PuTTYgen" -ForegroundColor $Yellow
            Write-Host "     - If you don't have it, download PuTTY from: https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html" -ForegroundColor $Yellow
            Write-Host ""
            Write-Host "  2) " -NoNewline; Write-Host "In PuTTYgen, click " -ForegroundColor $Green -NoNewline
            Write-Host "Conversions -> Import Key" -ForegroundColor $Yellow
            Write-Host "     - Select your .ppk file:" -ForegroundColor $Green
            Write-Host "       $file" -ForegroundColor $Cyan
            Write-Host ""
            Write-Host "  3) " -NoNewline; Write-Host "If prompted, enter your " -ForegroundColor $Green -NoNewline
            Write-Host "PPK passphrase" -ForegroundColor $Yellow
            Write-Host ""
            Write-Host "  4) " -NoNewline; Write-Host "In PuTTYgen, click " -ForegroundColor $Green -NoNewline
            Write-Host "Conversions -> Export OpenSSH key (force new file format)" -ForegroundColor $Yellow
            Write-Host "     - Choose a file name and location to save the new key." -ForegroundColor $Green
            Write-Host "     - Recommended: save it in your .ssh folder, for example:" -ForegroundColor $Green
            $sshDir = Join-Path $HOME ".ssh"
            Write-Host "       $sshDir\myserver_ed25519" -ForegroundColor $Cyan
            Write-Host ""
            Write-Host "  5) Close PuTTYgen when you are done." -ForegroundColor $Green
            Write-Host ""
            Write-Host "After you have exported the OpenSSH key," -ForegroundColor $Cyan
            Write-Host "you can select it in the next step." -ForegroundColor $Cyan
            Write-Host ""

            Read-Host "Press Enter once you have exported the OpenSSH key and are ready to select it..."

            continue
        }

        if (Test-IsSshPrivateKey $file) {
            return (Resolve-Path $file).Path
        }

        Write-Host "! Invalid SSH private key. Try again." -ForegroundColor $Yellow
    }
}

function Ensure-SshTools {
    $ssh = Get-Command ssh -ErrorAction SilentlyContinue
    $keygen = Get-Command ssh-keygen -ErrorAction SilentlyContinue

    return ($ssh -and $keygen)
}

function Test-P12Password {
    param([string]$Path)

    Write-Host "STEP 2: Verifying .p12 password..." -ForegroundColor $Cyan
    Write-Host "Your password is only checked locally." -ForegroundColor $Cyan
    Write-Host ""

    for ($i=1; $i -le 12; $i++) {
        $secure = Read-Host "Enter .p12 password (attempt $i of 12)" -AsSecureString

        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringUni($bstr)
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

        try {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($Path,$plain,'Exportable')

            Write-Host "+ Password verified." -ForegroundColor $Green

            if (-not [string]::IsNullOrWhiteSpace($cert.FriendlyName)) {
                $script:P12FriendlyName = $cert.FriendlyName
                Write-Host "+ P12 alias (friendlyName): $script:P12FriendlyName" -ForegroundColor $Green
                Write-Host "! Make sure to write this alias down and keep it documented for future use." -ForegroundColor $Yellow
            } else {
                Write-Host "! This P12 file does not contain a friendlyName/alias field." -ForegroundColor $Yellow
            }

            return $secure
        } catch {
            Write-Host "- Incorrect password." -ForegroundColor $Red
        }
    }

    Write-Host "- Too many incorrect attempts. Exiting." -ForegroundColor $Red
    return $null
}

function Generate-NewSshKeyPair {
    param([string]$Username,[string]$Server)

    $sshDir = Join-Path $HOME ".ssh"
    if (-not (Test-Path $sshDir)) { New-Item -ItemType Directory -Path $sshDir | Out-Null }

    while ($true) {
        $name = Read-Host "Enter a name for your new SSH key (no extension)"
        if (-not $name) {
            Write-Host "! Name cannot be empty." -ForegroundColor $Yellow
            continue
        }

        $path = Join-Path $sshDir $name
        if (Test-Path $path) {
            Write-Host "- A key with that name already exists. Choose another." -ForegroundColor $Red
            continue
        }

        Write-Host "Generating SSH key pair..." -ForegroundColor $Cyan

        & ssh-keygen -t ed25519 -f $path -N "" -C "nodecloud-$Username@$Server"

        if ($LASTEXITCODE -eq 0) {
            Write-Host "+ SSH key generated: $path" -ForegroundColor $Green
            return $path
        }

        Write-Host "- Failed to generate key. Try again." -ForegroundColor $Red
    }
}

function Get-KnownHostsPath {
    $profile = $env:USERPROFILE
    if (-not $profile) { $profile = $HOME }

    $dir = Join-Path $profile ".ssh"
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

    return (Join-Path $dir "known_hosts")
}

function Clean-HostFromKnownHosts {
    param([string]$Server)
    Write-Host "! Cleaning SSH host entry for $Server" -ForegroundColor $Yellow
    ssh-keygen -R $Server | Out-Null
}

function Invoke-ScpWithHostHandling {
    param(
        [string]$Username,
        [string]$Server,
        [string]$PrivateKeyPath,
        [string]$LocalPath
    )

    $known  = Get-KnownHostsPath
    $target = "$Username@$Server`:~/"

    $args = @()

    if ($PrivateKeyPath) {
        $normalizedKey = $PrivateKeyPath.Trim('"').Trim()

        try {
            $normalizedKey = [System.IO.Path]::GetFullPath($normalizedKey)
        } catch {
            Write-Host "! Invalid SSH key path: $PrivateKeyPath" -ForegroundColor $Red
            Write-Host "  $_" -ForegroundColor $Yellow
            return 1
        }

        if (-not (Test-Path $normalizedKey)) {
            Write-Host "! SSH private key file not found:" -ForegroundColor $Red
            Write-Host "  $normalizedKey" -ForegroundColor $Yellow
            return 1
        }

        Write-Host "Using SSH key: $normalizedKey" -ForegroundColor $Cyan

        $args += "-i"
        $args += $normalizedKey
    }

    $args += @(
        "-p",
        "-o","StrictHostKeyChecking=no",
        "-o","UserKnownHostsFile=$known",
        $LocalPath,
        $target
    )

    for ($i=1; $i -le 2; $i++) {
        Write-Host "Uploading .p12 file (attempt $i)..." -ForegroundColor $Cyan
        $out = & scp @args 2>&1

        if ($LASTEXITCODE -eq 0) {
            return 0
        }

        if ($out -match "IDENTIFICATION HAS CHANGED" -or $out -match "Offending") {
            Clean-HostFromKnownHosts $Server
            continue
        }

        Write-Host $out -ForegroundColor $Yellow
        return 1
    }

    return 1
}

function Install-PubKeyToServer {
    param([string]$Username,[string]$Server,[string]$KeyPath)

    Write-Host "Installing SSH public key on server..." -ForegroundColor $Cyan

    $pub = & ssh-keygen -y -f $KeyPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "- Failed to extract public key." -ForegroundColor $Red
        return $false
    }

    $known = Get-KnownHostsPath
    $cmd = "mkdir -p ~/.ssh && chmod 700 ~/.ssh && echo '$pub' >> ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys"

    $out = & ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=$known "$Username@$Server" $cmd 2>&1

    if ($LASTEXITCODE -ne 0) {
        if ($out -match "IDENTIFICATION HAS CHANGED" -or $out -match "Offending") {
            Clean-HostFromKnownHosts $Server
            $out = & ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=$known "$Username@$Server" $cmd 2>&1
        }
    }

    if ($LASTEXITCODE -eq 0) {
        Write-Host "+ SSH key installed." -ForegroundColor $Green
        return $true
    }

    Write-Host "- Failed to install SSH key." -ForegroundColor $Red
    Write-Host $out -ForegroundColor $Yellow
    return $false
}

function Test-SshKeyAgainstServer {
    param([string]$Username,[string]$Server,[string]$KeyPath)

    Write-Host "Testing SSH key authentication..." -ForegroundColor $Cyan
    Write-Host "SSH may prompt you for the key passphrase." -ForegroundColor $Yellow

    $known = Get-KnownHostsPath
    & ssh -i $KeyPath -o StrictHostKeyChecking=no -o UserKnownHostsFile=$known "$Username@$Server" "echo ok" 2>&1 | Out-Null

    if ($LASTEXITCODE -eq 0) {
        Write-Host "+ SSH key authentication succeeded." -ForegroundColor $Green
        return $true
    }

    Write-Host "- SSH key authentication failed." -ForegroundColor $Red
    return $false
}

Write-Host ""
Write-Host "STEP 1: Select your .p12 file" -ForegroundColor $Cyan

$p12File = Select-P12File
if (-not $p12File) {
    Write-Host "- No file selected. Exiting." -ForegroundColor $Red
    exit
}

Write-Host "+ Selected: $p12File" -ForegroundColor $Green
Write-Host ""

$p12Pass = Test-P12Password $p12File
if (-not $p12Pass) { exit }

Write-Host ""
Write-Host "STEP 3: Enter server IP or hostname" -ForegroundColor $Cyan
$server = Read-Host "Server IP"
if (-not $server) {
    Write-Host "- Server IP required. Exiting." -ForegroundColor $Red
    exit
}

if (Get-Command ssh-keygen -ErrorAction SilentlyContinue) {
    Write-Host "! Cleaning SSH host entry for $server" -ForegroundColor $Yellow
    ssh-keygen -R $server | Out-Null
}

Write-Host ""
Write-Host "STEP 4: Enter server username" -ForegroundColor $Cyan
$username = Read-Host "Username"
if (-not $username) {
    Write-Host "- Username required. Exiting." -ForegroundColor $Red
    exit
}

Write-Host ""
Write-Host "STEP 5: Choose SSH authentication method" -ForegroundColor $Cyan
$useKey = Read-Host "Use SSH private key? [Y/n]"
if ([string]::IsNullOrWhiteSpace($useKey)) { $useKey = "Y" }

$sshKey = $null

if ($useKey -match "^[Yy]") {
    Write-Host ""
    Write-Host "Select your SSH private key..." -ForegroundColor $Cyan
    $sshKey = Select-SshKeyFile
    if ($sshKey) {
        Write-Host "+ Using SSH key: $sshKey" -ForegroundColor $Green
    } else {
        Write-Host "! No key selected. Falling back to password authentication." -ForegroundColor $Yellow
    }
}

Write-Host ""
Write-Host "STEP 6: Uploading your .p12 file" -ForegroundColor $Cyan

$upload = Invoke-ScpWithHostHandling -Username $username -Server $server -PrivateKeyPath $sshKey -LocalPath $p12File

if ($upload -eq 0) {
    Write-Host "+ Upload successful." -ForegroundColor $Green
} else {
    Write-Host "- Upload failed." -ForegroundColor $Red
    Read-Host "Press Enter to exit..."
    exit
}

if (-not $sshKey) {
    Write-Host ""
    Write-Host "STEP 7: Optional SSH key setup" -ForegroundColor $Cyan

    $do = Read-Host "Set up SSH key-based login now? [Y/n]"
    if ([string]::IsNullOrWhiteSpace($do)) { $do = "Y" }

    if ($do -match "^[Yy]") {

        Write-Host ""
        Write-Host "Choose an option:" -ForegroundColor $Cyan
        Write-Host "  1) Generate a NEW SSH key pair"
        Write-Host "  2) Import an EXISTING SSH key"
        Write-Host "  3) Skip SSH key setup"
        Write-Host ""

        $choice = Read-Host "Enter 1, 2, or 3"

        switch ($choice) {

            "1" {
                if (-not (Ensure-SshTools)) { break }

                $newKey = Generate-NewSshKeyPair -Username $username -Server $server

                if ($newKey) {
                    if (Install-PubKeyToServer -Username $username -Server $server -KeyPath $newKey) {
                        Test-SshKeyAgainstServer -Username $username -Server $server -KeyPath $newKey | Out-Null
                    }
                }
            }

            "2" {
                if (-not (Ensure-SshTools)) { break }

                $import = Select-SshKeyFile -PromptTitle "Select existing SSH private key to import"

                if ($import) {
                    if (Install-PubKeyToServer -Username $username -Server $server -KeyPath $import) {
                        Test-SshKeyAgainstServer -Username $username -Server $server -KeyPath $import | Out-Null
                    }
                }
            }

            default {
                Write-Host "! SSH key setup skipped." -ForegroundColor $Yellow
            }
        }
    }
}

if ($script:P12FriendlyName) {
    Write-Host ""
    Write-Host "REMINDER: The alias (friendlyName) for this P12 is: $script:P12FriendlyName" -ForegroundColor $Cyan
    Write-Host "Please keep this alias documented somewhere safe. You may need it later for tools or imports." -ForegroundColor $Yellow
}

Write-Host ""
Write-Host "Login reminder:" -ForegroundColor $Cyan
if ($sshKey) {
    Write-Host "You can log into your server using this command:" -ForegroundColor $Cyan
    Write-Host "ssh -i `"$sshKey`" $username@$server" -ForegroundColor $Green
} else {
    Write-Host "You can log into your server using this command:" -ForegroundColor $Cyan
    Write-Host "ssh $username@$server" -ForegroundColor $Green
}

Write-Host ""
Read-Host "Press Enter to exit..."

