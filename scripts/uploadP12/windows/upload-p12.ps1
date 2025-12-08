# ============================================================
#  STARDUST COLLECTIVE
#  SECURE P12 UPLOAD & SSH SETUP TOOL (WINDOWS)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms

$Cyan   = "Cyan"
$Green  = "Green"
$Yellow = "Yellow"
$Red    = "Red"

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
        if ($line -match '^(-----BEGIN|PuTTY-User-Key-File-)') {
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

function Find-PuTTYgen {
    $paths = @(
        "C:\Program Files\PuTTY\puttygen.exe",
        "C:\Program Files (x86)\PuTTY\puttygen.exe"
    )

    foreach ($p in $paths) {
        if (Test-Path $p) { return $p }
    }

    $gcm = Get-Command puttygen.exe -ErrorAction SilentlyContinue
    if ($gcm) { return $gcm.Source }

    return $null
}

function Install-PuTTYgen {
    Write-Host "! PuTTYgen not found." -ForegroundColor $Yellow
    Write-Host ""

    $choice = Read-Host "Install PuTTYgen now? (via Chocolatey if available, otherwise direct download) [Y/n]"
    if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "Y" }

    if ($choice -match "^[Nn]") {
        Write-Host "- PuTTYgen is required to convert this key. Aborting PuTTY key usage." -ForegroundColor $Red
        return $null
    }

    $choco = Get-Command choco -ErrorAction SilentlyContinue
    if ($choco) {
        Write-Host "Installing PuTTY via Chocolatey..." -ForegroundColor $Cyan
        choco install putty -y | Out-Null
        $gen = Find-PuTTYgen
        if ($gen) {
            Write-Host "+ PuTTYgen installed successfully." -ForegroundColor $Green
            return $gen
        }
    }

    Write-Host "Downloading PuTTYgen.exe..." -ForegroundColor $Cyan

    $url = "https://the.earth.li/~sgtatham/putty/latest/w64/puttygen.exe"
    $dest = Join-Path $env:TEMP "puttygen.exe"

    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
        Write-Host "+ PuTTYgen downloaded: $dest" -ForegroundColor $Green
        return $dest
    } catch {
        Write-Host "- Failed to download PuTTYgen." -ForegroundColor $Red
        return $null
    }
}

function Convert-PuTTYKey {
    param(
        [string]$PpkPath,
        [string]$OutputPath
    )

    $gen = Find-PuTTYgen
    if (-not $gen) {
        $gen = Install-PuTTYgen
        if (-not $gen) { return $null }
    }

    Write-Host "Converting PuTTY (.ppk) key to OpenSSH format..." -ForegroundColor $Cyan

    $cmd = "`"$gen`" `"$PpkPath`" -O private-openssh -o `"$OutputPath`""
    cmd.exe /c $cmd | Out-Null

    if (Test-IsSshPrivateKey $OutputPath) {
        Write-Host "+ Conversion successful." -ForegroundColor $Green
        return $OutputPath
    }

    Write-Host "- Conversion failed." -ForegroundColor $Red
    return $null
}

function Select-SshKeyFile {
    param([string]$PromptTitle = "Select your SSH private key")

    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Title = $PromptTitle
    $dlg.Filter = "All Files (*.*)|*.*"
    $sshDir = Join-Path $HOME ".ssh"
    if (Test-Path $sshDir) { $dlg.InitialDirectory = $sshDir }

    while ($true) {
        if ($dlg.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { return $null }
        $file = $dlg.FileName

        if (Detect-PuTTYKey $file) {
            Write-Host "! PuTTY key detected (.ppk). Conversion is required." -ForegroundColor $Yellow

            $name = Read-Host "Enter a name for the converted SSH key (no extension)"
            $dest = Join-Path $sshDir $name

            if (Test-Path $dest) {
                Write-Host "- A key with this name already exists. Choose a different name." -ForegroundColor $Red
                continue
            }

            $converted = Convert-PuTTYKey -PpkPath $file -OutputPath $dest
            return $converted
        }

        if (Test-IsSshPrivateKey $file) {
            return $file
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

    $known = Get-KnownHostsPath
    $target = "$Username@$Server`:~/"

    $args = @()

    if ($PrivateKeyPath) { $args += "-i"; $args += $PrivateKeyPath }

    $args += @(
        "-o","StrictHostKeyChecking=no",
        "-o","UserKnownHostsFile=$known",
        $LocalPath,
        $target
    )

    for ($i=1; $i -le 2; $i++) {
        Write-Host "Uploading .p12 file (attempt $i)..." -ForegroundColor $Cyan
        $out = & scp @args 2>&1
        if ($LASTEXITCODE -eq 0) { return 0 }

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

$dlg = New-Object Windows.Forms.OpenFileDialog
$dlg.Filter = "P12 Files (*.p12)|*.p12|All Files (*.*)|*.*"
$dlg.Title = "Select .p12 file"
if ($dlg.ShowDialog() -ne 'OK') {
    Write-Host "- No file selected. Exiting." -ForegroundColor $Red
    exit
}

$p12File = $dlg.FileName
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
    if ([string]::IsNullOrWhiteSpace($do)) { $do="Y" }

    if ($do -match "^[Yy]") {
        Write-Host ""
        Write-Host "Choose an option:" -
