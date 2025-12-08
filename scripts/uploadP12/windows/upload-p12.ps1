Add-Type -AssemblyName System.Windows.Forms

Write-Host "This script will:" -ForegroundColor Cyan
Write-Host "  1) Let you select a .p12 file" 
Write-Host "  2) Optionally let you select an SSH private key (defaulting to your .ssh folder)"
Write-Host "  3) Upload the .p12 to your Ubuntu server user's HOME directory using scp"
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

$useKey = Read-Host "Do you want to use an SSH key file for authentication? [Y/n]"
if ([string]::IsNullOrWhiteSpace($useKey)) { $useKey = "Y" }

$sshKey = $null

if ($useKey -match '^[Yy]') {
    Write-Host ""
    Write-Host "A file dialog will open in your .ssh folder if it exists." 
    Write-Host "Pick your SSH private key (for example: id_rsa, id_ed25519, etc.)."
    Write-Host ""

    $sshDialog = New-Object System.Windows.Forms.OpenFileDialog
    $sshDialog.Multiselect = $false
    $sshDialog.Title = "Select your SSH private key"
    $sshDialog.Filter = "All Files (*.*)|*.*"

    $sshDir = Join-Path $HOME ".ssh"
    if (Test-Path $sshDir) {
        $sshDialog.InitialDirectory = $sshDir
    }

    $sshResult = $sshDialog.ShowDialog()
    if ($sshResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $sshKey = $sshDialog.FileName
        Write-Host "Selected SSH key:"
        Write-Host "  $sshKey"
        Write-Host ""
    } else {
        Write-Host "No SSH key selected. Will fall back to password authentication."
        Write-Host ""
    }
}

$server = Read-Host "Enter server IP or hostname (Ubuntu server)"
$username = Read-Host "Enter SSH username (on the Ubuntu server)"

if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($username)) {
    Write-Host "Username and server are required. Exiting."
    exit 1
}

$remoteTarget = "$username@$server`:~/"

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

if ($sshKey) {
    scp -i "$sshKey" "`"$p12File`"" "$remoteTarget"
} else {
    scp "`"$p12File`"" "$remoteTarget"
}

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "+ Upload completed successfully."
    Write-Host "Your .p12 file should now be in the HOME directory of user '$username' on $server."
} else {
    Write-Host ""
    Write-Host "- Upload failed. Exit code: $LASTEXITCODE"
}

Write-Host ""
Read-Host "Press Enter to close..."
