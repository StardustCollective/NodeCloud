<# 
    server-toolkit.ps1
    Stardust Collective - Server Setup Toolkit (Windows Edition, GUI)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Global config, logging, colors

$Script:HomeDir    = [Environment]::GetFolderPath("UserProfile")
$Script:ProfilesDir = Join-Path $Script:HomeDir ".ssh"
$Script:KnownHosts  = Join-Path $Script:ProfilesDir "known_hosts"
$Script:LogPath     = Join-Path $Script:HomeDir "server_setup.log"

if (-not (Test-Path $Script:ProfilesDir)) {
    New-Item -ItemType Directory -Path $Script:ProfilesDir -Force | Out-Null
}
if (-not (Test-Path $Script:KnownHosts)) {
    New-Item -ItemType File -Path $Script:KnownHosts -Force | Out-Null
}

# colors
$Script:Color_Background   = "#2b2b2b"
$Script:Color_Text         = "#ffffff"
$Script:Color_Panel        = "#3c3f41"
$Script:Color_Button       = "#1284b6"
$Script:Color_ButtonHover  = "#13628b"
$Script:Color_Accent       = "#11c3f0"
$Script:Color_Warning      = "#ffc857"
$Script:Color_Error        = "#ff4d4d"
$Script:Color_Success      = "#67e480"

# in-memory cached connections
$Script:NewServerConn = $null
$Script:OldServerConn = $null

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    try {
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $line = "{0} [{1}] {2}" -f $timestamp, $Level.ToUpper(), $Message
        Add-Content -Path $Script:LogPath -Value $line
    } catch {
        # last resort: ignore logging failures
    }
}

function Write-Info  { param([string]$Msg) Write-Log $Msg "INFO"  }
function Write-Warn  { param([string]$Msg) Write-Log $Msg "WARN"  }
function Write-ErrorLog { param([string]$Msg) Write-Log $Msg "ERROR" }

#endregion

#region WPF helpers (load, style, message boxes)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

function New-Window {
    param(
        [string]$Title,
        [int]$Width = 600,
        [int]$Height = 400
    )

    $window = New-Object Windows.Window
    $window.Title = $Title
    $window.Width = $Width
    $window.Height = $Height
    $window.WindowStartupLocation = "CenterScreen"
    $window.ResizeMode = "CanMinimize"
    $window.Background = $Color_Background
    $window.Foreground = $Color_Text
    $window.FontFamily = "Segoe UI"
    $window.FontSize = 12
    return $window
}

function New-StackPanel {
    param(
        [string]$Orientation = "Vertical",
        [int]$Margin = 10
    )
    $sp = New-Object Windows.Controls.StackPanel
    $sp.Orientation = $Orientation
    $sp.Margin = [Windows.Thickness]::new($Margin)
    return $sp
}

function New-Label {
    param(
        [string]$Text,
        [int]$FontSize = 12,
        [bool]$Bold = $false
    )
    $lbl = New-Object Windows.Controls.TextBlock
    $lbl.Text = $Text
    $lbl.Foreground = $Color_Text
    $lbl.TextWrapping = "Wrap"
    $lbl.Margin = "0,0,0,5"
    $lbl.FontSize = $FontSize
    if ($Bold) {
        $lbl.FontWeight = "Bold"
    }
    return $lbl
}

function New-TextBox {
    param(
        [string]$Text = "",
        [bool]$IsPassword = $false
    )
    if ($IsPassword) {
        $pwd = New-Object Windows.Controls.PasswordBox
        $pwd.Margin = "0,0,0,8"
        $pwd.Background = $Color_Panel
        $pwd.Foreground = $Color_Text
        $pwd.BorderBrush = "#5c5c5c"
        $pwd.BorderThickness = 1
        $pwd.Padding = "4,2,4,2"
        return $pwd
    } else {
        $tb = New-Object Windows.Controls.TextBox
        $tb.Text = $Text
        $tb.Margin = "0,0,0,8"
        $tb.Background = $Color_Panel
        $tb.Foreground = $Color_Text
        $tb.BorderBrush = "#5c5c5c"
        $tb.BorderThickness = 1
        $tb.Padding = "4,2,4,2"
        return $tb
    }
}

function New-Button {
    param(
        [string]$Content,
        [int]$Height = 32,
        [int]$MarginTop = 8
    )
    $btn = New-Object Windows.Controls.Button
    $btn.Content = $Content
    $btn.Height = $Height
    $btn.Margin = "0,$MarginTop,0,0"
    $btn.Background = $Color_Button
    $btn.Foreground = $Color_Text
    $btn.BorderThickness = 0
    $btn.Padding = "8,2,8,2"
    $btn.HorizontalAlignment = "Stretch"

    $btn.Add_MouseEnter({
        param($s,$e)
        $s.Background = $Color_ButtonHover
    })
    $btn.Add_MouseLeave({
        param($s,$e)
        $s.Background = $Color_Button
    })
    return $btn
}

function Show-MessageBoxInfo {
    param(
        [string]$Message,
        [string]$Title = "Server Toolkit"
    )
    Write-Info "$($Title): $Message"
    [Windows.MessageBox]::Show($Message, $Title, 'OK', 'Information') | Out-Null
}

function Show-MessageBoxWarn {
    param(
        [string]$Message,
        [string]$Title = "Server Toolkit"
    )
    Write-Warn "$($Title): $Message"
    [Windows.MessageBox]::Show($Message, $Title, 'OK', 'Warning') | Out-Null
}

function Show-MessageBoxError {
    param(
        [string]$Message,
        [string]$Title = "Server Toolkit"
    )
    Write-ErrorLog "$($Title): $Message"
    [Windows.MessageBox]::Show($Message, $Title, 'OK', 'Error') | Out-Null
}

function Show-Confirm {
    param(
        [string]$Message,
        [string]$Title = "Confirm",
        [bool]$DefaultYes = $false
    )
    Write-Info "Confirm: $Message"
    $defaultButton = if ($DefaultYes) { 'Yes' } else { 'No' }
    $result = [Windows.MessageBox]::Show($Message, $Title, 'YesNo', 'Question')
    return ($result -eq 'Yes')
}

#endregion

#region Profiles: parsing, saving, listing

class ConnectionProfile {
    [string]$Name
    [string]$Host
    [string]$User
    [string]$Port
    [string]$IdentityFile

    ConnectionProfile([string]$name,[string]$hostName,[string]$user,[string]$port,[string]$identity) {
        $this.Name        = $name
        $this.Host        = $hostName
        $this.User        = $user
        $this.Port        = $port
        $this.IdentityFile = $identity
    }
}

function Test-ProfileNameValid {
    param([string]$Name)
    return ($Name -match '^[A-Za-z0-9-]+$')
}

function Get-ProfilePath {
    param([string]$Name)
    return (Join-Path $Script:ProfilesDir ("{0}_ssh_config.txt" -f $Name))
}

function Get-StoredProfiles {
    Write-Info "Scanning for stored profiles in $Script:ProfilesDir"
    $files = Get-ChildItem -Path $Script:ProfilesDir -Filter "*_ssh_config.txt" -ErrorAction SilentlyContinue
    if (-not $files) {
        Write-Info "No *_ssh_config.txt profiles found."
        return @()
    }

    $profiles = @()
    foreach ($f in $files) {
        try {
            $content = Get-Content -LiteralPath $f.FullName -ErrorAction Stop
        } catch {
            Write-Warn "Failed to read profile file: $($f.FullName): $_"
            continue
        }

        $hostLine     = $content | Where-Object { $_ -match '^\s*Host\s+' }      | Select-Object -First 1
        $hostNameLine = $content | Where-Object { $_ -match '^\s*HostName\s+' }  | Select-Object -First 1
        $userLine     = $content | Where-Object { $_ -match '^\s*User\s+' }      | Select-Object -First 1
        $portLine     = $content | Where-Object { $_ -match '^\s*Port\s+' }      | Select-Object -First 1
        $identLine    = $content | Where-Object { $_ -match '^\s*IdentityFile\s+' } | Select-Object -First 1

        if (-not $hostLine -or -not $hostNameLine -or -not $userLine) {
            Write-Warn "Skipping invalid ssh_config file: $($f.FullName)"
            continue
        }

        $name          = ($hostLine     -replace '^\s*Host\s+', '').Trim()
        $hostNameValue = ($hostNameLine -replace '^\s*HostName\s+', '').Trim()
        $user          = ($userLine     -replace '^\s*User\s+', '').Trim()
        $port          = "22"
        if ($portLine) {
            $port = ($portLine -replace '^\s*Port\s+', '').Trim()
        }
        $identity = ""
        if ($identLine) {
            $identity = ($identLine -replace '^\s*IdentityFile\s+', '').Trim()
        }

        $profiles += [ConnectionProfile]::new($name,$hostNameValue,$user,$port,$identity)
    }

    $profiles = $profiles | Sort-Object Name
    Write-Info "Found $($profiles.Count) stored profile(s)."
    return $profiles
}
function Save-SSHProfile {
    param(
        [ConnectionProfile]$Profile
    )

    if (-not (Test-ProfileNameValid $Profile.Name)) {
        Show-MessageBoxError "Invalid profile name. Use letters, numbers, and dashes only." "Profile Name"
        return
    }

    $file = Get-ProfilePath -Name $Profile.Name
    Write-Info "Saving profile '$($Profile.Name)' to $file"

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
    $login = "ssh"
    $sftp  = "sftp"
    if ($Profile.IdentityFile) {
        $login += " -i $($Profile.IdentityFile)"
        $sftp  += " -i $($Profile.IdentityFile)"
    }
    $login += " $($Profile.User)@$($Profile.Host)"
    $sftp  += " $($Profile.User)@$($Profile.Host)"
    $lines += "# SSH login: $login"
    $lines += "# SFTP: $sftp"

    Set-Content -LiteralPath $file -Value $lines -Encoding UTF8 -Force
    Write-Info "Profile '$($Profile.Name)' saved."
}

#endregion

#region PuTTY / PuTTYgen helpers (.ppk conversion)

function Test-IsPuTTYKey {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return $false }
    try {
        $firstLine = (Get-Content -LiteralPath $FilePath -TotalCount 1 -ErrorAction Stop)
        return ($firstLine -match '^PuTTY-User-Key-File-')
    } catch {
        return $false
    }
}

function Find-PuTTYgen {
    Write-Info "Locating PuTTYgen.exe"
    $candidates = @()

    $pf = ${env:ProgramFiles}
    $pf86 = ${env:ProgramFiles(x86)}

    if ($pf)   { $candidates += (Join-Path $pf   "PuTTY\puttygen.exe") }
    if ($pf86) { $candidates += (Join-Path $pf86 "PuTTY\puttygen.exe") }

    # PATH lookup
    $cmd = Get-Command "puttygen.exe" -ErrorAction SilentlyContinue
    if ($cmd) {
        $candidates += $cmd.Source
    }

    foreach ($c in $candidates | Select-Object -Unique) {
        if (Test-Path $c) {
            Write-Info "Found PuTTYgen at $c"
            return $c
        }
    }

    Write-Warn "PuTTYgen not found in standard locations."
    return $null
}

function Install-PuTTYgen {
    # best effort; user may not have Chocolatey; fallback is manual download
    Write-Info "Attempting to install or download PuTTYgen."

    $pg = Find-PuTTYgen
    if ($pg) { return $pg }

    $hasChoco = Get-Command choco -ErrorAction SilentlyContinue
    if ($hasChoco) {
        if (Show-Confirm "PuTTYgen not found. Install PuTTY via Chocolatey now?" "PuTTYgen Install" $true) {
            try {
                choco install putty -y | Out-Null
                $pg = Find-PuTTYgen
                if ($pg) { return $pg }
            } catch {
                Write-ErrorLog "Chocolatey install of PuTTY failed: $_"
            }
        }
    }

    # fallback: download puttygen.exe to temp
    if (Show-Confirm "PuTTYgen still not found. Download a portable PuTTYgen.exe to your TEMP folder?" "PuTTYgen Download") {
        try {
            $url = "https://the.earth.li/~sgtatham/putty/latest/w64/puttygen.exe"
            $dest = Join-Path $env:TEMP "puttygen.exe"
            Write-Info "Downloading PuTTYgen from $url to $dest"
            Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -ErrorAction Stop
            if (Test-Path $dest) {
                Write-Info "PuTTYgen downloaded to $dest"
                return $dest
            }
        } catch {
            Write-ErrorLog "Failed to download PuTTYgen: $_"
        }
    }

    Show-MessageBoxWarn "PuTTYgen could not be installed or downloaded. .ppk files cannot be converted." "PuTTYgen"
    return $null
}

function Convert-PuTTYKey {
    param(
        [string]$PpkPath
    )

    if (-not (Test-Path $PpkPath)) {
        Show-MessageBoxError "PuTTY key file not found: $PpkPath" "PuTTY Key"
        return $null
    }

    $puttygen = Find-PuTTYgen
    if (-not $puttygen) {
        $puttygen = Install-PuTTYgen
        if (-not $puttygen) { return $null }
    }

    $dest = [System.IO.Path]::ChangeExtension($PpkPath, ".openssh.key")
    if (Test-Path $dest) {
        if (-not (Show-Confirm "Converted key '$dest' already exists. Overwrite it?" "PuTTY Conversion")) {
            return $dest
        }
    }

    Write-Info "Converting PuTTY key $PpkPath to OpenSSH format at $dest"
    $args = "`"$PpkPath`" -O private-openssh -o `"$dest`""
    $proc = Start-Process -FilePath $puttygen -ArgumentList $args -PassThru -WindowStyle Hidden -Wait
    if ($proc.ExitCode -eq 0 -and (Test-Path $dest)) {
        Show-MessageBoxInfo "PuTTY key converted successfully to $dest" "PuTTY Conversion"
        Write-Info "PuTTY->OpenSSH conversion succeeded: $dest"
        return $dest
    } else {
        Show-MessageBoxError "PuTTYgen conversion failed. ExitCode: $($proc.ExitCode)" "PuTTY Conversion"
        Write-ErrorLog "PuTTYgen conversion failed. ExitCode: $($proc.ExitCode)"
        return $null
    }
}

#endregion

#region File selection (SSH key, P12) via GUI

Add-Type -AssemblyName System.Windows.Forms

function Select-SSHKeyFile {
    Write-Info "Opening SSH key file browser."
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select SSH Private Key"
    $dlg.InitialDirectory = (Join-Path $Script:HomeDir ".ssh")
    $dlg.Filter = "All files (*.*)|*.*"

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Info "User cancelled SSH key file selection."
        return $null
    }

    $path = $dlg.FileName
    Write-Info "User selected SSH key file: $path"

    if (Test-IsPuTTYKey $path) {
        if (Show-Confirm "The selected key appears to be a PuTTY .ppk file. Convert it to OpenSSH format now?" "PuTTY Key" $true) {
            $converted = Convert-PuTTYKey -PpkPath $path
            return $converted
        } else {
            return $null
        }
    }

    return $path
}

function Select-P12File {
    Write-Info "Opening .p12 file browser."
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select .p12 File"
    $dlg.InitialDirectory = (Join-Path $Script:HomeDir "Downloads")
    $dlg.Filter = "P12 files (*.p12)|*.p12|All files (*.*)|*.*"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Info "User cancelled .p12 file selection."
        return $null
    }
    $path = $dlg.FileName
    Write-Info "Selected .p12 file: $path"
    return $path
}

#endregion

#region SSH helpers (ssh, scp, host-key cleanup)

function Prepare-HostKey {
    param(
        [string]$HostName,
        [string]$Port = "22"
    )

    Write-Info "Prepare-HostKey called for host: $HostName (port $Port)"

    try {
        Write-Info "Removing stale host keys for $HostName from known_hosts (if any)."
        & ssh-keygen -R $HostName       2>$null | Out-Null
        & ssh-keygen -R "[$HostName]:$Port" 2>$null | Out-Null
    } catch {
        Write-ErrorLog ("ssh-keygen -R failed for host {0}: {1}" -f $HostName, $_)
    }

    try {
        Write-Info ("Running ssh-keyscan for {0}:{1}" -f $HostName, $Port)
        $scan = & ssh-keyscan -p $Port $HostName 2>$null
        if ($scan) {
            if (-not (Test-Path $Script:KnownHosts)) {
                New-Item -ItemType File -Path $Script:KnownHosts -Force | Out-Null
            }
            $scan | Add-Content -Path $Script:KnownHosts
            Write-Info "Updated known_hosts entry for $HostName."
        } else {
            Write-Warn "ssh-keyscan returned no data for $HostName; known_hosts will be updated on first real ssh connection."
        }
    } catch {
        Write-ErrorLog ("ssh-keyscan failed for host {0}: {1}" -f $HostName, $_)
    }
}

function Invoke-SshCommand {
    param(
        [string]$HostName,
        [string]$User,
        [string]$Port = "22",
        [string]$IdentityFile,
        [string]$Command,
        [string]$InputData,
        [int]   $TimeoutMs = 30000
    )

    $args = @(
        "-p", $Port,
        "-o", "StrictHostKeyChecking=no",
        "-o", "UserKnownHostsFile=$($Script:KnownHosts)",
        "-o", "BatchMode=yes"
    )

    if ($IdentityFile) {
        $args += @("-i", $IdentityFile)
    }

    $args += "$User@$HostName"

    if ($Command) {
        $args += $Command
    }

    Write-Info "Running ssh command: ssh $($args -join ' ')"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "ssh.exe"
    $psi.Arguments              = ($args -join " ")
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.RedirectStandardInput  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()

    if ($InputData) {
        Write-Info "Invoke-SshCommand: sending InputData to stdin (length=$($InputData.Length))."
        $p.StandardInput.Write($InputData)
        $p.StandardInput.Close()
    }

    # HARD TIMEOUT so the GUI can't hang forever
    $timeoutMs = 30000
    $exited = $p.WaitForExit($timeoutMs)

    if (-not $exited) {
        Write-Warn "Invoke-SshCommand: ssh timed out after $timeoutMs ms for $User@$HostName. Killing process."
        try { $p.Kill() } catch {}
        return [PSCustomObject]@{
            ExitCode = 999
            StdOut   = ""
            StdErr   = "ssh command timed out after $timeoutMs ms for $User@$HostName"
        }
    }

    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()

    if ($stderr -match "REMOTE HOST IDENTIFICATION HAS CHANGED") {
        Write-Warn "Host key mismatch detected for $HostName. Cleaning known_hosts entry and retrying."
        try {
            & ssh-keygen -R $HostName | Out-Null
        } catch {
            Write-ErrorLog ("ssh-keygen -R failed for host {0}: {1}" -f $HostName, $_)
        }

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        [void]$p.Start()

        if ($InputData) {
            Write-Info "Invoke-SshCommand: retry sending InputData to stdin (length=$($InputData.Length))."
            $p.StandardInput.Write($InputData)
            $p.StandardInput.Close()
        }

        $exited = $p.WaitForExit($timeoutMs)
        if (-not $exited) {
            Write-Warn "Invoke-SshCommand (retry): ssh timed out after $timeoutMs ms for $User@$HostName. Killing process."
            try { $p.Kill() } catch {}
            return [PSCustomObject]@{
                ExitCode = 999
                StdOut   = ""
                StdErr   = "ssh command (retry) timed out after $timeoutMs ms for $User@$HostName"
            }
        }

        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
    }

    if ($stderr) {
        Write-Warn ("ssh stderr for {0}: {1}" -f $HostName, $stderr)
    }

    Write-Info ("Invoke-SshCommand finished for {0}@{1} with ExitCode={2}" -f $User, $HostName, $p.ExitCode)

    # Trim outputs in the log but keep full text in return object
    if ($stdout) {
        $sample = ($stdout -split "`r?`n")[0..([Math]::Min(4, ($stdout -split "`r?`n").Count-1))] -join " | "
        Write-Info "ssh stdout (first lines): $sample"
    }

    if ($stderr) {
        $sampleErr = ($stderr -split "`r?`n")[0..([Math]::Min(4, ($stderr -split "`r?`n").Count-1))] -join " | "
        Write-Warn "ssh stderr (first lines): $sampleErr"
    }

    return [PSCustomObject]@{
        ExitCode = $p.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Invoke-ScpDownload {
    param(
        [string]$HostName,
        [string]$User,
        [string]$Port = "22",
        [string]$IdentityFile,
        [string]$RemotePath,
        [string]$LocalPath
    )

    $args = @("-P", $Port)
    if ($IdentityFile) {
        $args += @("-i", $IdentityFile)
    }
    $args += "$User@$($HostName):`"$RemotePath`""
    $args += "`"$LocalPath`""

    Write-Info "Running scp download: scp $($args -join ' ')"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "scp.exe"
    $psi.Arguments = ($args -join " ")
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    if ($stderr -match "REMOTE HOST IDENTIFICATION HAS CHANGED") {
        Write-Warn "Host key mismatch for $HostName during scp. Cleaning known_hosts and retrying."
        try {
            & ssh-keygen -R $HostName | Out-Null
        } catch {
            Write-ErrorLog ("ssh-keygen -R failed for host {0}: {1}" -f $HostName, $_)
        }

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        [void]$p.Start()
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()
    }

    if ($stderr) {
        Write-Warn "scp stderr: $stderr"
    }

    return [PSCustomObject]@{
        ExitCode = $p.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

function Invoke-ScpUpload {
    param(
        [string]$HostName,
        [string]$User,
        [string]$Port = "22",
        [string]$IdentityFile,
        [string]$LocalPath,
        [string]$RemotePath
    )

    $args = @("-P", $Port)
    if ($IdentityFile) {
        $args += @("-i", $IdentityFile)
    }
    $args += "`"$LocalPath`""
    $args += "$User@$($HostName):`"$RemotePath`""

    Write-Info "Running scp upload: scp $($args -join ' ')"

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "scp.exe"
    $psi.Arguments = ($args -join " ")
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    if ($stderr -match "REMOTE HOST IDENTIFICATION HAS CHANGED") {
        Write-Warn "Host key mismatch for $HostName during scp. Cleaning known_hosts and retrying."
        try {
            & ssh-keygen -R $HostName | Out-Null
        } catch {
            Write-ErrorLog ("ssh-keygen -R failed for host {0}: {1}" -f $HostName, $_)
        }

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        [void]$p.Start()
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        $p.WaitForExit()
    }

    if ($stderr) {
        Write-Warn "scp stderr: $stderr"
    }

    return [PSCustomObject]@{
        ExitCode = $p.ExitCode
        StdOut   = $stdout
        StdErr   = $stderr
    }
}

#endregion

#region Connection dialog & profile picker

function Show-ConnectionDialog {
    param(
        [string]$Purpose
    )

    Write-Info "Opening connection dialog for '$Purpose'."

    $win = New-Window -Title "Server Connection - $Purpose" -Width 480 -Height 380
    $win.SizeToContent = 'Height'
    $win.MaxHeight = 700
    $win.ResizeMode = 'CanResize'

    $root = New-StackPanel -Orientation Vertical -Margin 16

    $title = New-Label -Text $Purpose -FontSize 16 -Bold $true
    $title.Foreground = $Color_Accent
    $root.Children.Add($title) | Out-Null

    $hint = New-Label -Text "Enter your server connection details. Fields marked * are required."
    $hint.Foreground = $Color_Text
    $root.Children.Add($hint) | Out-Null

    $grid = New-Object Windows.Controls.Grid
    $grid.Margin = "0,8,0,0"
    for ($i=0; $i -lt 2; $i++) {
        $col = New-Object Windows.Controls.ColumnDefinition
        if ($i -eq 0) { $col.Width = "Auto" } else { $col.Width = "*" }
        $grid.ColumnDefinitions.Add($col)
    }

    function Add-GridRow([string]$labelText, $control) {
        $rowIndex = $grid.RowDefinitions.Count
        $row = New-Object Windows.Controls.RowDefinition
        $row.Height = "Auto"
        $grid.RowDefinitions.Add($row)

        $lbl = New-Label -Text $labelText
        $lbl.Margin = "0,0,8,4"

        [Windows.Controls.Grid]::SetRow($lbl, $rowIndex)
        [Windows.Controls.Grid]::SetColumn($lbl, 0)
        [Windows.Controls.Grid]::SetRow($control, $rowIndex)
        [Windows.Controls.Grid]::SetColumn($control, 1)

        $grid.Children.Add($lbl) | Out-Null
        $grid.Children.Add($control) | Out-Null
    }

    $tbHost = New-TextBox
    Add-GridRow "Server IP or Hostname *" $tbHost

    $tbUser = New-TextBox
    Add-GridRow "SSH Username *" $tbUser

    $tbPort = New-TextBox -Text "22"
    Add-GridRow "SSH Port" $tbPort

    $spIdent = New-Object Windows.Controls.StackPanel
    $spIdent.Orientation = "Horizontal"
    $spIdent.Margin = "0,0,0,8"

    $tbIdent = New-TextBox
    $tbIdent.Width = 260
    $btnBrowse = New-Button -Content "Browse..."
    $btnBrowse.Width = 100
    $btnBrowse.Margin = "8,0,0,0"
    $btnBrowse.Add_Click({
        $path = Select-SSHKeyFile
        if ($path) { $tbIdent.Text = $path }
    })
    $spIdent.Children.Add($tbIdent) | Out-Null
    $spIdent.Children.Add($btnBrowse) | Out-Null

    Add-GridRow "SSH Private Key (optional)" $spIdent

    $tbPassword = New-TextBox -IsPassword $true
    Add-GridRow "Password (optional)" $tbPassword

    $cbSaveProfile = New-Object Windows.Controls.CheckBox
    $cbSaveProfile.Content = "Save as connection profile"
    $cbSaveProfile.Margin = "0,4,0,4"
    $cbSaveProfile.Foreground = $Color_Text

    $tbProfileName = New-TextBox
    $tbProfileName.IsEnabled = $false

    $cbSaveProfile.Add_Checked({ $tbProfileName.IsEnabled = $true })
    $cbSaveProfile.Add_Unchecked({ $tbProfileName.IsEnabled = $false })

    Add-GridRow "Save Profile" $cbSaveProfile
    Add-GridRow "Profile Name" $tbProfileName

    $root.Children.Add($grid) | Out-Null

    # Buttons
    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Center"
    $btnPanel.Margin = "0,16,0,0"

    $btnCancel = New-Button -Content "Cancel"
    $btnCancel.Width = 100
    $btnCancel.Margin = "0,0,12,0"
    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

    $btnOK = New-Button -Content "Connect"
    $btnOK.Width = 120
    $btnOK.Add_Click({
        if (-not $tbHost.Text.Trim()) {
            Show-MessageBoxWarn "Please enter a server hostname or IP." "Validation"
            return
        }
        if (-not $tbUser.Text.Trim()) {
            Show-MessageBoxWarn "Please enter an SSH username." "Validation"
            return
        }
        if ($cbSaveProfile.IsChecked -and -not $tbProfileName.Text.Trim()) {
            Show-MessageBoxWarn "Please enter a profile name or uncheck 'Save as connection profile'." "Validation"
            return
        }
        $win.DialogResult = $true
        $win.Close()
    })

    $btnPanel.Children.Add($btnCancel) | Out-Null
    $btnPanel.Children.Add($btnOK) | Out-Null
    $root.Children.Add($btnPanel) | Out-Null

    $scroll = New-Object Windows.Controls.ScrollViewer
    $scroll.VerticalScrollBarVisibility = 'Auto'
    $scroll.Content = $root

    $win.Content = $scroll
    $null = $win.ShowDialog()

    if (-not $win.DialogResult) {
        Write-Info "Connection dialog cancelled by user."
        return $null
    }

    $conn = [PSCustomObject]@{
        Name         = if ($cbSaveProfile.IsChecked) { $tbProfileName.Text.Trim() } else { "" }
        Host         = $tbHost.Text.Trim()
        User         = $tbUser.Text.Trim()
        Port         = (if ($tbPort.Text.Trim()) { $tbPort.Text.Trim() } else { "22" })
        IdentityFile = $tbIdent.Text.Trim()
        Password     = $tbPassword.Password
    }

    Write-Info "Connection dialog collected: Host=$($conn.Host), User=$($conn.User), Port=$($conn.Port), IdentityFile=$($conn.IdentityFile), SaveProfile=$($cbSaveProfile.IsChecked)"

    return $conn
}

function Select-ProfileFromDisk {
    param([string]$Purpose)

    $profiles = Get-StoredProfiles
    if (-not $profiles -or $profiles.Count -eq 0) {
        Show-MessageBoxWarn "No stored connection profiles were found in $($Script:ProfilesDir)." "Profiles"
        return $null
    }

    $win = New-Window -Title "Select Profile - $Purpose" -Width 520 -Height 420
    $win.SizeToContent = 'Height'
    $win.MaxHeight = 700
    $win.ResizeMode = 'CanResize'

    $root = New-StackPanel -Orientation Vertical -Margin 16

    $title = New-Label -Text "$Purpose - Select a stored connection profile:" -FontSize 16 -Bold $true
    $title.Foreground = $Color_Accent
    $root.Children.Add($title) | Out-Null

    $lb = New-Object Windows.Controls.ListBox
    $lb.Margin = "0,8,0,8"
    $lb.Background = $Color_Panel
    $lb.Foreground = $Color_Text
    $lb.MaxHeight = 260
    [Windows.Controls.ScrollViewer]::SetVerticalScrollBarVisibility($lb, 'Auto')

    foreach ($p in $profiles) {
        $item = New-Object Windows.Controls.ListBoxItem
        $item.Content = "{0} ({1}@{2}:{3})" -f $p.Name, $p.User, $p.Host, $p.Port
        $item.Tag     = $p
        $lb.Items.Add($item) | Out-Null
    }

    $root.Children.Add($lb) | Out-Null

    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Center"
    $btnPanel.Margin = "0,16,0,0"

    $btnCancel = New-Button -Content "Cancel"
    $btnCancel.Width = 120
    $btnCancel.Margin = "0,0,16,0"
    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

    $btnOK = New-Button -Content "Use Profile"
    $btnOK.Width = 120
    $btnOK.Margin = "0,0,0,0"
    $btnOK.Add_Click({
        if (-not $lb.SelectedItem) {
            Show-MessageBoxWarn "Please select a profile or click Cancel." "Profile Selection"
            return
        }
        $win.DialogResult = $true
        $win.Close()
    })

    $btnPanel.Children.Add($btnCancel) | Out-Null
    $btnPanel.Children.Add($btnOK) | Out-Null
    $root.Children.Add($btnPanel) | Out-Null

    $scroll = New-Object Windows.Controls.ScrollViewer
    $scroll.VerticalScrollBarVisibility = 'Auto'
    $scroll.Content = $root

    $win.Content = $scroll
    $null = $win.ShowDialog()

    if (-not $win.DialogResult) {
        Write-Info "Profile selection cancelled."
        return $null
    }

    $selectedItem = [Windows.Controls.ListBoxItem]$lb.SelectedItem
    $p = [ConnectionProfile]$selectedItem.Tag
    Write-Info "User selected profile: $($p.Name) ($($p.User)@$($p.Host):$($p.Port))"
    return $p
}

function Prompt-Connection {
    param(
        [string]$Purpose,
        [ref]$CachedConn
    )

    Write-Info "Prompt-Connection started for '$Purpose'. CachedConn set = $([bool]$CachedConn.Value)"

    # 1) Offer cached connection reuse
    if ($CachedConn.Value) {
        $c = $CachedConn.Value
        $msg = "Reuse $Purpose connection: $($c.User)@$($c.Host):$($c.Port)?"
        if (Show-Confirm $msg "Reuse Connection" $true) {
            Prepare-HostKey $c.Host
            return $c
        }
    }

    # 2) Show profile selector if any exist
    $profiles = Get-StoredProfiles
    if ($profiles -and $profiles.Count -gt 0) {
        Write-Info "Prompt-Connection: offering profile list for '$Purpose'."
        $p = Select-ProfileFromDisk -Purpose $Purpose
        if ($p) {
            $CachedConn.Value = $p
            Prepare-HostKey $p.Host
            return $p
        } else {
            Write-Info "No profile chosen, falling back to manual/GUI entry."
        }
    } else {
        Write-Info "Prompt-Connection: no stored profiles found; going straight to manual/GUI entry."
    }

    # 3) Manual GUI entry
    $conn = Show-ConnectionDialog -Purpose $Purpose
    if (-not $conn) {
        return $null
    }

    $obj = [PSCustomObject]@{
        Name         = $conn.Name
        Host         = $conn.Host
        User         = $conn.User
        Port         = $conn.Port
        IdentityFile = $conn.IdentityFile
    }

    if ($conn.PSObject.Properties.Match('Password').Count -gt 0 -and $conn.Password) {
        $obj | Add-Member -NotePropertyName Password -NotePropertyValue $conn.Password -Force
    }

    $CachedConn.Value = $obj
    Prepare-HostKey $obj.Host
    return $obj
}

#endregion

#region P12 helpers (.p12 password check via openssl)

function Test-P12Password {
    param(
        [string]$P12Path,
        [string]$Password
    )

    if (-not (Test-Path $P12Path)) {
        Write-ErrorLog ".p12 file not found: $P12Path"
        return $false
    }

    $cmd = "echo `"`" | openssl pkcs12 -in `"$P12Path`" -nokeys -passin pass:`"$Password`" -passout pass:`"dummy`""
    Write-Info "Testing P12 password via: $cmd"
    $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -PassThru -WindowStyle Hidden -Wait
    if ($proc.ExitCode -eq 0) { return $true }

    $cmd2 = "echo `"`" | openssl pkcs12 -in `"$P12Path`" -legacy -nokeys -passin pass:`"$Password`" -passout pass:`"dummy`""
    Write-Info "Testing P12 password (legacy) via: $cmd2"
    $proc2 = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd2" -PassThru -WindowStyle Hidden -Wait
    return ($proc2.ExitCode -eq 0)
}

function Harden-SSHRoot {
    param(
        [string]$HostName,
        [string]$Port = "22",
        [string]$User,
        [string]$IdentityFile
    )

    Write-Info ("Harden-SSHRoot invoked for {0}@{1}:{2}" -f $User, $HostName, $Port)

    $hardenScript = @'
set -e

SSHD_CONFIG="/etc/ssh/sshd_config"
TS=$(date +%Y%m%d%H%M%S)
BACKUP="/etc/ssh/sshd_config.stardust_backup_${TS}"

if [ ! -f "$SSHD_CONFIG" ]; then
  echo "sshd_config not found at $SSHD_CONFIG"
  exit 1
fi

cp "$SSHD_CONFIG" "$BACKUP"

if grep -qE '^[#[:space:]]*PermitRootLogin' "$SSHD_CONFIG"; then
  sed -i 's/^[#[:space:]]*PermitRootLogin.*/PermitRootLogin no/' "$SSHD_CONFIG"
else
  echo "" >> "$SSHD_CONFIG"
  echo "PermitRootLogin no" >> "$SSHD_CONFIG"
fi

if command -v systemctl >/dev/null 2>&1; then
  systemctl restart sshd 2>/dev/null || systemctl restart ssh 2>/dev/null
else
  service sshd restart 2>/dev/null || service ssh restart 2>/dev/null || /etc/init.d/ssh restart 2>/dev/null || true
fi

echo "Root SSH login disabled. Backup saved to $BACKUP"
'@

    $cmd = "sudo bash -s"
    $res = Invoke-SshCommand `
        -HostName    $HostName `
        -User        $User `
        -Port        $Port `
        -IdentityFile $IdentityFile `
        -Command     $cmd `
        -InputData   $hardenScript `
        -TimeoutMs   120000

    if ($res.ExitCode -ne 0) {
        $msg = "Failed to harden sshd on $HostName.`nExitCode: $($res.ExitCode)`nError: $($res.StdErr)"
        Write-ErrorLog $msg
        Show-MessageBoxError $msg "Disable Root SSH"
    } else {
        Write-Info "Harden-SSHRoot success output: $($res.StdOut)"
        Show-MessageBoxInfo "Root SSH login disabled on $HostName.`n`n$($res.StdOut)" "Disable Root SSH"
    }
}

function Run-NewServerSetup {
    Write-Info "Run-NewServerSetup invoked."

    $conn = Prompt-Connection -Purpose "New Server (initial root or admin login)" -CachedConn ([ref]$Script:NewServerConn)
    if (-not $conn) {
        Show-MessageBoxWarn "No connection selected; New Server Setup cancelled." "New Server Setup"
        return
    }

    Show-MessageBoxInfo "New Server Setup will create a non-root sudo user, copy SSH keys, test login, and optionally harden sshd." "New Server Setup"

    # Dialog for new user + password
    $win  = New-Window -Title "New User on Remote Server" -Width 420 -Height 260
    $root = New-StackPanel -Orientation Vertical -Margin 16
    $root.Children.Add((New-Label -Text "Create a new non-root sudo user on the remote server" -FontSize 14 -Bold $true)) | Out-Null

    $grid = New-Object Windows.Controls.Grid
    for ($i=0; $i -lt 2; $i++) {
        $col = New-Object Windows.Controls.ColumnDefinition
        if ($i -eq 0) { $col.Width = "Auto" } else { $col.Width = "*" }
        $grid.ColumnDefinitions.Add($col)
    }

    function Add-Row($labelText,$control) {
        $row = New-Object Windows.Controls.RowDefinition
        $row.Height = "Auto"
        $rowIndex = $grid.RowDefinitions.Count
        $grid.RowDefinitions.Add($row)

        $lbl = New-Label -Text $labelText
        $lbl.Margin = "0,0,8,4"

        [Windows.Controls.Grid]::SetRow($lbl, $rowIndex)
        [Windows.Controls.Grid]::SetColumn($lbl, 0)
        [Windows.Controls.Grid]::SetRow($control, $rowIndex)
        [Windows.Controls.Grid]::SetColumn($control, 1)

        $grid.Children.Add($lbl)    | Out-Null
        $grid.Children.Add($control)| Out-Null
    }

    $tbUser = New-TextBox -Text "nodeadmin"
    $pwd1   = New-TextBox -IsPassword $true
    $pwd2   = New-TextBox -IsPassword $true

    Add-Row "New username:" $tbUser
    Add-Row "Password:"     $pwd1
    Add-Row "Confirm:"      $pwd2

    $root.Children.Add($grid) | Out-Null

    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Center"
    $btnPanel.Margin = "0,20,0,0"

    $btnCancel = New-Button -Content "Cancel"
    $btnCancel.Width = 120
    $btnCancel.Margin = "0,0,16,0"
    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

    $btnOK = New-Button -Content "Create User"
    $btnOK.Width = 120
    $btnOK.Margin = "0,0,0,0"
    $btnOK.Add_Click({
        if (-not $tbUser.Text.Trim()) {
            Write-Warn "New user dialog validation failed: empty username."
            Show-MessageBoxWarn "Username cannot be empty." "Validation"
            return
        }
        if (-not $pwd1.Password) {
            Write-Warn "New user dialog validation failed: empty password."
            Show-MessageBoxWarn "Password cannot be empty." "Validation"
            return
        }
        if ($pwd1.Password -ne $pwd2.Password) {
            Write-Warn "New user dialog validation failed: password mismatch."
            Show-MessageBoxWarn "Passwords do not match." "Validation"
            return
        }
        Write-Info "New user dialog validation passed. Proceeding to create user."
        $win.DialogResult = $true
        $win.Close()
    })

    $btnPanel.Children.Add($btnCancel) | Out-Null
    $btnPanel.Children.Add($btnOK)     | Out-Null

    $root.Children.Add($btnPanel) | Out-Null
    $win.Content = $root
    $null = $win.ShowDialog()

    if (-not $win.DialogResult) {
        $userPreview = $tbUser.Text.Trim()
        Write-Info "User creation dialog ended without confirmation. DialogResult=$($win.DialogResult); EnteredUser='$userPreview'."
        Show-MessageBoxWarn "User creation was cancelled. No changes were made on the server." "New Server Setup"
        return
    }

    $newUser = $tbUser.Text.Trim()
    $pass    = $pwd1.Password
    Write-Info "Creating new user '$newUser' on $($conn.Host)."

    #
    # Build a reusable bash script for both root and non-root flows
    #
    $remoteScript = @'
set -e
NEWUSER='{0}'
NEWPASS='{1}'

echo "=== Bootstrap: creating user $NEWUSER with sudo access ==="

if id "$NEWUSER" >/dev/null 2>&1; then
  echo "User $NEWUSER already exists. Skipping creation."
  exit 0
fi

if command -v useradd >/dev/null 2>&1; then
  useradd -m -s /bin/bash "$NEWUSER" || echo "useradd returned non-zero (possibly user already exists). Continuing..."
elif command -v adduser >/dev/null 2>&1; then
  adduser --disabled-password --gecos "" "$NEWUSER" || echo "adduser returned non-zero (possibly user already exists). Continuing..."
else
  echo "Neither useradd nor adduser is available on this system."
  exit 1
fi

echo "$NEWUSER:$NEWPASS" | chpasswd

if getent group sudo >/dev/null 2>&1; then
  usermod -aG sudo "$NEWUSER"
elif getent group wheel >/dev/null 2>&1; then
  usermod -aG wheel "$NEWUSER"
fi

SRC_KEYS=""
if [ -f "/root/.ssh/authorized_keys" ]; then
  SRC_KEYS="/root/.ssh/authorized_keys"
elif [ -n "$SUDO_USER" ] && [ -f "/home/$SUDO_USER/.ssh/authorized_keys" ]; then
  SRC_KEYS="/home/$SUDO_USER/.ssh/authorized_keys"
fi

if [ -n "$SRC_KEYS" ]; then
  HOME_DIR=$(eval echo "~$NEWUSER")
  mkdir -p "$HOME_DIR/.ssh"
  cp "$SRC_KEYS" "$HOME_DIR/.ssh/authorized_keys"
  chown -R "$NEWUSER:$NEWUSER" "$HOME_DIR/.ssh"
  chmod 700 "$HOME_DIR/.ssh"
  chmod 600 "$HOME_DIR/.ssh/authorized_keys"
  echo "Copied $SRC_KEYS to $HOME_DIR/.ssh/authorized_keys"
else
  echo "No existing authorized_keys found to copy. You can add keys later."
fi

echo "Bootstrap complete for $NEWUSER."
'@ -f $newUser, $pass

    #
    # ROOT FLOW: open a new cmd.exe window that pipes the script into ssh.exe.
    # User types the root/key passphrase in that window; after auth, the script runs automatically.
    #
    if ($conn.User -eq "root") {
        Write-Warn "New Server Setup: detected root login. Using piped ssh.exe flow for '$newUser'."

        $tmpDir    = [System.IO.Path]::GetTempPath()
        $tmpScript = Join-Path $tmpDir ("server-toolkit-bootstrap-{0}-{1}.sh" -f $conn.Host, $newUser)
        Write-Info "Writing bootstrap script for '$newUser' to $tmpScript"
        Set-Content -LiteralPath $tmpScript -Value $remoteScript -Encoding ASCII

        $sshArgs = @()
        if ($conn.IdentityFile) {
            $sshArgs += "-i `"$($conn.IdentityFile)`""
        }
        $sshArgs += "-p $($conn.Port)"
        $sshArgs += "-T"
        $sshArgs += "-o StrictHostKeyChecking=no"
        $sshArgs += "-o UserKnownHostsFile=`"$($Script:KnownHosts)`""
        $sshArgs += "$($conn.User)@$($conn.Host)"

        $sshCommand = "ssh.exe " + ($sshArgs -join " ")
        $cmdLine    = "type `"$tmpScript`" | $sshCommand"

        Write-Info "Launching bootstrap console with: $cmdLine"
        Start-Process -FilePath "cmd.exe" -ArgumentList "/k $cmdLine" | Out-Null
        return
    }

    #
    # NON-ROOT FLOW: run fully scripted user creation via Invoke-SshCommand (key-based ssh)
    #
    $cmd = "bash -s"
    $res = Invoke-SshCommand `
        -HostName    $conn.Host `
        -User        $conn.User `
        -Port        $conn.Port `
        -IdentityFile $conn.IdentityFile `
        -Command     $cmd `
        -InputData   $remoteScript `
        -TimeoutMs   120000

    if ($res.ExitCode -ne 0) {
        $msg = "Remote user creation script failed.`nExitCode: $($res.ExitCode)`nError: $($res.StdErr)"
        Write-ErrorLog "New Server Setup: $msg"
        Show-MessageBoxError $msg "New Server Setup"
        return
    }

    Write-Info "Remote user script output: $($res.StdOut)"
    Show-MessageBoxInfo "New user '$newUser' created on $($conn.Host)." "New Server Setup"

    if (Show-Confirm "Do you want to save an SSH profile for '$newUser' now?" "New User Profile" $true) {
        $pname = "$($conn.Host)-$newUser"
        $profile = [ConnectionProfile]::new($pname,$conn.Host,$newUser,$conn.Port,$conn.IdentityFile)
        Save-SSHProfile -Profile $profile
        if (Show-Confirm "Do you want to launch an SSH session for '$newUser' now?" "Launch SSH" $false) {
            $sshArgs = @()
            if ($conn.IdentityFile) { $sshArgs += "-i `"$($conn.IdentityFile)`"" }
            $sshArgs += "-p $($conn.Port)"
            $sshArgs += "$newUser@$($conn.Host)"

            $cmdLine = "ssh.exe " + ($sshArgs -join " ")
            Write-Info "Launching interactive ssh session for new user in new console: $cmdLine"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/k $cmdLine" | Out-Null
        }
    }

    if (Show-Confirm "Do you want to disable SSH root login on $($conn.Host) now?`n`nThis will use 'sudo' as '$newUser' to update sshd_config." "Disable Root SSH" $false) {
        Harden-SSHRoot -HostName $conn.Host -Port $conn.Port -User $newUser -IdentityFile $conn.IdentityFile
    }
}

function Run-CreateNonRootUser {
    Write-Info "Run-CreateNonRootUser invoked."
    Run-NewServerSetup
}

function Run-BackupP12 {
    Write-Info "Run-BackupP12 invoked."

    $conn = Prompt-Connection -Purpose "Old Server (P12 backup source)" -CachedConn ([ref]$Script:OldServerConn)
    if (-not $conn) {
        Show-MessageBoxWarn "No connection selected; backup cancelled." "Backup P12"
        return
    }

    $remoteFind = "cd ~; find . -maxdepth 5 -type f -name '*.p12' ! -path '*\/hash\/*' ! -path '*\/ordinal\/*'"
    $res = Invoke-SshCommand -HostName $conn.Host -User $conn.User -Port $conn.Port -IdentityFile $conn.IdentityFile -Command $remoteFind
    if ($res.ExitCode -ne 0 -or -not $res.StdOut.Trim()) {
        Show-MessageBoxWarn "No .p12 files found on the remote server." "Backup P12"
        return
    }

    $paths = $res.StdOut.Trim().Split("`n") | Where-Object { $_ -and $_ -ne "." } | Sort-Object
    Write-Info "Remote .p12 files: $($paths -join ', ')"

    # simple list dialog
    $win = New-Window -Title "Select Remote .p12 File" -Width 500 -Height 360
    $root = New-StackPanel -Orientation Vertical -Margin 16
    $root.Children.Add((New-Label -Text "Select a .p12 file to download from $($conn.Host)" -FontSize 14 -Bold $true)) | Out-Null
    $lb = New-Object Windows.Controls.ListBox
    $lb.Margin = "0,8,0,8"
    $lb.Background = $Color_Panel
    $lb.Foreground = $Color_Text
    foreach ($p in $paths) {
        $lb.Items.Add($p) | Out-Null
    }
    $root.Children.Add($lb) | Out-Null

    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Right"

    $btnCancel = New-Button -Content "Cancel"
    $btnCancel.Width = 100
    $btnCancel.Margin = "0,0,12,0"
    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

    $btnOK = New-Button -Content "Download"
    $btnOK.Width = 120
    $btnOK.Add_Click({
        if (-not $lb.SelectedItem) {
            Show-MessageBoxWarn "Please select a .p12 file." "Backup P12"
            return
        }
        $win.DialogResult = $true
        $win.Close()
    })

    $btnPanel.Children.Add($btnCancel) | Out-Null
    $btnPanel.Children.Add($btnOK) | Out-Null
    $root.Children.Add($btnPanel) | Out-Null

    $win.Content = $root
    $null = $win.ShowDialog()

    if (-not $win.DialogResult) {
        Write-Info "User cancelled .p12 download selection."
        return
    }

    $selectedPath = [string]$lb.SelectedItem
    $localDir = Join-Path $Script:HomeDir "Downloads"
    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir -Force | Out-Null
    }
    $localPath = Join-Path $localDir ([System.IO.Path]::GetFileName($selectedPath))

    $res2 = Invoke-ScpDownload -HostName $conn.Host -User $conn.User -Port $conn.Port -IdentityFile $conn.IdentityFile -RemotePath $selectedPath -LocalPath $localPath
    if ($res2.ExitCode -ne 0) {
        Show-MessageBoxError "Failed to download .p12 file:`n$res2.StdErr" "Backup P12"
        return
    }

    Show-MessageBoxInfo "P12 file downloaded to: $localPath" "Backup P12"
}

function Run-UploadP12 {
    Write-Info "Run-UploadP12 invoked."

    $p12 = Select-P12File
    if (-not $p12) { return }

    # ask for password & verify
    $win = New-Window -Title "Verify P12 Password" -Width 420 -Height 220
    $root = New-StackPanel -Orientation Vertical -Margin 16
    $root.Children.Add((New-Label -Text "Enter the password for the selected .p12 file:" -FontSize 13 -Bold $true)) | Out-Null
    $pwdBox = New-TextBox -IsPassword $true
    $root.Children.Add($pwdBox) | Out-Null

    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Right"
    $btnCancel = New-Button -Content "Cancel"
    $btnCancel.Width = 100
    $btnCancel.Margin = "0,0,12,0"
    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })
    $btnOK = New-Button -Content "Verify"
    $btnOK.Width = 120
    $btnOK.Add_Click({
        if (-not $pwdBox.Password) {
            Show-MessageBoxWarn "Password cannot be empty." "P12 Password"
            return
        }
        if (-not (Test-P12Password -P12Path $p12 -Password $pwdBox.Password)) {
            Show-MessageBoxError "Incorrect P12 password." "P12 Password"
            return
        }
        $win.DialogResult = $true
        $win.Close()
    })
    $btnPanel.Children.Add($btnCancel) | Out-Null
    $btnPanel.Children.Add($btnOK) | Out-Null
    $root.Children.Add($btnPanel) | Out-Null
    $win.Content = $root
    $null = $win.ShowDialog()

    if (-not $win.DialogResult) {
        Write-Info "User cancelled P12 password verification."
        return
    }

    Show-MessageBoxInfo "P12 password verified. Next select the target server." "Upload P12"

    $conn = Prompt-Connection -Purpose "New Server (P12 upload target)" -CachedConn ([ref]$Script:NewServerConn)
    if (-not $conn) {
        Show-MessageBoxWarn "No connection selected; upload cancelled." "Upload P12"
        return
    }

    $res = Invoke-ScpUpload -HostName $conn.Host -User $conn.User -Port $conn.Port -IdentityFile $conn.IdentityFile -LocalPath $p12 -RemotePath "~/"
    if ($res.ExitCode -ne 0) {
        Show-MessageBoxError "Failed to upload P12 file:`n$res.StdErr" "Upload P12"
        return
    }

    Show-MessageBoxInfo "P12 file uploaded to remote home directory (~)." "Upload P12"
}

function Run-ExportSSHProfile {
    Write-Info "Run-ExportSSHProfile invoked."

    $win = New-Window -Title "Export SSH Profile" -Width 420 -Height 320
    $root = New-StackPanel -Orientation Vertical -Margin 16

    $root.Children.Add((New-Label -Text "Create a new SSH profile file in $($Script:ProfilesDir)" -FontSize 14 -Bold $true)) | Out-Null

    $grid = New-Object Windows.Controls.Grid
    for ($i=0; $i -lt 2; $i++) {
        $col = New-Object Windows.Controls.ColumnDefinition
        if ($i -eq 0) { $col.Width = "Auto" } else { $col.Width = "*" }
        $grid.ColumnDefinitions.Add($col)
    }

    function Add-Row([string]$labelText,$control) {
        $row = New-Object Windows.Controls.RowDefinition
        $row.Height = "Auto"
        $rowIndex = $grid.RowDefinitions.Count
        $grid.RowDefinitions.Add($row)

        $lbl = New-Label -Text $labelText
        $lbl.Margin = "0,0,8,4"

        [Windows.Controls.Grid]::SetRow($lbl, $rowIndex)
        [Windows.Controls.Grid]::SetColumn($lbl, 0)
        [Windows.Controls.Grid]::SetRow($control, $rowIndex)
        [Windows.Controls.Grid]::SetColumn($control, 1)

        $grid.Children.Add($lbl) | Out-Null
        $grid.Children.Add($control) | Out-Null
    }

    $tbName  = New-TextBox
    $tbHost  = New-TextBox
    $tbUser  = New-TextBox
    $tbPort  = New-TextBox -Text "22"

    $spIdent = New-Object Windows.Controls.StackPanel
    $spIdent.Orientation = "Horizontal"
    $tbIdent = New-TextBox
    $tbIdent.Width = 230
    $btnBrowse = New-Button -Content "Browse..."
    $btnBrowse.Width = 100
    $btnBrowse.Margin = "8,0,0,0"
    $btnBrowse.Add_Click({
        $path = Select-SSHKeyFile
        if ($path) { $tbIdent.Text = $path }
    })
    $spIdent.Children.Add($tbIdent) | Out-Null
    $spIdent.Children.Add($btnBrowse) | Out-Null

    Add-Row "Profile Name:" $tbName
    Add-Row "Server IP/Host:" $tbHost
    Add-Row "Username:" $tbUser
    Add-Row "Port:" $tbPort
    Add-Row "SSH Key (optional):" $spIdent

    $root.Children.Add($grid) | Out-Null

    $btnPanel = New-Object Windows.Controls.StackPanel
    $btnPanel.Orientation = "Horizontal"
    $btnPanel.HorizontalAlignment = "Center"
    $btnPanel.Margin = "0,16,0,0"

    $btnCancel = New-Button -Content "Cancel"
    $btnCancel.Width = 100
    $btnCancel.Margin = "0,0,12,0"
    $btnCancel.Add_Click({ $win.DialogResult = $false; $win.Close() })

    $btnOK = New-Button -Content "Save Profile"
    $btnOK.Width = 120
    $btnOK.Add_Click({
        if (-not $tbName.Text.Trim()) {
            Show-MessageBoxWarn "Profile name is required." "Export Profile"
            return
        }
        if (-not $tbHost.Text.Trim() -or -not $tbUser.Text.Trim()) {
            Show-MessageBoxWarn "Host and Username are required." "Export Profile"
            return
        }
        $win.DialogResult = $true
        $win.Close()
    })

    $btnPanel.Children.Add($btnCancel) | Out-Null
    $btnPanel.Children.Add($btnOK) | Out-Null
    $root.Children.Add($btnPanel) | Out-Null
    $win.Content = $root
    $null = $win.ShowDialog()

    if (-not $win.DialogResult) {
        Write-Info "Export SSH Profile cancelled."
        return
    }

    $profile = [ConnectionProfile]::new(
        $tbName.Text.Trim(),
        $tbHost.Text.Trim(),
        $tbUser.Text.Trim(),
        (if ($tbPort.Text.Trim()) { $tbPort.Text.Trim() } else { "22" }),
        $tbIdent.Text.Trim()
    )
    Save-SSHProfile -Profile $profile
    Show-MessageBoxInfo "Profile saved to: $(Get-ProfilePath -Name $profile.Name)" "Export SSH Profile"
}

function Run-ServerLogin {
    Write-Info "Run-ServerLogin invoked."

    $conn = Prompt-Connection -Purpose "Server Login" -CachedConn ([ref]$Script:NewServerConn)
    if (-not $conn) {
        Show-MessageBoxWarn "No connection selected; login cancelled." "Server Login"
        return
    }

    $sshArgs = @()
    if ($conn.IdentityFile) { $sshArgs += "-i `"$($conn.IdentityFile)`"" }
    $sshArgs += "-p $($conn.Port)"
    $sshArgs += "$($conn.User)@$($conn.Host)"

    $cmdLine = "ssh.exe " + ($sshArgs -join " ")
    Write-Info "Launching interactive ssh session in new console: $cmdLine"

    # Open a new cmd window and keep it open after ssh exits
    Start-Process -FilePath "cmd.exe" -ArgumentList "/k $cmdLine" | Out-Null
}

#endregion

#region Main GUI

function Show-MainWindow {
    Write-Info "Launching Server Toolkit main window."

    $win = New-Window -Title "Stardust Collective - Server Toolkit (Windows)" -Width 520 -Height 420
    $root = New-StackPanel -Orientation Vertical -Margin 16

    $title = New-Label -Text "Stardust Collective - Server Toolkit" -FontSize 18 -Bold $true
    $title.Foreground = $Color_Accent
    $root.Children.Add($title) | Out-Null

    $subtitle = New-Label -Text "Node server preperation - Brought to you by @Proph151Music"
    $root.Children.Add($subtitle) | Out-Null

    $btnNewServer     = New-Button -Content "New Server Setup"
    $btnCreateUser    = New-Button -Content "Create Non-Root User"
    $btnBackupP12     = New-Button -Content "Backup P12 File (From old server)"
    $btnUploadP12     = New-Button -Content "Upload P12 File"
    $btnExportProfile = New-Button -Content "Export SSH-Config File"
    $btnLogin         = New-Button -Content "Server Login (interactive shell)"
    $btnExit          = New-Button -Content "Exit"

    $btnNewServer.Add_Click({ Run-NewServerSetup })
    $btnCreateUser.Add_Click({ Run-CreateNonRootUser })
    $btnBackupP12.Add_Click({ Run-BackupP12 })
    $btnUploadP12.Add_Click({ Run-UploadP12 })
    $btnExportProfile.Add_Click({ Run-ExportSSHProfile })
    $btnLogin.Add_Click({ Run-ServerLogin })
    $btnExit.Add_Click({ $win.Close() })

    $root.Children.Add($btnNewServer)     | Out-Null
    $root.Children.Add($btnCreateUser)    | Out-Null
    $root.Children.Add($btnBackupP12)     | Out-Null
    $root.Children.Add($btnUploadP12)     | Out-Null
    $root.Children.Add($btnExportProfile) | Out-Null
    $root.Children.Add($btnLogin)         | Out-Null

    $root.Children.Add((New-Label -Text "")) | Out-Null
    $root.Children.Add($btnExit) | Out-Null

    $win.Content = $root
    $null = $win.ShowDialog()
    Write-Info "Main window closed."
}

#endregion

# Entry point
Write-Info "==== Server Toolkit started ===="
Show-MainWindow
Write-Info "==== Server Toolkit ended ===="
