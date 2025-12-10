<# 
    server-toolkit.ps1
    Stardust Collective - Server Setup Toolkit (Windows Edition, GUI)
#>
# Load required .NET assemblies for WPF
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Xaml

# Define the XAML for the WPF UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Server Toolkit" Height="500" Width="600" WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    <DockPanel LastChildFill="True">
        <!-- Menu bar -->
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem Header="_Load Profile..." x:Name="MenuLoadProfile"/>
                <MenuItem Header="_Save Profile..." x:Name="MenuSaveProfile"/>
                <Separator/>
                <MenuItem Header="E_xit" x:Name="MenuExit"/>
            </MenuItem>
        </Menu>
        <!-- Main content area -->
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>  <!-- Connection form -->
                <RowDefinition Height="Auto"/>  <!-- New user form -->
                <RowDefinition Height="Auto"/>  <!-- Options -->
                <RowDefinition Height="Auto"/>  <!-- Action button -->
                <RowDefinition Height="*"/>    <!-- Log output -->
            </Grid.RowDefinitions>
            <!-- Connection Details Group -->
            <GroupBox Header="Connection" Grid.Row="0" Margin="0,0,0,10">
                <Grid Margin="10,5,10,5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <!-- Host -->
                    <Label Content="Host:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="HostBox" Grid.Row="0" Grid.Column="1" Margin="5,2" />
                    <!-- Port -->
                    <Label Content="Port:" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="PortBox" Grid.Row="1" Grid.Column="1" Margin="5,2" Width="60" Text="22"/>
                    <!-- Username -->
                    <Label Content="Username:" Grid.Row="2" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="UserBox" Grid.Row="2" Grid.Column="1" Margin="5,2" />
                    <!-- Password -->
                    <Label Content="Password:" Grid.Row="3" Grid.Column="0" VerticalAlignment="Center"/>
                    <PasswordBox x:Name="PasswordBox" Grid.Row="3" Grid.Column="1" Margin="5,2" />
                    <!-- SSH Key -->
                    <Label Content="SSH Key File:" Grid.Row="4" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="KeyBox" Grid.Row="4" Grid.Column="1" Margin="5,2" IsReadOnly="True"/>
                    <Button x:Name="BrowseKeyButton" Content="Browse..." Grid.Row="4" Grid.Column="2" Margin="5,2,0,2" Padding="8,0" />
                </Grid>
            </GroupBox>
            <!-- New User Group -->
            <GroupBox Header="New User" Grid.Row="1" Margin="0,0,0,10">
                <Grid Margin="10,5,10,5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Label Content="New Username:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="NewUserBox" Grid.Row="0" Grid.Column="1" Margin="5,2" />
                    <Label Content="New User Password:" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center"/>
                    <PasswordBox x:Name="NewPasswordBox" Grid.Row="1" Grid.Column="1" Margin="5,2" />
                </Grid>
            </GroupBox>
            <!-- Options -->
            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="5,0,0,10">
                <CheckBox x:Name="DisableRootCheck" Content="Disable root login after setup" IsChecked="True"/>
            </StackPanel>
            <!-- Start Button -->
            <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,0,0,10">
                <Button x:Name="StartButton" Content="Start Setup" Width="100" HorizontalAlignment="Center"/>
            </StackPanel>
            <!-- Log Output -->
            <TextBox x:Name="LogBox" Grid.Row="4" Margin="0" Background="#FF000000" Foreground="#FF00FF00"
                     FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" AcceptsReturn="True"/>
            <!-- Overlays -->
            <!-- Sudo Password Prompt Overlay -->
            <Grid x:Name="SudoOverlay" Visibility="Collapsed" Background="#80000000" VerticalAlignment="Stretch" HorizontalAlignment="Stretch"
                  Grid.RowSpan="5">
                <Border Background="White" Padding="20" CornerRadius="5" HorizontalAlignment="Center" VerticalAlignment="Center">
                    <StackPanel>
                        <TextBlock Text="Elevation required: enter sudo password" Margin="0,0,0,10" />
                        <PasswordBox x:Name="SudoPassBox" Width="200" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="SudoOkButton" Content="OK" Width="60" Margin="5"/>
                            <Button x:Name="SudoCancelButton" Content="Cancel" Width="60" Margin="5"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
            </Grid>
            <!-- Requiretty Confirmation Overlay -->
            <Grid x:Name="RequireOverlay" Visibility="Collapsed" Background="#80000000" VerticalAlignment="Stretch" HorizontalAlignment="Stretch"
                  Grid.RowSpan="5">
                <Border Background="White" Padding="20" CornerRadius="5" HorizontalAlignment="Center" VerticalAlignment="Center" Width="300">
                    <StackPanel>
                        <TextBlock Text="The server requires a TTY for sudo." TextWrapping="Wrap" Margin="0,0,0,10"/>
                        <TextBlock Text="Disable 'requiretty' in sudoers to continue?" TextWrapping="Wrap" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,5,0,0">
                            <Button x:Name="RequireYesButton" Content="Yes" Width="60" Margin="5"/>
                            <Button x:Name="RequireNoButton" Content="No" Width="60" Margin="5"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
            </Grid>
        </Grid>
    </DockPanel>
</Window>
"@

# Load the XAML into a Window object
$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get references to named controls
$HostBox        = $Window.FindName("HostBox")
$PortBox        = $Window.FindName("PortBox")
$UserBox        = $Window.FindName("UserBox")
$PasswordBox    = $Window.FindName("PasswordBox")
$KeyBox         = $Window.FindName("KeyBox")
$BrowseKeyButton= $Window.FindName("BrowseKeyButton")
$NewUserBox     = $Window.FindName("NewUserBox")
$NewPasswordBox = $Window.FindName("NewPasswordBox")
$DisableRootCheck = $Window.FindName("DisableRootCheck")
$StartButton    = $Window.FindName("StartButton")
$LogBox         = $Window.FindName("LogBox")
$SudoOverlay    = $Window.FindName("SudoOverlay")
$SudoPassBox    = $Window.FindName("SudoPassBox")
$SudoOkButton   = $Window.FindName("SudoOkButton")
$SudoCancelButton = $Window.FindName("SudoCancelButton")
$RequireOverlay = $Window.FindName("RequireOverlay")
$RequireYesButton = $Window.FindName("RequireYesButton")
$RequireNoButton  = $Window.FindName("RequireNoButton")
$MenuLoadProfile  = $Window.FindName("MenuLoadProfile")
$MenuSaveProfile  = $Window.FindName("MenuSaveProfile")
$MenuExit         = $Window.FindName("MenuExit")

# Helper function to append text to log in a thread-safe way
$AppendLog = [Action[string]]{
    param($text)
    # Ensure we append on the UI thread
    $LogBox.Dispatcher.Invoke([Action]{
        $LogBox.AppendText($text)
        $LogBox.ScrollToEnd()
    })
}

# Another helper: find plink and puttygen paths or download if needed
function Ensure-PuttyTools {
    # Find plink.exe
    $global:PlinkPath = ""
    try {
        $plinkCmd = Get-Command plink.exe -ErrorAction Stop
        $global:PlinkPath = $plinkCmd.Source
    } catch {
        # plink not found, attempt download
        $AppendLog.Invoke("[ERROR] Plink (plink.exe) not found. Attempting to download PuTTY...`r`n")
        # Determine architecture (use 64-bit if possible)
        $arch = if ([IntPtr]::Size -eq 8) {"w64"} else {"w32"}
        $plinkUrl = "https://the.earth.li/~sgtatham/putty/latest/$arch/plink.exe"
        $puttygenUrl = "https://the.earth.li/~sgtatham/putty/latest/$arch/puttygen.exe"
        # Download plink
        $tempPath = [IO.Path]::GetTempPath()
        $plinkTarget = Join-Path $tempPath "plink.exe"
        try {
            (New-Object Net.WebClient).DownloadFile($plinkUrl, $plinkTarget)
            $AppendLog.Invoke("Downloaded plink.exe to $plinkTarget`r`n")
            $global:PlinkPath = $plinkTarget
        } catch {
            $AppendLog.Invoke("[ERROR] Failed to download plink.exe. Please install PuTTY and try again.`r`n")
            return $false
        }
        # Also download puttygen for convenience (optional)
        $puttygenTarget = Join-Path $tempPath "puttygen.exe"
        try {
            (New-Object Net.WebClient).DownloadFile($puttygenUrl, $puttygenTarget)
            $AppendLog.Invoke("Downloaded puttygen.exe to $puttygenTarget`r`n")
        } catch {
            # If puttygen fails, not fatal; user will be prompted later if needed
            $AppendLog.Invoke("[WARN] Could not download puttygen.exe. Key conversion might prompt for manual steps.`r`n")
        }
    }
    return $true
}

# Ensure Plink is available up-front
if (-not (Ensure-PuttyTools)) {
    # If we cannot proceed due to missing plink, disable start and log error
    $AppendLog.Invoke("[FATAL] plink.exe is required. Setup cannot continue.`r`n")
    $StartButton.IsEnabled = $false
}

# Browse for key file
$BrowseKeyButton.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = "Select Private Key File"
    $dlg.Filter = "Private Key Files (*.ppk;*.pem;*.key)|*.ppk;*.pem;*.key|All Files|*.*"
    if ($dlg.ShowDialog()) {
        $KeyBox.Text = $dlg.FileName
    }
})

# Load Profile
$MenuLoadProfile.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = "Load SSH Profile"
    $dlg.Filter = "SSH Profile (*.ssh_config.txt)|*_ssh_config.txt;*.ssh_config.txt|All Files|*.*"
    $homeSSH = [IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), ".ssh")
    if ([IO.Directory]::Exists($homeSSH)) { $dlg.InitialDirectory = $homeSSH }
    if ($dlg.ShowDialog()) {
        try {
            $lines = Get-Content $dlg.FileName
            # Parse profile lines
            foreach ($line in $lines) {
                if ($line.Trim().StartsWith("#") -or [string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match "HostName\s+(.+)$") {
                    $HostBox.Text = $matches[1].Trim()
                }
                elseif ($line -match "User\s+(.+)$") {
                    $UserBox.Text = $matches[1].Trim()
                }
                elseif ($line -match "Port\s+(\d+)$") {
                    $PortBox.Text = $matches[1].Trim()
                }
                elseif ($line -match "IdentityFile\s+(.+)$") {
                    $KeyPath = $matches[1].Trim()
                    # If identity path contains ~, expand it
                    if ($KeyPath.StartsWith("~")) {
                        $KeyPath = $KeyPath -replace "^~", [Environment]::GetFolderPath('UserProfile')
                    }
                    $KeyBox.Text = $KeyPath
                }
            }
            $AppendLog.Invoke("Profile loaded: $($dlg.FileName)`r`n")
        } catch {
            $AppendLog.Invoke("[ERROR] Failed to load profile: $_`r`n")
        }
    }
})

# Save Profile
$MenuSaveProfile.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title = "Save SSH Profile"
    $dlg.Filter = "SSH Profile (*.ssh_config.txt)|*.ssh_config.txt|All Files|*.*"
    $homeSSH = [IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), ".ssh")
    if (-not [IO.Directory]::Exists($homeSSH)) { [IO.Directory]::CreateDirectory($homeSSH) | Out-Null }
    $dlg.InitialDirectory = $homeSSH
    # Suggest file name based on host or profile name
    $baseName = $HostBox.Text
    if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = "profile" }
    $baseName = $baseName.Replace(":", "_")  # remove problematic chars
    $dlg.FileName = "${baseName}_ssh_config.txt"
    if ($dlg.ShowDialog()) {
        try {
            $profilePath = $dlg.FileName
            # Determine Host alias for profile (use filename without _ssh_config)
            $alias = [IO.Path]::GetFileNameWithoutExtension($profilePath)
            $alias = $alias -replace "_ssh_config$", ""
            $content = @()
            $content += "Host $alias"
            $content += "    HostName $($HostBox.Text)"
            $content += "    User $($UserBox.Text)"
            $content += "    Port $($PortBox.Text)"
            if (-not [string]::IsNullOrWhiteSpace($KeyBox.Text)) {
                $content += "    IdentityFile $($KeyBox.Text)"
            }
            # Add helpful comment lines
            $content += ""
            $content += "# CreatedOn: $(Get-Date -Format yyyy-MM-dd)"
            $content += "# SSH login: ssh $([string]::IsNullOrEmpty($KeyBox.Text) ? "" : "-i $($KeyBox.Text) ")$($UserBox.Text)@$($HostBox.Text)"
            $content += "# SFTP: sftp $([string]::IsNullOrEmpty($KeyBox.Text) ? "" : "-i $($KeyBox.Text) ")$($UserBox.Text)@$($HostBox.Text)"
            $content | Out-File -FilePath $profilePath -Encoding ASCII
            $AppendLog.Invoke("Profile saved to: $profilePath`r`n")
        } catch {
            $AppendLog.Invoke("[ERROR] Failed to save profile: $_`r`n")
        }
    }
})

# Exit menu
$MenuExit.Add_Click({
    $Window.Close()
})

# Core logic for starting the setup
$StartButton.Add_Click({
    # Clear log for a new run
    $LogBox.Clear()
    # Gather inputs
    $host = $HostBox.Text.Trim()
    $port = $PortBox.Text.Trim()
    $user = $UserBox.Text.Trim()
    $pass = $PasswordBox.Password      # login password
    $keyPath = $KeyBox.Text.Trim()
    $newUser = $NewUserBox.Text.Trim()
    $newPass = $NewPasswordBox.Password
    $disableRoot = $DisableRootCheck.IsChecked -and $DisableRootCheck.IsChecked.Value

    # Basic validation
    if ([string]::IsNullOrWhiteSpace($host) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($port)) {
        $AppendLog.Invoke("[ERROR] Host, Port, and Username are required fields.`r`n")
        return
    }
    if ($port -notmatch '^\d+$') {
        $AppendLog.Invoke("[ERROR] Port must be a number.`r`n")
        return
    }
    if ([string]::IsNullOrWhiteSpace($newUser)) {
        $AppendLog.Invoke("[ERROR] New Username is required.`r`n")
        return
    }

    # If a key file is provided but ends with .pem or .key, ensure conversion to .ppk
    $ppkPath = $null
    $useKeyAuth = $false
    if (-not [string]::IsNullOrWhiteSpace($keyPath)) {
        $ext = [IO.Path]::GetExtension($keyPath).ToLower()
        if ($ext -ne ".ppk") {
            # Need to convert key
            $AppendLog.Invoke("Converting SSH key to .ppk format...`r`n")
            # Check if puttygen is available
            $puttygenPath = ""
            try {
                $puttygenCmd = Get-Command puttygen.exe -ErrorAction Stop
                $puttygenPath = $puttygenCmd.Source
            } catch {
                # Try to find in temp if we downloaded earlier
                $tempPg = Join-Path [IO.Path]::GetTempPath() "puttygen.exe"
                if (Test-Path $tempPg) {
                    $puttygenPath = $tempPg
                }
            }
            $ppkPath = [IO.Path]::Combine([IO.Path]::GetTempPath(), ([IO.Path]::GetFileNameWithoutExtension($keyPath) + ".ppk"))
            if ($puttygenPath) {
                try {
                    # Use WinSCP's command-line conversion if available (winscp.com)
                    $winscp = Get-Command winscp.com -ErrorAction SilentlyContinue
                    if ($winscp) {
                        & $winscp /keygen "$keyPath" /output="$ppkPath" | Out-Null
                        if (Test-Path $ppkPath) {
                            $AppendLog.Invoke("Key converted to PPK: $ppkPath`r`n")
                            $keyPath = $ppkPath
                        }
                    } else {
                        # If winscp not present, use puttygen via GUI (non-automated)
                        $AppendLog.Invoke("[INFO] Launching PuTTYgen for manual conversion...`r`n")
                        Start-Process -FilePath $puttygenPath -ArgumentList "`"$keyPath`""
                        # Ask user to convert manually
                        $AppendLog.Invoke("Please save the key as .ppk in PuTTYgen, then click Continue.`r`n")
                        # Open file dialog to select the newly saved .ppk
                        $dlg = New-Object Microsoft.Win32.OpenFileDialog
                        $dlg.Title = "Select Converted .ppk Key"
                        $dlg.Filter = "PuTTY Private Key (*.ppk)|*.ppk"
                        if ($dlg.ShowDialog()) {
                            $keyPath = $dlg.FileName
                            $AppendLog.Invoke("Using converted key: $keyPath`r`n")
                        } else {
                            $AppendLog.Invoke("[ERROR] Key conversion cancelled by user.`r`n")
                            return
                        }
                    }
                } catch {
                    $AppendLog.Invoke("[ERROR] Key conversion failed: $_`r`n")
                    return
                }
            } else {
                $AppendLog.Invoke("[ERROR] No PuTTYgen available to convert key. Provide a .ppk key or install PuTTYgen.`r`n")
                return
            }
        }
        # After potential conversion, set flag
        if (-not [string]::IsNullOrWhiteSpace($keyPath)) {
            $useKeyAuth = $true
        }
    }

    # Determine if we will use password for initial login
    $usePasswordAuth = (-not $useKeyAuth)
    if ($useKeyAuth) {
        # If key auth is chosen, we won't use the provided login password for SSH (but might use it for sudo if given)
        $AppendLog.Invoke("Attempting key-based login for $user@$host...`r`n")
    } else {
        $AppendLog.Invoke("Attempting password login for $user@$host...`r`n")
    }

    # Helper function: run a command via plink and capture output in real-time
    function Run-PlinkCommand {
        param(
            [string]$remoteCmd,
            [switch]$elevated   # whether to run under sudo (prefix will be added outside if needed)
        )
        # Build base plink arguments
        $args = "-batch -ssh -P $port"
        if ($useKeyAuth) {
            $args += " -i `"$keyPath`""
        } elseif ($usePasswordAuth) {
            # Note: Using -pw for initial or test login commands if needed
            if ($remoteCmd -eq "exit") {
                # For initial host key accept, include password if available
                if ($pass) { $args += " -pw `"$pass`"" }
            } else {
                if ($pass) { $args += " -pw `"$pass`"" }
            }
        }
        $args += " -l $user `"$remoteCmd`""
        # Prepare process start info
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $global:PlinkPath
        $psi.Arguments = $args
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.RedirectStandardInput = $false
        $psi.CreateNoWindow = $true

        # Process and event handlers
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        # Output data handler
        $proc.add_OutputDataReceived({
            if ($_.Data -ne $null) {
                $AppendLog.Invoke("$($_.Data)`r`n")
            }
        })
        # Error data handler
        $proc.add_ErrorDataReceived({
            if ($_.Data -ne $null) {
                # Prefix error lines for clarity
                $AppendLog.Invoke("[ERROR] $($_.Data)`r`n")
            }
        })
        $proc.Start() | Out-Null
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
        # Wait until process exits, keeping UI responsive
        while (-not $proc.HasExited) {
            Start-Sleep -Milliseconds 100
            [System.Windows.Forms.Application]::DoEvents()  # process UI events
        }
        $proc.WaitForExit()  # ensure fully finished
        return $proc.ExitCode
    }

    # Special handling: initial connect to accept host key
    # Use echo 'y' piped to plink if host key is not cached
    $initialExit = 0
    $needHostKeyAccept = $false
    # Check if host key exists in registry
    $regPath = "HKCU:\Software\SimonTatham\PuTTY\SshHostKeys"
    $targetKeyPrefix = "*@$port:$host"
    try {
        $regItem = Get-Item $regPath -ErrorAction Stop
        $existingKeys = $regItem.Property -like $targetKeyPrefix
    } catch {
        $existingKeys = @()
    }
    if ($existingKeys.Count -eq 0) {
        $needHostKeyAccept = $true
    }
    # Also set needHostKeyAccept if known key might mismatch (we will know only after trying, so we try once)
    # Attempt an initial connection (non-batch) in case host key is cached but possibly changed
    $AppendLog.Invoke("Connecting to $host on port $port...`r`n")
    # Prepare plink for initial connect (without -batch to allow prompt, but we feed 'y')
    $psiInit = New-Object System.Diagnostics.ProcessStartInfo
    $psiInit.FileName = $global:PlinkPath
    $psiInit.UseShellExecute = $false
    $psiInit.RedirectStandardOutput = $true
    $psiInit.RedirectStandardError = $true
    $psiInit.RedirectStandardInput = $true
    $psiInit.CreateNoWindow = $true
    # Build arguments for initial connect
    $argsInit = "-ssh -P $port -legacy-stdio-prompts"
    if ($useKeyAuth) {
        $argsInit += " -i `"$keyPath`""
    } elseif ($usePasswordAuth -and $pass) {
        $argsInit += " -pw `"$pass`""
    }
    $argsInit += " -l $user `"exit`""
    $psiInit.Arguments = $argsInit
    $procInit = New-Object System.Diagnostics.Process
    $procInit.StartInfo = $psiInit
    # Attach handlers for host key prompt or errors
    $hostKeyChanged = $false
    $procInit.add_OutputDataReceived({
        if ($_.Data) {
            $AppendLog.Invoke("$($_.Data)`r`n")
            if ($_.Data -like "*host key is not cached*") {
                # Unknown host key scenario
                $needHostKeyAccept = $true
            }
            if ($_.Data -like "*WARNING:* *host key* *has changed*") {
                $hostKeyChanged = $true
            }
        }
    })
    $procInit.add_ErrorDataReceived({
        if ($_.Data) {
            $AppendLog.Invoke("[ERROR] $($_.Data)`r`n")
            if ($_.Data -like "*Host key verification failed*" -or $_.Data -like "*WARNING: REMOTE HOST IDENTIFICATION HAS CHANGED*") {
                $hostKeyChanged = $true
            }
            if ($_.Data -like "*Access denied*" -or $_.Data -like "*No supported authentication methods*") {
                # Authentication failure (bad password or key)
                # This will be handled after process exit
            }
        }
    })
    $procInit.Start() | Out-Null
    $procInit.BeginOutputReadLine()
    $procInit.BeginErrorReadLine()
    # If we suspect host key prompt needed, send "y"
    # (We always send 'y' to ensure acceptance if prompted)
    $procInit.StandardInput.WriteLine("y")
    $procInit.StandardInput.Close()
    $procInit.WaitForExit()
    $initialExit = $procInit.ExitCode

    # If host key changed, remove old key and retry initial connect
    if ($hostKeyChanged) {
        $AppendLog.Invoke("[WARN] Host key has changed! Removing old key from registry...`r`n")
        try {
            $reg = Get-Item $regPath -ErrorAction Stop
            $propsToRemove = $reg.Property | Where-Object { $_ -like "*@$port:$host" }
            foreach ($prop in $propsToRemove) {
                Remove-ItemProperty -Path $regPath -Name $prop -ErrorAction SilentlyContinue
            }
            $AppendLog.Invoke("Old host key removed. Accepting new key...`r`n")
        } catch {
            $AppendLog.Invoke("[ERROR] Unable to remove old host key: $_`r`n")
        }
        # Retry connection to accept new key
        $procInit = New-Object System.Diagnostics.Process
        $procInit.StartInfo = $psiInit
        $procInit.add_OutputDataReceived({
            if ($_.Data) { $AppendLog.Invoke("$($_.Data)`r`n") }
        })
        $procInit.add_ErrorDataReceived({
            if ($_.Data) { $AppendLog.Invoke("[ERROR] $($_.Data)`r`n") }
        })
        $procInit.Start() | Out-Null
        $procInit.BeginOutputReadLine()
        $procInit.BeginErrorReadLine()
        $procInit.StandardInput.WriteLine("y")
        $procInit.StandardInput.Close()
        $procInit.WaitForExit()
        $initialExit = $procInit.ExitCode
    }

    # Check initial connection result
    if ($initialExit -ne 0) {
        # If authentication failed (exit code nonzero and we saw Access denied or similar)
        if ($initialExit -ne 0) {
            $AppendLog.Invoke("[ERROR] Initial connection failed. Check host, username, and credentials.`r`n")
        }
        return
    }
    $AppendLog.Invoke("Connected and host key verified.`r`n")

    # Decide on prefix for privileged commands (if user is root, no sudo needed)
    $sudoPrefix = ""
    if ($user -ne "root") { $sudoPrefix = "sudo -S -p '' " }

    # Decide on sudo password usage
    $sudoPassword = $null
    if ($user -eq "root") {
        # root user, no sudo needed
    } else {
        # If user provided a login password and we suspect it's the same for sudo, use it.
        if ($pass -and -not $useKeyAuth) {
            $sudoPassword = $pass
        }
    }

    # Helper: run a remote command with optional sudo and password
    function Run-RemoteCommand {
        param([string]$cmd, [bool]$checkOutput=$false)
        # Build actual command string with sudo if needed
        $fullCmd = $cmd
        if ($user -ne "root") {
            if ($sudoPassword) {
                $fullCmd = "echo `"${sudoPassword}`" | sudo -S -p '' $cmd"
            } else {
                $fullCmd = "sudo -S -p '' $cmd"
            }
        }
        return Run-PlinkCommand -remoteCmd $fullCmd
    }

    # Create new user
    $AppendLog.Invoke("Creating new user '$newUser'...`r`n")
    $exitCode = Run-RemoteCommand -cmd "useradd -m -s /bin/bash $newUser"
    if ($exitCode -ne 0) {
        # If failed due to sudo needing password or requiretty, handle those:
        # Check last log lines for hints
        $logText = $LogBox.Text
        if ($logText -match "password is required" -or $logText -match "sudo: \d+ incorrect password attempt") {
            # Sudo password required or was wrong
            # Prompt user for sudo password
            $AppendLog.Invoke("[INFO] Sudo password required for creating user.`r`n")
            $Window.Dispatcher.Invoke([Action]{$SudoOverlay.Visibility = "Visible"})
            # Wait for user input (this is handled by SudoOkButton/SudoCancelButton events below)
            return
        }
        if ($logText -match "sorry, you must have a tty") {
            # requiretty enforced
            $AppendLog.Invoke("[INFO] sudo requires tty. Prompting to disable requiretty...`r`n")
            $Window.Dispatcher.Invoke([Action]{$RequireOverlay.Visibility = "Visible"})
            return
        }
        # Other errors (like user already exists) - stop
        $AppendLog.Invoke("[ERROR] Failed to create user. Setup aborted.`r`n")
        return
    }
    $AppendLog.Invoke("User '$newUser' created successfully.`r`n")

    # Add user to sudo group
    $AppendLog.Invoke("Adding '$newUser' to sudo group...`r`n")
    $exitCode = Run-RemoteCommand -cmd "usermod -aG sudo $newUser"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to add user to sudo group. (Exit code $exitCode)`r`n")
        return
    } else {
        $AppendLog.Invoke("User '$newUser' is now in sudo group.`r`n")
    }

    # Set password for new user if provided
    if (-not [string]::IsNullOrWhiteSpace($newPass)) {
        $AppendLog.Invoke("Setting password for '$newUser'...`r`n")
        # Note: send newuser:password to chpasswd
        $exitCode = Run-RemoteCommand -cmd "bash -c `"echo '$newUser:$newPass' | chpasswd`""
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to set password for $newUser. (Exit code $exitCode)`r`n")
            # Not critical enough to abort; continue without password
        } else {
            $AppendLog.Invoke("Password set for '$newUser'.`r`n")
        }
    } else {
        $AppendLog.Invoke("No password set for '$newUser' (account will be key-only).`r`n")
    }

    # Copy SSH authorized_keys to new user (if exists)
    $AppendLog.Invoke("Copying authorized_keys to '$newUser'...`r`n")
    # Ensure .ssh directory and copy key file
    # We'll do: mkdir .ssh, copy file if exists, set perms
    $cmdCopy = "bash -c 'mkdir -p /home/$newUser/.ssh && "
    $cmdCopy += "if [ -f ~/.ssh/authorized_keys ]; then cp ~/.ssh/authorized_keys /home/$newUser/.ssh/authorized_keys && chmod 600 /home/$newUser/.ssh/authorized_keys; else echo \"NO_AUTH_KEYS\"; fi && "
    $cmdCopy += "chown -R $newUser:$newUser /home/$newUser/.ssh && chmod 700 /home/$newUser/.ssh'"
    $exitCode = Run-RemoteCommand -cmd $cmdCopy
    if ($exitCode -ne 0) {
        if ($LogBox.Text.Contains("NO_AUTH_KEYS")) {
            $AppendLog.Invoke("[WARN] No authorized_keys found for $user; none copied for $newUser.`r`n")
        } else {
            $AppendLog.Invoke("[ERROR] Failed to copy authorized_keys (Exit code $exitCode).`r`n")
        }
        # Continue even if none copied, as user might still log in via password
    } else {
        $AppendLog.Invoke("SSH key authorized for '$newUser'.`r`n")
    }

    # Test new user login
    $AppendLog.Invoke("Testing login for new user '$newUser'...`r`n")
    $testExit = 1
    # Test with key or password depending on what's available
    if ($useKeyAuth -and (Test-Path $keyPath)) {
        # Use same key for new user
        $AppendLog.Invoke("(Using key authentication for test)`r`n")
        $testArgs = "-batch -ssh -P $port -i `"$keyPath`" -l $newUser `"echo OK`""
        $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = $global:PlinkPath; Arguments = $testArgs;
            UseShellExecute = $false; RedirectStandardOutput = $true; RedirectStandardError = $true; CreateNoWindow = $true
        }
        $procTest = New-Object System.Diagnostics.Process
        $procTest.StartInfo = $psiTest
        $out = New-Object System.Text.StringBuilder
        $err = New-Object System.Text.StringBuilder
        $procTest.add_OutputDataReceived({ if ($_.Data) { $out.AppendLine($_.Data) } })
        $procTest.add_ErrorDataReceived({ if ($_.Data) { $err.AppendLine($_.Data) } })
        $procTest.Start() | Out-Null
        $procTest.BeginOutputReadLine()
        $procTest.BeginErrorReadLine()
        $procTest.WaitForExit()
        $testExit = $procTest.ExitCode
        if ($out.Length -gt 0) {
            $AppendLog.Invoke("New user output: $($out.ToString().Trim())`r`n")
        }
        if ($err.Length -gt 0) {
            $AppendLog.Invoke("[ERROR] New user error: $($err.ToString().Trim())`r`n")
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($newPass)) {
        # Use password to test
        $AppendLog.Invoke("(Using password authentication for test)`r`n")
        $testArgs = "-batch -ssh -P $port -l $newUser -pw `"$newPass`" `"echo OK`""
        $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = $global:PlinkPath; Arguments = $testArgs;
            UseShellExecute = $false; RedirectStandardOutput = $true; RedirectStandardError = $true; CreateNoWindow = $true
        }
        $procTest = New-Object System.Diagnostics.Process
        $procTest.StartInfo = $psiTest
        $out = New-Object System.Text.StringBuilder
        $err = New-Object System.Text.StringBuilder
        $procTest.add_OutputDataReceived({ if ($_.Data) { $out.AppendLine($_.Data) } })
        $procTest.add_ErrorDataReceived({ if ($_.Data) { $err.AppendLine($_.Data) } })
        $procTest.Start() | Out-Null
        $procTest.BeginOutputReadLine()
        $procTest.BeginErrorReadLine()
        $procTest.WaitForExit()
        $testExit = $procTest.ExitCode
        if ($out.Length -gt 0) {
            $AppendLog.Invoke("New user output: $($out.ToString().Trim())`r`n")
        }
        if ($err.Length -gt 0) {
            $AppendLog.Invoke("[ERROR] New user error: $($err.ToString().Trim())`r`n")
        }
    } else {
        $AppendLog.Invoke("[WARN] No key or password available to test new user login.`r`n")
        $testExit = 1
    }
    if ($testExit -eq 0) {
        $AppendLog.Invoke("New user '$newUser' login test successful.`r`n")
    } else {
        $AppendLog.Invoke("[ERROR] New user '$newUser' login test failed. Root login will not be disabled.`r`n")
        $disableRoot = $false
    }

    # Disable root login if selected and test was successful
    if ($disableRoot -and $user -ne "root") {
        $AppendLog.Invoke("Disabling root SSH login...`r`n")
        $cmdDisableRoot = "sed -i 's/^#\?PermitRootLogin.*/PermitRootLogin no/' /etc/ssh/sshd_config"
        $exitCode = Run-RemoteCommand -cmd $cmdDisableRoot
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to edit sshd_config to disable root login.`r`n")
        } else {
            $AppendLog.Invoke("Root login disabled in SSH config.`r`n")
            # Reload SSH service to apply change
            $exitCode = Run-RemoteCommand -cmd "bash -c 'which systemctl >/dev/null 2>&1 && (systemctl reload sshd || systemctl reload ssh) || service ssh reload'"
            if ($exitCode -ne 0) {
                $AppendLog.Invoke("[WARN] SSH service reload failed. You may need to restart SSH manually for changes to take effect.`r`n")
            } else {
                $AppendLog.Invoke("SSH service reloaded to apply new settings.`r`n")
            }
        }
    }
    $AppendLog.Invoke("Setup process complete.`r`n")
})

# Handle Sudo password overlay OK
$SudoOkButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $SudoOverlay.Visibility = "Collapsed"
    })
    $sudoPassInput = $SudoPassBox.Password
    if ([string]::IsNullOrEmpty($sudoPassInput)) {
        $AppendLog.Invoke("[ERROR] Sudo password was not provided. Aborting.`r`n")
        return
    }
    # Store the sudo password and retry the last command that required it
    $AppendLog.Invoke("Sudo password received. Resuming operations...`r`n")
    # Set global sudo password for subsequent commands
    $script:sudoPassword = $sudoPassInput
    # Retry user creation (the point where we left off)
    $exitCode = Run-RemoteCommand -cmd "useradd -m -s /bin/bash $($NewUserBox.Text.Trim())"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to create user even after sudo password. Aborting.`r`n")
        return
    }
    $AppendLog.Invoke("User '$($NewUserBox.Text.Trim())' created successfully (after sudo password).`r`n")
    # Continue with subsequent steps (simulate clicking Start again from after user creation)
    # We will simply call the StartButton's click handler recursively to continue.
    # To avoid infinite recursion, perhaps factor remaining steps into separate function.
    # For simplicity, we call a separate function that continues the setup from after user creation.
    Continue-SetupAfterUserCreation
})

# Handle Sudo password overlay Cancel
$SudoCancelButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $SudoOverlay.Visibility = "Collapsed"
    })
    $AppendLog.Invoke("[ERROR] Sudo password prompt canceled by user. Aborting setup.`r`n")
})

# Handle requiretty overlay Yes
$RequireYesButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $RequireOverlay.Visibility = "Collapsed"
    })
    $AppendLog.Invoke("Disabling requiretty and resuming operations...`r`n")
    # If sudo password is needed for editing sudoers and not known, prompt for it
    if (-not $script:sudoPassword) {
        # We need a password to edit sudoers if one was not already provided
        $AppendLog.Invoke("Please enter sudo password to disable requiretty...`r`n")
        $Window.Dispatcher.Invoke([Action]{
            $SudoOverlay.Visibility = "Visible"
        })
        # (After getting sudo password via SudoOkButton, code will continue in that handler)
    } else {
        # We have a sudo password from before, proceed to disable requiretty
        Disable-RequirettyAndContinue
    }
})

# Handle requiretty overlay No
$RequireNoButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $RequireOverlay.Visibility = "Collapsed"
    })
    $AppendLog.Invoke("User chose not to disable requiretty. Aborting setup.`r`n")
})

# Function to disable requiretty and resume
function Disable-RequirettyAndContinue {
    # Use sudo to comment out requiretty in /etc/sudoers
    $cmd = "sed -i 's/^Defaults\s\+requiretty/#&/' /etc/sudoers"
    # This needs a tty, so use plink -t for this command
    $plinkArgs = "-ssh -t -P $($PortBox.Text.Trim()) -l $($UserBox.Text.Trim())"
    if ($useKeyAuth) {
        $plinkArgs += " -i `"$($KeyBox.Text)`""
    } elseif ($usePasswordAuth) {
        $plinkArgs += " -pw `"$($PasswordBox.Password)`""
    }
    $plinkArgs += " `"echo $script:sudoPassword | sudo -S -p '' sed -i 's/^Defaults\s\+requiretty/#&/' /etc/sudoers`""
    try {
        $output = & $global:PlinkPath $plinkArgs
        $AppendLog.Invoke("requiretty disabled in sudoers.`r`n")
    } catch {
        $AppendLog.Invoke("[ERROR] Failed to disable requiretty: $_`r`n")
    }
    # Resume creating user after disabling requiretty
    $AppendLog.Invoke("Resuming user creation...`r`n")
    $exitCode = Run-RemoteCommand -cmd "useradd -m -s /bin/bash $($NewUserBox.Text.Trim())"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to create user even after disabling requiretty. Aborting.`r`n")
        return
    }
    $AppendLog.Invoke("User '$($NewUserBox.Text.Trim())' created successfully.`r`n")
    Continue-SetupAfterUserCreation
}

# Function to continue setup steps after creating the new user (used when resuming after prompts)
function Continue-SetupAfterUserCreation {
    # (This function will carry out the steps: add to sudo group, set password, copy keys, test, disable root)
    $newUser = $NewUserBox.Text.Trim()
    # Add to sudo group
    $AppendLog.Invoke("Adding '$newUser' to sudo group...`r`n")
    $exitCode = Run-RemoteCommand -cmd "usermod -aG sudo $newUser"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to add user to sudo group. Aborting.`r`n")
        return
    } else {
        $AppendLog.Invoke("User '$newUser' added to sudo group.`r`n")
    }
    # Set password if given
    $newPass = $NewPasswordBox.Password
    if (-not [string]::IsNullOrEmpty($newPass)) {
        $AppendLog.Invoke("Setting password for '$newUser'...`r`n")
        $exitCode = Run-RemoteCommand -cmd "bash -c `"echo '$newUser:$newPass' | chpasswd`""
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to set password for $newUser.`r`n")
        } else {
            $AppendLog.Invoke("Password set for '$newUser'.`r`n")
        }
    }
    # Copy authorized_keys
    $AppendLog.Invoke("Copying authorized_keys to '$newUser'...`r`n")
    $cmdCopy = "bash -c 'mkdir -p /home/$newUser/.ssh && "
    $cmdCopy += "if [ -f ~/.ssh/authorized_keys ]; then cp ~/.ssh/authorized_keys /home/$newUser/.ssh/authorized_keys && chmod 600 /home/$newUser/.ssh/authorized_keys; else echo \"NO_AUTH_KEYS\"; fi && "
    $cmdCopy += "chown -R $newUser:$newUser /home/$newUser/.ssh && chmod 700 /home/$newUser/.ssh'"
    $exitCode = Run-RemoteCommand -cmd $cmdCopy
    if ($exitCode -ne 0) {
        if ($LogBox.Text.Contains("NO_AUTH_KEYS")) {
            $AppendLog.Invoke("[WARN] No authorized_keys to copy.`r`n")
        } else {
            $AppendLog.Invoke("[ERROR] Error copying authorized_keys. (Exit $exitCode)`r`n")
        }
    } else {
        $AppendLog.Invoke("SSH key authorized for '$newUser'.`r`n")
    }
    # Test new user login (same logic as above)
    $AppendLog.Invoke("Testing login for new user '$newUser'...`r`n")
    $testExit = 1
    $useKey = (-not [string]::IsNullOrEmpty($KeyBox.Text))
    if ($useKey -and (Test-Path $KeyBox.Text)) {
        $AppendLog.Invoke("(Using key authentication for test)`r`n")
        $testArgs = "-batch -ssh -P $($PortBox.Text.Trim()) -i `"$($KeyBox.Text)`" -l $newUser `"echo OK`""
        $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = $global:PlinkPath; Arguments = $testArgs;
            UseShellExecute = $false; RedirectStandardOutput = $true; RedirectStandardError = $true; CreateNoWindow = $true
        }
        $procTest = New-Object System.Diagnostics.Process
        $procTest.StartInfo = $psiTest
        $out = New-Object System.Text.StringBuilder
        $err = New-Object System.Text.StringBuilder
        $procTest.add_OutputDataReceived({ if ($_.Data) { $out.AppendLine($_.Data) } })
        $procTest.add_ErrorDataReceived({ if ($_.Data) { $err.AppendLine($_.Data) } })
        $procTest.Start() | Out-Null
        $procTest.BeginOutputReadLine()
        $procTest.BeginErrorReadLine()
        $procTest.WaitForExit()
        $testExit = $procTest.ExitCode
        if ($out.Length -gt 0) {
            $AppendLog.Invoke("New user output: $($out.ToString().Trim())`r`n")
        }
        if ($err.Length -gt 0) {
            $AppendLog.Invoke("[ERROR] $($err.ToString().Trim())`r`n")
        }
    } elseif (-not [string]::IsNullOrEmpty($newPass)) {
        $AppendLog.Invoke("(Using password authentication for test)`r`n")
        $testArgs = "-batch -ssh -P $($PortBox.Text.Trim()) -l $newUser -pw `"$newPass`" `"echo OK`""
        $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = $global:PlinkPath; Arguments = $testArgs;
            UseShellExecute = $false; RedirectStandardOutput = $true; RedirectStandardError = $true; CreateNoWindow = $true
        }
        $procTest = New-Object System.Diagnostics.Process
        $procTest.StartInfo = $psiTest
        $out = New-Object System.Text.StringBuilder
        $err = New-Object System.Text.StringBuilder
        $procTest.add_OutputDataReceived({ if ($_.Data) { $out.AppendLine($_.Data) } })
        $procTest.add_ErrorDataReceived({ if ($_.Data) { $err.AppendLine($_.Data) } })
        $procTest.Start() | Out-Null
        $procTest.BeginOutputReadLine()
        $procTest.BeginErrorReadLine()
        $procTest.WaitForExit()
        $testExit = $procTest.ExitCode
        if ($out.Length -gt 0) {
            $AppendLog.Invoke("New user output: $($out.ToString().Trim())`r`n")
        }
        if ($err.Length -gt 0) {
            $AppendLog.Invoke("[ERROR] $($err.ToString().Trim())`r`n")
        }
    }
    if ($testExit -eq 0) {
        $AppendLog.Invoke("New user '$newUser' login test successful.`r`n")
    } else {
        $AppendLog.Invoke("[ERROR] New user '$newUser' login test failed. Root login will not be disabled.`r`n")
        $DisableRootCheck.IsChecked = $false
    }
    # Disable root if opted and test was success
    if ($DisableRootCheck.IsChecked -and $DisableRootCheck.IsChecked.Value -and $testExit -eq 0 -and $UserBox.Text.Trim() -ne "root") {
        $AppendLog.Invoke("Disabling root SSH login...`r`n")
        $exitCode = Run-RemoteCommand -cmd "sed -i 's/^#\?PermitRootLogin.*/PermitRootLogin no/' /etc/ssh/sshd_config"
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to update sshd_config for PermitRootLogin.`r`n")
        } else {
            $AppendLog.Invoke("Root login disabled in sshd_config.`r`n")
            $exitCode = Run-RemoteCommand -cmd "bash -c 'which systemctl >/dev/null 2>&1 && (systemctl reload sshd || systemctl reload ssh) || service ssh reload'"
            if ($exitCode -ne 0) {
                $AppendLog.Invoke("[WARN] SSH reload failed; you may need to restart SSH manually.`r`n")
            } else {
                $AppendLog.Invoke("SSH service reloaded to apply new configuration.`r`n")
            }
        }
    }
    $AppendLog.Invoke("Setup process complete.`r`n")
}

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
