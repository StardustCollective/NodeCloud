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
            $keyPart = if (-not [string]::IsNullOrEmpty($KeyBox.Text)) { "-i $($KeyBox.Text) " } else { "" }

            $content += "# SSH login: ssh $keyPart$($UserBox.Text)@$($HostBox.Text)"
            $content += "# SFTP: sftp $keyPart$($UserBox.Text)@$($HostBox.Text)"
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
                $tempPg = Join-Path ([IO.Path]::GetTempPath()) "puttygen.exe"
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
    $targetKeyPrefix = "*@${port}:${host}"
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
        $escapedPair = "$newUser`:$newPass"
$exitCode = Run-RemoteCommand -cmd "bash -c `"echo '$escapedPair' | chpasswd`""
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
