# # Relaunch elevated if not already running as Administrator
# if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
#     [Security.Principal.WindowsBuiltInRole] "Administrator"
# ))
# {
#     $scriptPath = $PSCommandPath
#     if (-not $scriptPath) {
#         $scriptPath = $MyInvocation.MyCommand.Path
#     }

#     if (-not $scriptPath) {
#         Write-Host "Unable to determine script path for elevation. Please run this script from a .ps1 file."
#         pause
#         exit
#     }

#     $quoted = '"' + $scriptPath + '"'
#     Start-Process powershell.exe -Verb RunAs `
#         -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File $quoted" `
#         -WorkingDirectory (Split-Path $scriptPath -Parent)

#     exit
# }

# Load required .NET assemblies for WPF
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Xaml

# Define the XAML for the WPF UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Server Toolkit" Height="650" Width="700" WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        Background="#1e1f22"
        Foreground="#f0f0f0"
        FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="WindowBg" Color="#1e1f22" />
        <SolidColorBrush x:Key="PanelBg" Color="#25272b" />
        <SolidColorBrush x:Key="InputBg" Color="#2f3238" />
        <SolidColorBrush x:Key="InputBorder" Color="#4a4f5a" />
        <SolidColorBrush x:Key="InputBorderFocused" Color="#00a8ff" />
        <SolidColorBrush x:Key="LabelFg" Color="#d0d0d0" />
        <SolidColorBrush x:Key="TextFg" Color="#f0f0f0" />
        <SolidColorBrush x:Key="AccentBrush" Color="#00a8ff" />
        <SolidColorBrush x:Key="ButtonBg" Color="#00a8ff" />
        <SolidColorBrush x:Key="ButtonBgHover" Color="#14b5ff" />
        <SolidColorBrush x:Key="ButtonBgPressed" Color="#0090d0" />
        <SolidColorBrush x:Key="BorderBrushDark" Color="#3b3f46" />

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource LabelFg}" />
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="{StaticResource LabelFg}" />
            <Setter Property="Margin" Value="0,0,4,4" />
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
        </Style>

        <Style TargetType="GroupBox">
            <Setter Property="Background" Value="{StaticResource PanelBg}" />
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushDark}" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="10" />
            <Setter Property="Margin" Value="0,0,0,12" />
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource InputBg}" />
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
            <Setter Property="BorderBrush" Value="{StaticResource InputBorder}" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="4,2" />
            <Style.Triggers>
                <Trigger Property="IsKeyboardFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource InputBorderFocused}" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="{StaticResource InputBg}" />
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
            <Setter Property="BorderBrush" Value="{StaticResource InputBorder}" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="4,2" />
            <Style.Triggers>
                <Trigger Property="IsKeyboardFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource InputBorderFocused}" />
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="{StaticResource ButtonBg}" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="10,4" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="4,1"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource ButtonBgHover}" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource ButtonBgPressed}" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Menu">
            <Setter Property="Background" Value="#262a30" />
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
        </Style>
        <Style TargetType="MenuItem">
            <Setter Property="Background" Value="#262a30" />
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
            <Setter Property="Padding" Value="8,2" />
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#333842" />
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <DockPanel LastChildFill="True">
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem Header="_Load Profile..." x:Name="MenuLoadProfile"/>
                <MenuItem Header="_Save Profile..." x:Name="MenuSaveProfile"/>
                <Separator/>
                <MenuItem Header="E_xit" x:Name="MenuExit"/>
            </MenuItem>
        </Menu>
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

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
                    <Label Content="Host:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="HostBox" Grid.Row="0" Grid.Column="1" Margin="5,2" />
                    <Label Content="Port:" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="PortBox" Grid.Row="1" Grid.Column="1" Margin="5,2" Width="60" Text="22"/>
                    <Label Content="Username:" Grid.Row="2" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="UserBox" Grid.Row="2" Grid.Column="1" Margin="5,2" />
                    <Label Content="Password:" Grid.Row="3" Grid.Column="0" VerticalAlignment="Center"/>
                    <PasswordBox x:Name="PasswordBox" Grid.Row="3" Grid.Column="1" Margin="5,2" />
                    <Label Content="SSH Key File:" Grid.Row="4" Grid.Column="0" VerticalAlignment="Center"/>
                    <TextBox x:Name="KeyBox" Grid.Row="4" Grid.Column="1" Margin="5,2" IsReadOnly="True"/>
                    <Button x:Name="BrowseKeyButton" Content="Browse..." Grid.Row="4" Grid.Column="2" Margin="5,2,0,2" Padding="8,0" />
                </Grid>
            </GroupBox>

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

            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="5,0,0,10">
                <CheckBox x:Name="DisableRootCheck" Content="Disable root login after setup" IsChecked="True"/>
            </StackPanel>

            <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,0,0,10">
                <Button x:Name="StartButton" Content="Start Setup" Width="100" HorizontalAlignment="Center"/>
            </StackPanel>

            <TextBox x:Name="LogBox" Grid.Row="4" Margin="0" Background="#FF000000" Foreground="#FF00FF00"
                     FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" AcceptsReturn="True"/>

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

$reader = New-Object System.Xml.XmlNodeReader $xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

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

$global:ServerToolkitLogPath = Join-Path $env:USERPROFILE "server-toolkit.log"

$AppendLog = [Action[string]]{
    param($text)
    try {
        if (-not [string]::IsNullOrEmpty($global:ServerToolkitLogPath)) {
            $dir = Split-Path -Parent $global:ServerToolkitLogPath
            if (-not [IO.Directory]::Exists($dir)) {
                [IO.Directory]::CreateDirectory($dir) | Out-Null
            }
            Add-Content -Path $global:ServerToolkitLogPath -Value $text -ErrorAction SilentlyContinue
        }
    } catch {}

    $isDebug = $false
    if ($text -like "DEBUG:*") {
        $isDebug = $true
    }

    if (-not $isDebug) {
        try {
            if ($LogBox -and $LogBox.Dispatcher) {
                $LogBox.Dispatcher.Invoke([Action]{
                    $LogBox.AppendText($text)
                    $LogBox.ScrollToEnd()
                })
            }
        } catch {}
    }
}

$script:useKeyAuth       = $false
$script:usePasswordAuth  = $false
$script:sudoPassword     = $null
$script:agentKeyAddedThisSession = $false
$global:PlinkPath        = ""

function Ensure-PuttyTools {
    $global:PlinkPath = ""
    try {
        $plinkCmd = Get-Command plink.exe -ErrorAction Stop
        $global:PlinkPath = $plinkCmd.Source
    } catch {
        $AppendLog.Invoke("[INFO] plink.exe not found. Attempting to download PuTTY tools...`r`n")
        $arch = if ([IntPtr]::Size -eq 8) {"w64"} else {"w32"}
        $plinkUrl = "https://the.earth.li/~sgtatham/putty/latest/$arch/plink.exe"
        $puttygenUrl = "https://the.earth.li/~sgtatham/putty/latest/$arch/puttygen.exe"
        $tempPath = [IO.Path]::GetTempPath()
        $plinkTarget = Join-Path $tempPath "plink.exe"
        try {
            (New-Object Net.WebClient).DownloadFile($plinkUrl, $plinkTarget)
            $AppendLog.Invoke("Downloaded plink.exe to $plinkTarget`r`n")
            $global:PlinkPath = $plinkTarget
        } catch {
            $AppendLog.Invoke("[ERROR] Failed to download plink.exe. Please install PuTTY or provide plink.exe manually.`r`n")
            return $false
        }
        $puttygenTarget = Join-Path $tempPath "puttygen.exe"
        try {
            (New-Object Net.WebClient).DownloadFile($puttygenUrl, $puttygenTarget)
            $AppendLog.Invoke("Downloaded puttygen.exe to $puttygenTarget`r`n")
        } catch {
            $AppendLog.Invoke("[WARN] Could not download puttygen.exe. .ppk conversion might require manual steps.`r`n")
        }
    }
    return $true
}

function Reset-HostKeysForServer {
    param(
        [string]$HostName,
        [string]$Port
    )

    try {
        $regPath = "HKCU:\Software\SimonTatham\PuTTY\SshHostKeys"
        if (Test-Path $regPath) {
            $regItem      = Get-Item $regPath
            $targetPrefix = "*@${Port}:${HostName}"
            $propsToRemove = $regItem.Property | Where-Object { $_ -like $targetPrefix }
            foreach ($prop in $propsToRemove) {
                Remove-ItemProperty -Path $regPath -Name $prop -ErrorAction SilentlyContinue
            }
            if ($propsToRemove.Count -gt 0) {
                $AppendLog.Invoke("Reset PuTTY host-key cache entries for ${HostName}:$Port.`r`n")
            }
        }
    } catch {
        $AppendLog.Invoke("[WARN] Failed to reset PuTTY host-key cache for ${HostName}:${Port}: $($_.Exception.Message)`r`n")
    }

    try {
        $knownHostsPath = Join-Path (Join-Path $env:USERPROFILE ".ssh") "known_hosts"
        if (Test-Path $knownHostsPath) {
            $allLines    = Get-Content -Path $knownHostsPath
            $escapedHost = [regex]::Escape($HostName)
            $escapedPort = [regex]::Escape($Port)
            $pattern1 = "^$escapedHost\s"
            $pattern2 = "^\[$escapedHost\]:$escapedPort\s"
            $filtered = $allLines | Where-Object {
                $_ -notmatch $pattern1 -and $_ -notmatch $pattern2
            }
            if ($filtered.Count -ne $allLines.Count) {
                $filtered | Set-Content -Path $knownHostsPath
                $AppendLog.Invoke("Reset OpenSSH known_hosts entries for $HostName (port $Port).`r`n")
            }
        }
    } catch {
        $AppendLog.Invoke("[WARN] Failed to clean OpenSSH known_hosts for ${HostName}:${Port}: $($_.Exception.Message)`r`n")
    }
}

function Run-SshCommand {
    param(
        [string]$RemoteCommand,
        [string]$RemoteHost,
        [string]$Port,
        [string]$User,
        [string]$KeyPath
    )

    $sshExe = $null
    $winSsh = Join-Path $env:WINDIR "System32\OpenSSH\ssh.exe"
    if (Test-Path $winSsh) {
        $sshExe = $winSsh
    } else {
        try {
            $sshCmd = Get-Command ssh -ErrorAction Stop
            $sshExe = $sshCmd.Source
        } catch {
            $AppendLog.Invoke("[FATAL] ssh.exe not found. Install Windows OpenSSH Client or ensure ssh is on PATH.`r`n")
            return 1
        }
    }

    $args = @(
        "-o","BatchMode=yes",
        "-o","StrictHostKeyChecking=accept-new",
        "-p",$Port
    )
    if (-not [string]::IsNullOrWhiteSpace($KeyPath)) {
        $args += @("-i",$KeyPath)
    }
    $args += "$User@$RemoteHost"
    $args += $RemoteCommand

    $fullArgs = ($args -join ' ')
    if ($fullArgs -like "*chpasswd*" -or $fullArgs -like "*base64 -d*") {
        $AppendLog.Invoke("DEBUG: Running ssh ($sshExe) with password update command (details redacted).`r`n")
    } else {
        $AppendLog.Invoke("DEBUG: Running ssh ($sshExe) with args: $fullArgs`r`n")
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $sshExe
        $psi.Arguments = ($args -join " ")
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow         = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi

        $proc.Start() | Out-Null

        $timeoutMs = 15000
        $finished = $proc.WaitForExit($timeoutMs)

        if (-not $finished) {
            try { $proc.Kill() } catch {}
            $AppendLog.Invoke("[ERROR] ssh command timed out after ${timeoutMs}ms for: $($args -join ' ')`r`n")
            return 1
        }

        $out = $proc.StandardOutput.ReadToEnd()
        $err = $proc.StandardError.ReadToEnd()

        if ($out) {
            foreach ($line in $out -split "`r?`n") {
                if ($line -ne "") {
                    if ($RemoteCommand -like "*echo OK*" -and $line -eq "OK") {
                        continue
                    }
                    $AppendLog.Invoke("$line`r`n")
                }
            }
        }
        if ($err) {
            foreach ($line in $err -split "`r?`n") {
                if ($line -ne "") {
                    if ($line -like "Warning: Permanently added*") {
                        $AppendLog.Invoke("First-time connection to $RemoteHost confirmed. The server's identity has been saved for future connections.`r`n")
                    } else {
                        $AppendLog.Invoke("[ERROR] $line`r`n")
                    }
                }
            }
        }

        $exitCode = $proc.ExitCode
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] ssh exited with code $exitCode for command: $($args -join ' ')`r`n")
        }
        return $exitCode
    } catch {
        $AppendLog.Invoke("[FATAL] Exception while running ssh: $($_.Exception.Message)`r`n")
        return 1
    }
}

function Get-AgentRegistryKeyForPath {
    param(
        [string]$KeyPath
    )

    $base = "HKCU:\Software\OpenSSH\Agent\Keys"
    if (-not (Test-Path $base)) { return $null }

    $normalized = [IO.Path]::GetFullPath($KeyPath)

    foreach ($sub in Get-ChildItem $base -ErrorAction SilentlyContinue) {
        try {
            $props   = Get-ItemProperty -Path $sub.PSPath -ErrorAction SilentlyContinue
            $comment = $props.comment
            if ($comment) {
                $commentFull = [IO.Path]::GetFullPath($comment)
                if ($commentFull -ieq $normalized) {
                    return $sub.PSPath
                }
            }
        } catch {
            # ignore broken entries
        }
    }

    return $null
}

function Mark-AgentKeyOwned {
    param(
        [string]$KeyPath
    )

    $regKey = Get-AgentRegistryKeyForPath -KeyPath $KeyPath
    if ($regKey) {
        try {
            Set-ItemProperty -Path $regKey -Name "ServerToolkitOwned" -Value 1 -Type DWord
            $AppendLog.Invoke("Marked ssh-agent key as owned by GUI: $KeyPath`r`n")
        } catch {
            $AppendLog.Invoke("[WARN] Failed to mark agent key as owned: $($_.Exception.Message)`r`n")
        }
    } else {
        $AppendLog.Invoke("DEBUG: ssh-agent registry entry not found for key (cannot mark as GUI-owned): $KeyPath`r`n")
    }
}

function Cleanup-AgentOwnedKeys {
    $base = "HKCU:\Software\OpenSSH\Agent\Keys"
    if (-not (Test-Path $base)) { return }

    $sshAddExe = $null
    $winSshAdd = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
    if (Test-Path $winSshAdd) {
        $sshAddExe = $winSshAdd
    } else {
        try {
            $sshAddCmd = Get-Command ssh-add -ErrorAction Stop
            $sshAddExe = $sshAddCmd.Source
        } catch {
            $AppendLog.Invoke("[WARN] ssh-add not found. Cannot clean up agent keys.`r`n")
            return
        }
    }

    foreach ($sub in Get-ChildItem $base -ErrorAction SilentlyContinue) {
        try {
            $props = Get-ItemProperty -Path $sub.PSPath -ErrorAction SilentlyContinue
            if ($props.ServerToolkitOwned -eq 1) {
                $comment = $props.comment
                if ($comment) {
                    $AppendLog.Invoke("Removing GUI-added key from ssh-agent: $comment`r`n")
                    try { & $sshAddExe -d "$comment" | Out-Null } catch {}
                }
                try {
                    Remove-ItemProperty -Path $sub.PSPath -Name "ServerToolkitOwned" -ErrorAction SilentlyContinue
                } catch {}
            }
        } catch {
            # ignore broken entries
        }
    }
}

function Ensure-SshAgentWithKey {
    param(
        [string]$KeyPath,
        [string]$Passphrase
    )

    if ([string]::IsNullOrWhiteSpace($KeyPath)) {
        return $true
    }

    $agentService = $null
    try {
        $agentService = Get-Service ssh-agent -ErrorAction SilentlyContinue
    } catch {
        $agentService = $null
    }

    if (-not $agentService) {
        $AppendLog.Invoke("[WARN] ssh-agent service not found on this system. Skipping agent integration.`r`n")
        return $true
    }

    if ($agentService.StartType -eq 'Disabled') {
        try {
            Set-Service ssh-agent -StartupType Automatic -ErrorAction Stop
            $AppendLog.Invoke("Set ssh-agent StartupType to Automatic.`r`n")
        } catch {
            $AppendLog.Invoke("[WARN] Failed to change ssh-agent StartupType: $($_.Exception.Message)`r`n")
        }
    }

    if ($agentService.Status -ne 'Running') {
        try {
            Start-Service ssh-agent -ErrorAction Stop
            $AppendLog.Invoke("Started ssh-agent service.`r`n")
        } catch {
            $AppendLog.Invoke("[WARN] Could not start ssh-agent service: $($_.Exception.Message)`r`n")
            return $true
        }
    } else {
        $AppendLog.Invoke("ssh-agent service is already running.`r`n")
    }

    $sshAddExe = $null
    $winSshAdd = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
    if (Test-Path $winSshAdd) {
        $sshAddExe = $winSshAdd
    } else {
        try {
            $sshAddCmd = Get-Command ssh-add -ErrorAction Stop
            $sshAddExe = $sshAddCmd.Source
        } catch {
            $AppendLog.Invoke("[WARN] ssh-add not found on PATH. Skipping automatic agent checks.`r`n")
            return $true
        }
    }

    $identitiesLoaded = $false
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $sshAddExe
        $psi.Arguments = "-L"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow         = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        $proc.Start() | Out-Null

        if ($proc.WaitForExit(3000)) {
            $out = $proc.StandardOutput.ReadToEnd()
            $err = $proc.StandardError.ReadToEnd()

            if ($proc.ExitCode -eq 0 -and $out) {
                $keyPathEscaped = [regex]::Escape($KeyPath)
                foreach ($line in $out -split "`r?`n") {
                    if ($line -match $keyPathEscaped) {
                        $identitiesLoaded = $true
                        break
                    }
                }
            } elseif ($err) {
                $AppendLog.Invoke("[WARN] ssh-add -L stderr: $err`r`n")
            }
        } else {
            try { $proc.Kill() } catch {}
            $AppendLog.Invoke("[WARN] ssh-add -L did not finish within 3000ms. Assuming no identities.`r`n")
        }
    } catch {
        $AppendLog.Invoke("[WARN] Failed to query ssh-agent identities: $($_.Exception.Message)`r`n")
    }

    if ($identitiesLoaded) {
        $AppendLog.Invoke("ssh-agent already has this key loaded: $KeyPath`r`n")
        if ($script:agentKeyAddedThisSession) {
            Mark-AgentKeyOwned -KeyPath $KeyPath
            $script:agentKeyAddedThisSession = $false
        }
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($Passphrase)) {
        $AppendLog.Invoke("No key passphrase provided and ssh-agent has no identity for this key.`r`n")
        $AppendLog.Invoke("If your key is encrypted, run this in a PowerShell window BEFORE using this toolkit:`r`n  ssh-add `"$KeyPath`"`r`n")
        return $true
    }

    $sshAddCmd = @"
`$host.UI.RawUI.WindowTitle = 'Server Toolkit - SSH Key Passphrase';

Write-Host ''
Write-Host 'Server Toolkit - SSH Key Passphrase Required' -ForegroundColor Cyan
Write-Host '------------------------------------------------' -ForegroundColor Cyan
Write-Host ''
Write-Host 'A secure terminal window has opened to unlock your SSH key.' -ForegroundColor Yellow
Write-Host 'Please enter the passphrase for your SSH key when prompted.' -ForegroundColor Yellow
Write-Host ''
Write-Host 'Once the passphrase is accepted, you can close this window.' -ForegroundColor Green
Write-Host 'Then return to the Server Toolkit and click "Start Setup" again.' -ForegroundColor Green
Write-Host ''
ssh-add "$KeyPath"
"@

    try {
        $script:agentKeyAddedThisSession = $true
        Start-Process powershell.exe `
            -ArgumentList "-NoProfile -NoLogo -Command $sshAddCmd" `
            -WorkingDirectory (Split-Path $KeyPath -Parent) `
            -WindowStyle Normal | Out-Null

        $AppendLog.Invoke("A new terminal window has opened so you can enter the passphrase for your SSH key.`r`n")
        $AppendLog.Invoke("After entering the passphrase there and closing that window, click Start Setup again to continue.`r`n")


        return $false
    } catch {
        $AppendLog.Invoke("[WARN] Failed to launch interactive ssh-add console: $($_.Exception.Message)`r`n")
        $script:agentKeyAddedThisSession = $false
        return $true
    }
}

$BrowseKeyButton.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = "Select Private Key File"
    $dlg.Filter = "Private Key Files (*.ppk;*.pem;*.key)|*.ppk;*.pem;*.key|All Files|*.*"
    if ($dlg.ShowDialog()) {
        $KeyBox.Text = $dlg.FileName
    }
})

$MenuLoadProfile.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title = "Load SSH Profile"
    $dlg.Filter = "SSH Profile (*.ssh_config.txt)|*_ssh_config.txt;*.ssh_config.txt|All Files|*.*"
    $homeSSH = [IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), ".ssh")
    if ([IO.Directory]::Exists($homeSSH)) { $dlg.InitialDirectory = $homeSSH }
    if ($dlg.ShowDialog()) {
        try {
            $lines = Get-Content $dlg.FileName
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
                    $KeyPathLocal = $matches[1].Trim()
                    if ($KeyPathLocal.StartsWith("~")) {
                        $KeyPathLocal = $KeyPathLocal -replace "^~", [Environment]::GetFolderPath('UserProfile')
                    }
                    $KeyBox.Text = $KeyPathLocal
                }
            }
            $AppendLog.Invoke("Profile loaded: $($dlg.FileName)`r`n")
        } catch {
            $AppendLog.Invoke("[ERROR] Failed to load profile: $_`r`n")
        }
    }
})

$MenuSaveProfile.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title = "Save SSH Profile"
    $dlg.Filter = "SSH Profile (*.ssh_config.txt)|*.ssh_config.txt|All Files|*.*"
    $homeSSH = [IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), ".ssh")
    if (-not [IO.Directory]::Exists($homeSSH)) { [IO.Directory]::CreateDirectory($homeSSH) | Out-Null }
    $dlg.InitialDirectory = $homeSSH
    $baseName = $HostBox.Text
    if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = "profile" }
    $baseName = $baseName.Replace(":", "_")
    $dlg.FileName = "${baseName}_ssh_config.txt"
    if ($dlg.ShowDialog()) {
        try {
            $profilePath = $dlg.FileName
            $alias = [IO.Path]::GetFileNameWithoutExtension($profilePath)
            $alias = $alias -replace "_ssh_config$", ""
            $hostName = $HostBox.Text
            $port = $PortBox.Text
            $rootUser = $UserBox.Text
            $newUser = $NewUserBox.Text
            $profileUser = if (-not [string]::IsNullOrWhiteSpace($newUser)) { $newUser } else { $rootUser }
            $identity = $KeyBox.Text
            $content = @()
            $content += "### This ssh_config file can also be used to import this server's settings into Termius. ###"
            $content += ""
            $content += "Host $alias"
            $content += "    HostName $hostName"
            $content += "    User $profileUser"
            $content += "    Port $port"
            if (-not [string]::IsNullOrWhiteSpace($identity)) {
                $content += "    IdentityFile $identity"
            }
            $content += ""
            $content += "# CreatedOn: $(Get-Date -Format yyyy-MM-dd)"
            $keyPart = if (-not [string]::IsNullOrEmpty($identity)) { "-i $identity " } else { "" }
            $content += "# SSH login: ssh ${keyPart}${profileUser}@${hostName}"
            $content += "# SFTP: sftp ${keyPart}${profileUser}@${hostName}"
            $content | Out-File -FilePath $profilePath -Encoding ASCII
            $AppendLog.Invoke("Profile saved to: $profilePath`r`n")
        } catch {
            $AppendLog.Invoke("[ERROR] Failed to save profile: $_`r`n")
        }
    }
})

$MenuExit.Add_Click({
    $Window.Close()
})

$StartButton.Add_Click({
    $LogBox.Clear()
    try {
        $AppendLog.Invoke("=== StartButton clicked at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===`r`n")

        $serverHost = $HostBox.Text.Trim()
        $port       = $PortBox.Text.Trim()
        $user       = $UserBox.Text.Trim()
        $pass       = $PasswordBox.Password
        $keyPath    = $KeyBox.Text.Trim()
        $newUser    = $NewUserBox.Text.Trim()
        $newPass    = $NewPasswordBox.Password
        $disableRoot = $DisableRootCheck.IsChecked -and $DisableRootCheck.IsChecked.Value

        if ([string]::IsNullOrWhiteSpace($serverHost) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($port)) {
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

        Reset-HostKeysForServer -HostName $serverHost -Port $port

        $script:useKeyAuth = $false
        $script:usePasswordAuth = $false
        $authMode = ""

        if (-not [string]::IsNullOrWhiteSpace($keyPath)) {
            if (-not (Test-Path $keyPath)) {
                $AppendLog.Invoke("[ERROR] SSH key file not found: $keyPath`r`n")
                return
            }

            $ext       = [IO.Path]::GetExtension($keyPath).ToLower()
            $firstLine = (Get-Content -Path $keyPath -TotalCount 1 -ErrorAction Stop)

            $isPuttyHeader   = $firstLine -like "PuTTY-User-Key-File-*"
            $isOpenSshHeader = $firstLine -like "-----BEGIN OPENSSH PRIVATE KEY-----"
            $isPemHeader     = $firstLine -like "-----BEGIN *PRIVATE KEY-----"

            if ($ext -eq ".ppk" -or $isPuttyHeader) {
                $conversionMessage = "You selected a PuTTY (.ppk) private key.`r`n`r`n" +
                    "This type of key cannot be used directly with standard SSH tools or automation.`r`n`r`n" +
                    "Your original .ppk file will not be changed.`r`n" +
                    "The toolkit will create a new OpenSSH-compatible private key in your .ssh folder`r`n" +
                    "and use that for SSH connections.`r`n`r`n" +
                    "Do you want to convert this .ppk key now?"

                $conversionResult = [System.Windows.MessageBox]::Show(
                    $conversionMessage,
                    "Convert PuTTY Key?",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )

                if ($conversionResult -ne [System.Windows.MessageBoxResult]::Yes) {
                    $AppendLog.Invoke("You chose not to convert the PuTTY key. Key-based authentication will not be used.`r`n")
                } else {
                    $AppendLog.Invoke("Converting your PuTTY key into a standard SSH key for this setup...`r`n")

                    $puttygenPath = ""
                    try {
                        $puttygenCmd = Get-Command puttygen.exe -ErrorAction Stop
                        $puttygenPath = $puttygenCmd.Source
                    } catch {
                        $tempPg = Join-Path ([IO.Path]::GetTempPath()) "puttygen.exe"
                        if (Test-Path $tempPg) {
                            $puttygenPath = $tempPg
                        }
                    }

                    if (-not $puttygenPath) {
                        $AppendLog.Invoke("[ERROR] PuTTYgen (puttygen.exe) not found. Cannot convert .ppk to OpenSSH.`r`n")
                        return
                    }

                    $sshDir = Join-Path $HOME ".ssh"
                    if (-not (Test-Path $sshDir)) {
                        New-Item -ItemType Directory -Path $sshDir | Out-Null
                    }

                    $openSshFileName = [IO.Path]::GetFileNameWithoutExtension($keyPath) + "-openssh.key"
                    $openSshPath     = Join-Path $sshDir $openSshFileName

                    if (Test-Path $openSshPath) {
                        $existingHeader = ""
                        try {
                            $existingHeader = (Get-Content -Path $openSshPath -TotalCount 1 -ErrorAction Stop)
                        } catch {
                            $existingHeader = "<unable to read header>"
                        }

                        $msg = "An SSH private key file already exists at:`r`n`r`n" +
                               "  $openSshPath`r`n`r`n" +
                               "First line of the existing file:`r`n" +
                               "  $existingHeader`r`n`r`n" +
                               "If you overwrite this file, any other tools or servers using it may break.`r`n`r`n" +
                               "Choose an option:`r`n" +
                               "  Yes    = Overwrite this file with a new OpenSSH key derived from:`r`n" +
                               "           $keyPath`r`n" +
                               "  No     = Choose a different filename inside .ssh`r`n" +
                               "  Cancel = Abort conversion"

                        $overwriteResult = [System.Windows.MessageBox]::Show(
                            $msg,
                            "Existing SSH key detected",
                            [System.Windows.MessageBoxButton]::YesNoCancel,
                            [System.Windows.MessageBoxImage]::Warning
                        )

                        switch -Exact ($overwriteResult) {
                            ([System.Windows.MessageBoxResult]::Yes) { }
                            ([System.Windows.MessageBoxResult]::No) {
                                $saveDlg = New-Object Microsoft.Win32.SaveFileDialog
                                $saveDlg.Title            = "Save converted OpenSSH key as..."
                                $saveDlg.InitialDirectory = $sshDir
                                $saveDlg.FileName         = $openSshFileName
                                $saveDlg.Filter           = "OpenSSH Private Key (*.key;*.*)|*.key;*.*"
                                $saveResult = $saveDlg.ShowDialog()
                                if (-not $saveResult) {
                                    $AppendLog.Invoke("[INFO] User cancelled Save As dialog. Conversion cancelled.`r`n")
                                    return
                                }
                                $openSshPath = $saveDlg.FileName
                                $AppendLog.Invoke("User chose to save converted key as: $openSshPath`r`n")
                            }
                            ([System.Windows.MessageBoxResult]::Cancel) {
                                $AppendLog.Invoke("[INFO] User cancelled conversion of .ppk key. No files were changed.`r`n")
                                return
                            }
                            default {
                                $AppendLog.Invoke("[INFO] Unexpected dialog result. Conversion cancelled.`r`n")
                                return
                            }
                        }
                    }

                    $AppendLog.Invoke("Converted OpenSSH key will be saved to: $openSshPath`r`n")

                    try {
                        & $puttygenPath $keyPath -O private-openssh -o $openSshPath
                        if (-not (Test-Path $openSshPath)) {
                            $AppendLog.Invoke("[ERROR] PuTTYgen did not produce an OpenSSH key at: $openSshPath`r`n")
                            return
                        }
                        $AppendLog.Invoke("Converted .ppk to OpenSSH key: $openSshPath`r`n")
                        $keyPath = $openSshPath
                        $script:useKeyAuth = $true
                    } catch {
                        $AppendLog.Invoke("[ERROR] Failed to convert .ppk to OpenSSH via PuTTYgen: $_`r`n")
                        return
                    }
                }
            }
            elseif ($isOpenSshHeader -or $isPemHeader) {
                $AppendLog.Invoke("Valid OpenSSH/PEM private key detected - using as-is: $keyPath`r`n")
                $script:useKeyAuth = $true
            }
            else {
                $AppendLog.Invoke("[ERROR] Unsupported SSH key format in file: $keyPath`r`n")
                $AppendLog.Invoke("First line was: $firstLine`r`n")
                return
            }
        }

        $haveKey = $script:useKeyAuth
        $havePass = -not [string]::IsNullOrWhiteSpace($pass)

        if ($haveKey) {
            $authMode = "key"
            $script:usePasswordAuth = $false
        } elseif ($havePass) {
            $authMode = "password"
            $script:usePasswordAuth = $true
        } else {
            $AppendLog.Invoke("[ERROR] You must provide either an SSH key (for agent-based auth) or a password for password-only setup.`r`n")
            return
        }

        if ($authMode -eq "key") {
            $AppendLog.Invoke("Using key-based authentication for $user@$serverHost (Password is treated as key passphrase).`r`n")
            $agentOk = Ensure-SshAgentWithKey -KeyPath $keyPath -Passphrase $pass
            if (-not $agentOk) {
                $AppendLog.Invoke("[INFO] Waiting for you to complete ssh-add in the opened PowerShell window.`r`n")
                $AppendLog.Invoke("[INFO] Once ssh-add succeeds and that window is closed, click Start Setup again in this GUI.`r`n")
                return
            }
        } else {
            $AppendLog.Invoke("Using password-based authentication for $user@$serverHost.`r`n")
            if (-not (Ensure-PuttyTools)) {
                $AppendLog.Invoke("[FATAL] plink.exe is required for password-only mode. Setup cannot continue.`r`n")
                return
            }
        }

        function Run-PlinkCommand {
            param(
                [string]$remoteCmd
            )
            $args = "-batch -ssh -P $port"
            if ($script:useKeyAuth -and $authMode -eq "password") {
                $args += " -i `"$keyPath`""
            } elseif ($script:usePasswordAuth -and $pass) {
                $args += " -pw `"$($pass)`""
            }
            $args += " -l $user $serverHost `"$remoteCmd`""

            if ($args -like "* -pw *") {
                $AppendLog.Invoke("DEBUG: Using plink at '$($global:PlinkPath)' with password-based authentication (details redacted).`r`n")
            } else {
                $AppendLog.Invoke("DEBUG: Using plink at '$($global:PlinkPath)' with args: $args`r`n")
            }

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $global:PlinkPath
            $psi.Arguments = $args
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.RedirectStandardInput = $false
            $psi.CreateNoWindow = $true

            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $psi
            $proc.add_OutputDataReceived({ if ($_.Data -ne $null) { $AppendLog.Invoke("$($_.Data)`r`n") } })
            $proc.add_ErrorDataReceived({ if ($_.Data -ne $null) { $AppendLog.Invoke("[ERROR] $($_.Data)`r`n") } })
            $proc.Start() | Out-Null
            $proc.BeginOutputReadLine()
            $proc.BeginErrorReadLine()
            while (-not $proc.HasExited) {
                Start-Sleep -Milliseconds 100
            }
            $proc.WaitForExit()
            return $proc.ExitCode
        }

        if ($authMode -eq "password") {
            $initialExit = 0
            $AppendLog.Invoke("Connecting to $serverHost on port $port (password lane, performing host-key acceptance)...`r`n")
            try {
                if (-not $global:PlinkPath -or -not (Test-Path $global:PlinkPath)) {
                    $AppendLog.Invoke("[FATAL] plink.exe not found at '$($global:PlinkPath)'. Cannot perform host-key acceptance.`r`n")
                    return
                }
                $hostKeyCmd = "echo y | `"$($global:PlinkPath)`" -ssh -P $port -l $user $serverHost exit"
                $AppendLog.Invoke("DEBUG: Running automatic host-key acceptance via: cmd.exe /c $hostKeyCmd`r`n")
                $proc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $hostKeyCmd" -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
                $initialExit = $proc.ExitCode
                $AppendLog.Invoke("DEBUG: Host-key acceptance plink exit code: $initialExit`r`n")
            } catch {
                $AppendLog.Invoke("[WARN] Automatic host-key acceptance command failed: $($_.Exception.Message)`r`n")
                $initialExit = 0
            }
            if ($initialExit -ne 0) {
                $AppendLog.Invoke("[WARN] Host-key acceptance command returned exit code $initialExit. Continuing with setup.`r`n")
            } else {
                $AppendLog.Invoke("Host key accepted or already trusted for ${serverHost}:${port}.`r`n")
            }
        } else {
            $AppendLog.Invoke("Connecting to $serverHost on port $port using OpenSSH...`r`n")
        }

        $sudoPrefix = ""
        if ($user -ne "root") { $sudoPrefix = "sudo -S -p '' " }

        $sudoPassword = $null
        if ($user -ne "root") {
            if ($pass -and $authMode -eq "password") {
                $sudoPassword        = $pass
                $script:sudoPassword = $pass
            }
        }

        function Run-RemoteCommand {
            param(
                [string]$cmd,
                [bool]$checkOutput=$false
            )

            if ($user -eq "root") {
                $fullCmd = $cmd
            } else {
                if ($sudoPassword) {
                    $fullCmd = "echo `"${sudoPassword}`" | sudo -S -p '' $cmd"
                } else {
                    $fullCmd = "sudo -S -p '' $cmd"
                }
            }

            if ($authMode -eq "key") {
                if ($cmd -like "*chpasswd*") {
                    $AppendLog.Invoke("DEBUG: Running remote ssh command: [REDACTED password update command]`r`n")
                } elseif ($fullCmd -like "*sudo -S -p*") {
                    $AppendLog.Invoke("DEBUG: Running remote ssh command: [REDACTED sudo password command]`r`n")
                } else {
                    $AppendLog.Invoke("DEBUG: Running remote ssh command: $fullCmd`r`n")
                }
                $exit = Run-SshCommand -RemoteCommand $fullCmd -RemoteHost $serverHost -Port $port -User $user -KeyPath $keyPath
            } else {
                if ($fullCmd -like "*sudo -S -p*") {
                    $AppendLog.Invoke("DEBUG: Running remote plink command: [REDACTED sudo password command]`r`n")
                } else {
                    $AppendLog.Invoke("DEBUG: Running remote plink command: $fullCmd`r`n")
                }
                $exit = Run-PlinkCommand -remoteCmd $fullCmd
            }

            if ($exit -ne 0) {
                $AppendLog.Invoke("[ERROR] Remote command failed (exit code $exit): $fullCmd`r`n")
            }
            return $exit
        }

        $AppendLog.Invoke("Creating '$newUser' user...`r`n")
        $ensureUserCmd = "bash -c 'id -u $newUser >/dev/null 2>&1 || useradd -m -s /bin/bash $newUser'"
        $exitCode = Run-RemoteCommand -cmd $ensureUserCmd
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to ensure user '$newUser' exists (exit code $exitCode). Setup aborted.`r`n")
            return
        }
        $AppendLog.Invoke("User '$newUser' exists.`r`n")

        $AppendLog.Invoke("Adding '$newUser' to sudo group...`r`n")
        $exitCode = Run-RemoteCommand -cmd "usermod -aG sudo $newUser"
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to add user '$newUser' to sudo group. (Exit code $exitCode)`r`n")
            return
        } else {
            $AppendLog.Invoke("User '$newUser' is now in sudo group.`r`n")
        }

        if (-not [string]::IsNullOrWhiteSpace($newPass)) {
            $AppendLog.Invoke("Setting password for '$newUser'...`r`n")
            $rawPair = $newUser + ":" + $newPass
            $b64Pair = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($rawPair))

            $cmdSetPass = "bash -c 'echo $b64Pair | base64 -d | chpasswd'"
            $exitCode = Run-RemoteCommand -cmd $cmdSetPass

            if ($exitCode -ne 0) {
                $AppendLog.Invoke("[ERROR] Failed to set password for $newUser (Exit code $exitCode)`r`n")
            } else {
                $AppendLog.Invoke("Password set for '$newUser'.`r`n")
            }
        } else {
            $AppendLog.Invoke("No password set for '$newUser' (account may be key-only).`r`n")
        }

        $AppendLog.Invoke("Copying authorized_keys to '$newUser'...`r`n")
        $cmdCopy = 'bash -c ''mkdir -p /home/' + $newUser + '/.ssh && ' +
            'if [ -f ~/.ssh/authorized_keys ]; then cp ~/.ssh/authorized_keys /home/' + $newUser + '/.ssh/authorized_keys && chmod 600 /home/' + $newUser + '/.ssh/authorized_keys; else echo "NO_AUTH_KEYS"; fi && ' +
            'chown -R ' + $newUser + ':' + $newUser + ' /home/' + $newUser + '/.ssh && chmod 700 /home/' + $newUser + '/.ssh'''
        $exitCode = Run-RemoteCommand -cmd $cmdCopy
        if ($exitCode -ne 0) {
            if ($LogBox.Text.Contains("NO_AUTH_KEYS")) {
                $AppendLog.Invoke("[WARN] No authorized_keys found for $user; none copied for $newUser.`r`n")
            } else {
                $AppendLog.Invoke("[ERROR] Failed to copy authorized_keys (Exit code $exitCode).`r`n")
            }
        } else {
            $AppendLog.Invoke("SSH key authorized for '$newUser'.`r`n")
        }

        $AppendLog.Invoke("Testing login for new user '$newUser'...`r`n")
        $testExit = 1

        if ($authMode -eq "key" -and -not [string]::IsNullOrWhiteSpace($keyPath) -and (Test-Path $keyPath)) {
            $testCmd = "echo OK"
            $testExit = Run-SshCommand -RemoteCommand $testCmd -RemoteHost $serverHost -Port $port -User $newUser -KeyPath $keyPath
        } elseif (-not [string]::IsNullOrWhiteSpace($newPass) -and $authMode -eq "password") {
            $AppendLog.Invoke("(Using password authentication for test)`r`n")
            $testArgs = "-batch -ssh -P $port -l $newUser -pw `"$($newPass)`" `"echo OK`""
            $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                FileName = $global:PlinkPath
                Arguments = $testArgs
                UseShellExecute = $false
                RedirectStandardOutput = $true
                RedirectStandardError = $true
                CreateNoWindow = $true
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

        if ($disableRoot) {
            $AppendLog.Invoke("Disabling root SSH login and global password auth...`r`n")
            $cmdDisableRoot = @"
bash -c '
if grep -q "^[#[:space:]]*PermitRootLogin" /etc/ssh/sshd_config; then
sed -i "s/^[#[:space:]]*PermitRootLogin.*/PermitRootLogin no/" /etc/ssh/sshd_config;
else
echo "PermitRootLogin no" >> /etc/ssh/sshd_config;
fi;
if grep -q "^[#[:space:]]*PasswordAuthentication" /etc/ssh/sshd_config; then
sed -i "s/^[#[:space:]]*PasswordAuthentication.*/PasswordAuthentication no/" /etc/ssh/sshd_config;
else
echo "PasswordAuthentication no" >> /etc/ssh/sshd_config;
fi'
"@
            $exitCode = Run-RemoteCommand -cmd $cmdDisableRoot
            if ($exitCode -ne 0) {
                $AppendLog.Invoke("[ERROR] Failed to update sshd_config to disable root login and password auth.`r`n")
            } else {
                $AppendLog.Invoke("Root login and password authentication disabled in SSH config.`r`n")
                $exitCode = Run-RemoteCommand -cmd "bash -c 'which systemctl >/dev/null 2>&1 && (systemctl reload sshd || systemctl reload ssh) || service ssh reload'"
                if ($exitCode -ne 0) {
                    $AppendLog.Invoke("[WARN] SSH service reload failed. You may need to restart SSH manually for changes to take effect.`r`n")
                } else {
                    $AppendLog.Invoke("SSH service reloaded to apply new settings.`r`n")
                }
            }
        }

        $AppendLog.Invoke("Setup process complete.`r`n")
    }
    catch {
        $AppendLog.Invoke("[FATAL] Unexpected error during setup: $($_.Exception.Message)`r`n")
    }
})

$SudoOkButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $SudoOverlay.Visibility = "Collapsed"
    })
    $sudoPassInput = $SudoPassBox.Password
    if ([string]::IsNullOrEmpty($sudoPassInput)) {
        $AppendLog.Invoke("[ERROR] Sudo password was not provided. Aborting.`r`n")
        return
    }
    $AppendLog.Invoke("Sudo password received. Resuming operations...`r`n")
    $script:sudoPassword = $sudoPassInput
    $exitCode = Run-RemoteCommand -cmd "useradd -m -s /bin/bash $($NewUserBox.Text.Trim())"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to create user even after sudo password. Aborting.`r`n")
        return
    }
    $AppendLog.Invoke("User '$($NewUserBox.Text.Trim())' created successfully (after sudo password).`r`n")
    Continue-SetupAfterUserCreation
})

$SudoCancelButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $SudoOverlay.Visibility = "Collapsed"
    })
    $AppendLog.Invoke("[ERROR] Sudo password prompt canceled by user. Aborting setup.`r`n")
})

$RequireYesButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $RequireOverlay.Visibility = "Collapsed"
    })
    $AppendLog.Invoke("Disabling requiretty and resuming operations...`r`n")
    if (-not $script:sudoPassword) {
        $AppendLog.Invoke("Please enter sudo password to disable requiretty...`r`n")
        $Window.Dispatcher.Invoke([Action]{
            $SudoOverlay.Visibility = "Visible"
        })
    } else {
        Disable-RequirettyAndContinue
    }
})

$RequireNoButton.Add_Click({
    $Window.Dispatcher.Invoke([Action]{
        $RequireOverlay.Visibility = "Collapsed"
    })
    $AppendLog.Invoke("User chose not to disable requiretty. Aborting setup.`r`n")
})

function Disable-RequirettyAndContinue {
    $cmd = "sed -i 's/^Defaults\s\+requiretty/#&/' /etc/sudoers"
    $plinkArgs = "-ssh -t -P $($PortBox.Text.Trim()) -l $($UserBox.Text.Trim())"
    if ($script:useKeyAuth) {
        $plinkArgs += " -i `"$($KeyBox.Text)`""
    } elseif ($script:usePasswordAuth) {
        $plinkArgs += " -pw `"$($PasswordBox.Password)`""
    }
    $plinkArgs += " `"echo $script:sudoPassword | sudo -S -p '' sed -i 's/^Defaults\s\+requiretty/#&/' /etc/sudoers`""
    try {
        $output = & $global:PlinkPath $plinkArgs
        $AppendLog.Invoke("requiretty disabled in sudoers.`r`n")
    } catch {
        $AppendLog.Invoke("[ERROR] Failed to disable requiretty: $_`r`n")
    }
    $AppendLog.Invoke("Resuming user creation...`r`n")
    $exitCode = Run-RemoteCommand -cmd "useradd -m -s /bin/bash $($NewUserBox.Text.Trim())"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to create user even after disabling requiretty. Aborting.`r`n")
        return
    }
    $AppendLog.Invoke("User '$($NewUserBox.Text.Trim())' created successfully.`r`n")
    Continue-SetupAfterUserCreation
}

function Continue-SetupAfterUserCreation {
    $newUser = $NewUserBox.Text.Trim()
    $AppendLog.Invoke("Adding '$newUser' to sudo group...`r`n")
    $exitCode = Run-RemoteCommand -cmd "usermod -aG sudo $newUser"
    if ($exitCode -ne 0) {
        $AppendLog.Invoke("[ERROR] Failed to add user to sudo group. Aborting.`r`n")
        return
    } else {
        $AppendLog.Invoke("User '$newUser' added to sudo group.`r`n")
    }

    $newPass = $NewPasswordBox.Password
    if (-not [string]::IsNullOrEmpty($newPass)) {
        $AppendLog.Invoke("Setting password for '$newUser'...`r`n")
        $rawPair = $newUser + ":" + $newPass
        $b64Pair = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($rawPair))

        $cmdSetPass = "bash -c 'echo $b64Pair | base64 -d | chpasswd'"
        $exitCode = Run-RemoteCommand -cmd $cmdSetPass

        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to set password for $newUser.`r`n")
        } else {
            $AppendLog.Invoke("Password set for '$newUser'.`r`n")
        }
    }

    $AppendLog.Invoke("Copying authorized_keys to '$newUser'...`r`n")
    $cmdCopy = 'bash -c ''mkdir -p /home/' + $newUser + '/.ssh && ' +
    'if [ -f ~/.ssh/authorized_keys ]; then cp ~/.ssh/authorized_keys /home/' + $newUser + '/.ssh/authorized_keys && chmod 600 /home/' + $newUser + '/.ssh/authorized_keys; else echo "NO_AUTH_KEYS"; fi && ' +
    'chown -R ' + $newUser + ':' + $newUser + ' /home/' + $newUser + '/.ssh && chmod 700 /home/' + $newUser + '/.ssh'''
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

    $AppendLog.Invoke("Testing login for new user '$newUser'...`r`n")
    $testExit = 1
    $useKey = (-not [string]::IsNullOrEmpty($KeyBox.Text))

    if ($useKey -and (Test-Path $KeyBox.Text)) {
        $testArgs = "-batch -ssh -P $($PortBox.Text.Trim()) -i `"$($KeyBox.Text)`" -l $newUser `"echo OK`""
        $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = $global:PlinkPath
            Arguments = $testArgs
            UseShellExecute = $false
            RedirectStandardOutput = $true
            RedirectStandardError = $true
            CreateNoWindow = $true
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
        $testArgs = "-batch -ssh -P $($PortBox.Text.Trim()) -l $newUser -pw `"$($newPass)`" `"echo OK`""
        $psiTest = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = $global:PlinkPath
            Arguments = $testArgs
            UseShellExecute = $false
            RedirectStandardOutput = $true
            RedirectStandardError = $true
            CreateNoWindow = $true
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
    } else {
        $AppendLog.Invoke("[WARN] No key or password available to test new user login.`r`n")
        $testExit = 1
    }

    if ($testExit -eq 0) {
        $AppendLog.Invoke("New user '$newUser' login test successful.`r`n")
    } else {
        $AppendLog.Invoke("[ERROR] New user '$newUser' login test failed. Root login will not be disabled.`r`n")
        $DisableRootCheck.IsChecked = $false
    }

    if ($DisableRootCheck.IsChecked -and $DisableRootCheck.IsChecked.Value -and $testExit -eq 0) {
        $AppendLog.Invoke("Disabling root SSH login and global password auth...`r`n")
        $cmdDisableRoot = @"
bash -c '
if grep -q "^[#[:space:]]*PermitRootLogin" /etc/ssh/sshd_config; then
sed -i "s/^[#[:space:]]*PermitRootLogin.*/PermitRootLogin no/" /etc/ssh/sshd_config;
else
echo "PermitRootLogin no" >> /etc/ssh/sshd_config;
fi;
if grep -q "^[#[:space:]]*PasswordAuthentication" /etc/ssh/sshd_config; then
sed -i "s/^[#[:space:]]*PasswordAuthentication.*/PasswordAuthentication no/" /etc/ssh/sshd_config;
else
echo "PasswordAuthentication no" >> /etc/ssh/sshd_config;
fi'
"@
        $exitCode = Run-RemoteCommand -cmd $cmdDisableRoot
        if ($exitCode -ne 0) {
            $AppendLog.Invoke("[ERROR] Failed to update sshd_config for PermitRootLogin/PasswordAuthentication.`r`n")
        } else {
            $AppendLog.Invoke("Root login and password auth disabled in sshd_config.`r`n")
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

$Window.add_Closed({
    Cleanup-AgentOwnedKeys
})

$null = $Window.ShowDialog()
