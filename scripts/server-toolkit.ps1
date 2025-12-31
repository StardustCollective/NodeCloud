$scriptversion = "1.2"

try {
    $global:ServerToolkitLogPath = Join-Path $env:USERPROFILE "server-toolkit.log"
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $hostName = $Host.Name
    $psver = $PSVersionTable.PSVersion.ToString()
    $edition = $PSVersionTable.PSEdition
    $apartment = [System.Threading.Thread]::CurrentThread.ApartmentState.ToString()
    Add-Content -Path $global:ServerToolkitLogPath -Value "$ts [BOOT] Host=$hostName PSEdition=$edition PSVersion=$psver Apartment=$apartment" -Encoding UTF8 -ErrorAction SilentlyContinue
} catch {}

$script:DevMode = ($env:SERVER_TOOLKIT_DEV -eq "1")

try {
    $apt = [System.Threading.Thread]::CurrentThread.ApartmentState.ToString()
    if ($apt -ne "STA") {
        $exe = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"

        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.MyCommand.Path }

        if ($scriptPath -and (Test-Path -LiteralPath $scriptPath) -and (Test-Path -LiteralPath $exe)) {
            Start-Process -FilePath $exe -ArgumentList @(
                "-NoProfile",
                "-ExecutionPolicy","Bypass",
                "-STA",
                "-File", "`"$scriptPath`""
            ) | Out-Null
            exit
        }
    }
} catch {}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Xaml
try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue } catch {}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Server Toolkit ${scriptversion} - Stardust Collective" Height="700" Width="715" WindowStartupLocation="CenterScreen"
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

        <!-- Start button (green) palette -->
        <SolidColorBrush x:Key="StartButtonBg" Color="#47d16c" />
        <SolidColorBrush x:Key="StartButtonBgHover" Color="#58e07c" />
        <SolidColorBrush x:Key="StartButtonBgPressed" Color="#33b958" />

        <!-- Start button style -->
        <Style x:Key="StartButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource StartButtonBg}" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="10,4" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center"
                                            VerticalAlignment="Center"
                                            Margin="4,1"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource StartButtonBgHover}" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource StartButtonBgPressed}" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary (grey) button palette -->
        <SolidColorBrush x:Key="ButtonBgSecondary" Color="#3b3f46" />
        <SolidColorBrush x:Key="ButtonBgSecondaryHover" Color="#4a4f5a" />
        <SolidColorBrush x:Key="ButtonBgSecondaryPressed" Color="#2f3238" />

        <SolidColorBrush x:Key="BorderBrushDark" Color="#3b3f46" />

        <!-- Secondary button style (grey) -->
        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource ButtonBgSecondary}" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="10,4" />
            <Setter Property="HorizontalAlignment" Value="Left" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center"
                                            VerticalAlignment="Center"
                                            Margin="4,1"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource ButtonBgSecondaryHover}" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource ButtonBgSecondaryPressed}" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource LabelFg}" />
            <Setter Property="TextWrapping" Value="Wrap" />
        </Style>

        <Style TargetType="Label">
            <Setter Property="Foreground" Value="{StaticResource LabelFg}" />
            <Setter Property="Margin" Value="0,0,4,4" />
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
            <Setter Property="IsTabStop" Value="True" />
            <Setter Property="Focusable" Value="True" />
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

        <!-- Global ToggleButton theme override (dark) -->
        <Style TargetType="ToggleButton">
            <Setter Property="Background" Value="#2f3238" />
            <Setter Property="Foreground" Value="#f0f0f0" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="4" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                            VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#25272b" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#25272b" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#25272b" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.4" />
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

        <!-- Submenu item style (dropdown items) - supports hover background without breaking menu -->
        <Style x:Key="SubMenuItemStyle" TargetType="MenuItem">
            <Setter Property="Background" Value="#262a30" />
            <Setter Property="Foreground" Value="#f0f0f0" />
            <Setter Property="Padding" Value="10,4" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type MenuItem}">
                        <Grid SnapsToDevicePixels="True">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="22" />
                            </Grid.ColumnDefinitions>

                            <!-- Item surface -->
                            <Border x:Name="Bd"
                                    Grid.ColumnSpan="2"
                                    Background="{TemplateBinding Background}"
                                    CornerRadius="3"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter x:Name="HeaderHost"
                                                ContentSource="Header"
                                                RecognizesAccessKey="True"
                                                VerticalAlignment="Center"/>
                            </Border>

                            <!-- Submenu arrow (only shown when HasItems) -->
                            <TextBlock x:Name="Arrow"
                                    Grid.Column="1"
                                    Text="›"
                                    FontSize="16"
                                    Margin="0,0,8,0"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Right"
                                    Foreground="#b0b0b0"
                                    Visibility="Collapsed"/>

                            <!-- Popup for submenu -->
                            <Popup x:Name="PART_Popup"
                                Placement="Right"
                                IsOpen="{TemplateBinding IsSubmenuOpen}"
                                AllowsTransparency="True"
                                Focusable="False"
                                PopupAnimation="Fade">
                                <Border Background="#262a30"
                                        BorderBrush="#3b3f46"
                                        BorderThickness="1"
                                        CornerRadius="6"
                                        Padding="4">
                                    <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Cycle" />
                                </Border>
                            </Popup>
                        </Grid>

                        <ControlTemplate.Triggers>

                            <!-- Show arrow only when this item has children -->
                            <Trigger Property="HasItems" Value="True">
                                <Setter TargetName="Arrow" Property="Visibility" Value="Visible" />
                            </Trigger>

                            <!-- Hover highlight -->
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#3a3f4b" />
                                <Setter Property="Foreground" Value="#ffffff" />
                                <Setter TargetName="Arrow" Property="Foreground" Value="#ffffff" />
                            </Trigger>

                            <!-- Keyboard focus / selection -->
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#3a3f4b" />
                                <Setter Property="Foreground" Value="#ffffff" />
                                <Setter TargetName="Arrow" Property="Foreground" Value="#ffffff" />
                            </Trigger>

                            <!-- Disabled -->
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.45" />
                                <Setter Property="Foreground" Value="#9aa0a6" />
                                <Setter TargetName="Arrow" Property="Foreground" Value="#9aa0a6" />
                            </Trigger>

                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Themed ToolTips (match dark UI) -->
       <Style TargetType="{x:Type ToolTip}">
            <Setter Property="Background" Value="#25272b" />
            <Setter Property="Foreground" Value="#f0f0f0" />
            <Setter Property="BorderBrush" Value="#3b3f46" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Padding" Value="10,8" />
            <Setter Property="FontFamily" Value="Segoe UI" />
            <Setter Property="FontSize" Value="12" />
            <Setter Property="MaxWidth" Value="340" />

            <!-- Tooltip positioning -->
            <Setter Property="Placement" Value="Mouse" />
            <Setter Property="HorizontalOffset" Value="10" />
            <Setter Property="VerticalOffset" Value="10" />
        </Style>

        <Style TargetType="MenuItem">
            <Setter Property="Background" Value="#262a30" />
            <Setter Property="Foreground" Value="{StaticResource TextFg}" />
            <Setter Property="Padding" Value="10,4" />
        </Style>

    </Window.Resources>

    <DockPanel LastChildFill="True">
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File" ItemContainerStyle="{StaticResource SubMenuItemStyle}">
                <MenuItem Header="_Load Profile..." x:Name="MenuLoadProfile"/>
                <MenuItem Header="_Save Profile..." x:Name="MenuSaveProfile"/>
                <Separator/>
                <MenuItem Header="E_xit" x:Name="MenuExit"/>
            </MenuItem>
        </Menu>
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>  <!-- Connection -->
                <RowDefinition Height="Auto"/>  <!-- Create non-root toggle + NewUserPanel -->
                <RowDefinition Height="Auto"/>  <!-- Upload P12 toggle -->
                <RowDefinition Height="Auto"/>  <!-- P12Panel -->
                <RowDefinition Height="Auto"/>  <!-- Start -->
                <RowDefinition Height="*"/>     <!-- Log -->
            </Grid.RowDefinitions>

            <GroupBox Header="Connection" Grid.Row="0" Margin="0,0,0,10">
                <Grid Margin="10,5,10,5">

                    <!-- 4 columns: LeftLabel, LeftInput, RightLabel, RightInput
                        Browse button sits in RightInput column but aligned right -->
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="150"/>   <!-- Host width -->
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>  <!-- Right input column (UserBox + Passphrase) -->
                        <ColumnDefinition Width="*"/>     <!-- Filler space (so layouts can still breathe) -->
                    </Grid.ColumnDefinitions>

                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>  <!-- Host/Port -->
                        <RowDefinition Height="Auto"/>  <!-- User/Pass -->
                        <RowDefinition Height="Auto"/>  <!-- Key/Browse -->
                    </Grid.RowDefinitions>

                    <Label Content="Server IP / Host:" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Right"/>
                    <TextBox x:Name="HostBox" Grid.Row="0" Grid.Column="1" Margin="5,2" Width="140" HorizontalAlignment="Left"/>

                    <Label Content="SSH Port:" Grid.Row="1" Grid.Column="0" HorizontalAlignment="Right" Margin="10,0,4,0"/>
                    <TextBox x:Name="PortBox" Grid.Row="1" Grid.Column="1" Margin="5,2" Width="60" HorizontalAlignment="Left" Text="22"/>

                    <Label Content="Server Username:" Grid.Row="0" Grid.Column="2" HorizontalAlignment="Left" Margin="4,0,4,0"/>

                    <!-- Inner grid so UserBox + OpenServer button share the same row -->
                    <Grid Grid.Row="0" Grid.Column="3" Margin="0,2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <TextBox x:Name="UserBox" Grid.Column="0" Width="173" HorizontalAlignment="Left" CharacterCasing="Lower">
                            <TextBox.ToolTip>
                                <ToolTip>
                                    <TextBlock TextWrapping="Wrap" Width="300">
Ubuntu username rules:
* Must start with a lowercase letter (a-z)
* Allowed: a-z, 0-9, underscore (_), hyphen (-)
* No spaces, dots, or uppercase
* Max 32 characters
                                    </TextBlock>
                                </ToolTip>
                            </TextBox.ToolTip>
                        </TextBox>
                        <!-- Hidden until Connection inputs are valid -->
                        <Button x:Name="OpenServerButton"
                            Grid.Column="1"
                            Content="Open >_"
                            Style="{StaticResource SecondaryButtonStyle}"
                            Margin="10,0,0,0"
                            Padding="10,2"
                            Visibility="Visible"
                            Width="65"
                            IsEnabled="False"
                            Opacity="0.4"/>
                    </Grid>

                    <StackPanel Grid.Row="1" Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0,4,0">
                        <Label Content="SSH Passphrase:" HorizontalAlignment="Left" Margin="0,0,3,0"/>
                        <TextBlock>
                            <TextBlock.ToolTip>
                                <ToolTip>
                                    <TextBlock TextWrapping="Wrap" Width="260">
                                        This field is for your SSH key passphrase.
                                    </TextBlock>
                                </ToolTip>
                            </TextBlock.ToolTip>
                        </TextBlock>
                    </StackPanel>

                    <Grid Grid.Row="1" Grid.Column="3" Margin="0,2" HorizontalAlignment="Left">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <!-- Source-of-truth -->
                        <PasswordBox x:Name="PasswordBox"
                                    Grid.Column="0"
                                    Width="173"
                                    HorizontalAlignment="Left"/>

                        <!-- Reveal textbox -->
                        <TextBox x:Name="PasswordRevealBox"
                                Grid.Column="0"
                                Width="173"
                                HorizontalAlignment="Left"
                                Visibility="Collapsed"/>

                        <!-- Toggle reveal -->
                        <ToggleButton x:Name="TogglePassReveal"
                                    Grid.Column="1"
                                    Width="28"
                                    Height="24"
                                    Margin="6,0,0,0"
                                    VerticalAlignment="Center"
                                    ToolTip="Show / hide passphrase"
                                    Background="#2f3238"
                                    BorderBrush="#2f3238"
                                    BorderThickness="0"
                                    Foreground="#f0f0f0"
                                    Padding="4"
                                    Cursor="Hand">
                            <TextBlock x:Name="PassRevealIcon"
                                    Text="&#xE890;"
                                    FontFamily="Segoe MDL2 Assets"
                                    FontSize="14"
                                    Foreground="#f0f0f0"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"/>
                        </ToggleButton>
                    </Grid>
                    <Label Content="SSH Key File:" Grid.Row="2" Grid.Column="0" HorizontalAlignment="Right"/>

                    <!-- Fixed-width container so Browse can "come in" like the screenshot -->
                    <Grid Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="3" Margin="0" HorizontalAlignment="Stretch">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <TextBox x:Name="KeyBox"
                                Grid.Column="0"
                                Margin="5,2,8,2"
                                IsReadOnly="True"/>

                        <StackPanel Grid.Column="1"
                                Orientation="Horizontal"
                                Margin="0,2,8,2">

                        <Button x:Name="BrowseKeyButton"
                                Content="Browse..."
                                Padding="18,10"
                                ToolTip="Select an existing SSH private key file"/>

                        <Button x:Name="CreateKeyButton"
                                Content="+"
                                FontSize="16"
                                Width="34"
                                Height="24"
                                Margin="8,0,0,0"
                                ToolTip="Create a new SSH key pair"/>
                    </StackPanel>
                    </Grid>

                </Grid>
                </GroupBox>

            <!-- Create non-root toggle + New User panel (disabled until checked - Phase 2 wiring) -->
            <StackPanel Grid.Row="1" Margin="0,0,0,10">
                <CheckBox x:Name="CreateNonRootCheck" Content="Create non-root user" IsChecked="False" Margin="5,0,0,10"/>

                <StackPanel x:Name="NewUserPanel" IsEnabled="False">
                    <GroupBox Header="New User" Margin="0,0,0,0">
                        <StackPanel>

                            <Grid Margin="10,5,10,5">
                                <!-- 4 columns like Connection: LeftLabel, LeftInput, RightLabel, RightInput -->
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>

                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <Label Content="New Username:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center"/>
                                    <TextBox x:Name="NewUserBox" Grid.Row="0" Grid.Column="1" Margin="5,2" CharacterCasing="Lower">
                                        <TextBox.ToolTip>
                                            <ToolTip>
                                                <TextBlock TextWrapping="Wrap" Width="300">
Ubuntu username rules:
* Must start with a lowercase letter (a-z)
* Allowed: a-z, 0-9, underscore (_), hyphen (-)
* No spaces, dots, or uppercase
* Max 32 characters
                                                </TextBlock>
                                            </ToolTip>
                                        </TextBox.ToolTip>
                                    </TextBox>
                                <StackPanel Grid.Row="0" Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="10,0,4,0">
                                    <Label Content="Password:" VerticalAlignment="Center" Margin="0,0,3,0"/>
                                    <TextBlock>
                                        <TextBlock.ToolTip>
                                            <ToolTip>
                                                <TextBlock TextWrapping="Wrap" Width="260">
                                                    This password is used for the new non-root user account on the server.
                                                </TextBlock>
                                            </ToolTip>
                                        </TextBlock.ToolTip>
                                    </TextBlock>
                                </StackPanel>
                                <Grid Grid.Row="0" Grid.Column="3" Margin="0,2" HorizontalAlignment="Left">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>

                                    <!-- Source-of-truth -->
                                    <PasswordBox x:Name="NewPasswordBox"
                                                Grid.Column="0"
                                                Width="173"
                                                HorizontalAlignment="Left"/>

                                    <!-- Reveal textbox -->
                                    <TextBox x:Name="NewPasswordRevealBox"
                                            Grid.Column="0"
                                            Width="173"
                                            HorizontalAlignment="Left"
                                            Visibility="Collapsed"/>

                                    <!-- Toggle reveal -->
                                    <ToggleButton x:Name="ToggleNewPassReveal"
                                                Grid.Column="1"
                                                Width="28"
                                                Height="24"
                                                Margin="6,0,0,0"
                                                VerticalAlignment="Center"
                                                ToolTip="Show / hide password">
                                        <TextBlock x:Name="NewPassRevealIcon"
                                                Text="&#xE890;"
                                                FontFamily="Segoe MDL2 Assets"
                                                FontSize="14"
                                                Foreground="#f0f0f0"
                                                VerticalAlignment="Center"
                                                HorizontalAlignment="Center"/>
                                    </ToggleButton>
                                </Grid>
                            </Grid>

                            <StackPanel Orientation="Horizontal" Margin="5,6,0,0">
                                <CheckBox x:Name="DisableRootCheck"
                                        Content="Disable root login after setup"
                                        IsChecked="True"/>
                            </StackPanel>

                        </StackPanel>
                    </GroupBox>
                </StackPanel>
            </StackPanel>

            <!-- Upload P12 toggle -->
            <Grid Grid.Row="2" Margin="5,0,0,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <CheckBox x:Name="UploadP12Check"
                        Grid.Column="0"
                        Content="Upload P12 file"
                        IsChecked="False"/>

                <!-- Spacer -->
                <Border Grid.Column="1"/>

                <Button x:Name="BackupP12Button"
                        Grid.Column="2"
                        Content="Backup P12"
                        Style="{StaticResource SecondaryButtonStyle}"
                        Width="110"
                        IsEnabled="False"
                        Opacity="0.4"
                        ToolTip="Search this system for existing P12 and back them up safely"/>
            </Grid>

            <!-- Upload P12 area (disabled by default until UploadP12Check is checked - Phase 2 wiring) -->
            <GroupBox x:Name="P12Panel" Header="P12 Upload" Grid.Row="3" IsEnabled="False" Margin="0,0,0,10">
                <Grid Margin="10,5,10,5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Row 0: file picker -->
                    <Label Content="P12 File:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" />
                    <TextBox x:Name="P12PathBox" Grid.Row="0" Grid.Column="1" Margin="8,2" IsReadOnly="True"/>
                    <Button x:Name="BrowseP12Button" Grid.Row="0" Grid.Column="2" Content="Browse..." Margin="5,2,0,2" Padding="8,0"/>

                    <!-- Row 1: passphrase -->
                    <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                        <Label Content="P12 Passphrase:" VerticalAlignment="Center" Margin="0,0,3,0"/>
                        <TextBlock>
                            <TextBlock.ToolTip>
                                <ToolTip>
                                    <TextBlock TextWrapping="Wrap" Width="260">
                                        This passphrase unlocks your P12 locally so it can be validated and uploaded.
                                    </TextBlock>
                                </ToolTip>
                            </TextBlock.ToolTip>
                        </TextBlock>
                    </StackPanel>
                    <Grid Grid.Row="1" Grid.Column="1" Margin="8,2" HorizontalAlignment="Left">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <!-- Source-of-truth -->
                        <PasswordBox x:Name="P12PasswordBox"
                                    Grid.Column="0"
                                    Width="420"
                                    HorizontalAlignment="Left"/>

                        <!-- Reveal textbox -->
                        <TextBox x:Name="P12PasswordRevealBox"
                                Grid.Column="0"
                                Width="420"
                                HorizontalAlignment="Left"
                                Visibility="Collapsed"/>

                        <!-- Toggle reveal -->
                        <ToggleButton x:Name="ToggleP12PassReveal"
                                    Grid.Column="1"
                                    Width="28"
                                    Height="24"
                                    Margin="6,0,0,0"
                                    VerticalAlignment="Center"
                                    ToolTip="Show / hide P12 passphrase">
                            <TextBlock x:Name="P12PassRevealIcon"
                                    Text="&#xE890;"
                                    FontFamily="Segoe MDL2 Assets"
                                    FontSize="14"
                                    Foreground="#f0f0f0"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"/>
                        </ToggleButton>
                    </Grid>
                    <!-- keep column 2 empty so it aligns with Browse row -->
                </Grid>
            </GroupBox>

            <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,0,0,10">
                <Button x:Name="StartButton"
                    Content="Start Setup"
                    Width="100"
                    Height="24"
                    HorizontalAlignment="Center"
                    Style="{StaticResource StartButtonStyle}"/>
                <Button x:Name="ResetHostKeysButton" Content="Reset Host Keys" Width="130" Margin="10,0,0,0" Visibility="Collapsed"/>
            </StackPanel>

            <RichTextBox x:Name="LogBox" Grid.Row="5" Margin="0"
                        Background="#FF000000"
                        FontFamily="Consolas" FontSize="12"
                        IsReadOnly="True"
                        VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Auto"
                        BorderThickness="0">
                <FlowDocument PagePadding="0">
                    <Paragraph Margin="0"/>
                </FlowDocument>
            </RichTextBox>

        </Grid>
    </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml

try {
    $MainWindow = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    $msg = $_.Exception.ToString()

    try {
        $tmp = Join-Path $env:TEMP "server-toolkit-xaml-error.log"
        Add-Content -Path $tmp -Value ("===== XAML LOAD FAILED " + (Get-Date) + " =====`r`n" + $msg + "`r`n") -Encoding UTF8
    } catch {}

    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show(
            "XAML failed to load. The detailed error was saved to:`r`n`r`n$($env:TEMP)\server-toolkit-xaml-error.log`r`n`r`n" +
            "Top of error:`r`n" +
            ($msg.Split("`n") | Select-Object -First 12) -join "`r`n",
            "Server Toolkit - XAML Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    } catch {}

    throw
}

try {
    [AppDomain]::CurrentDomain.UnhandledException += {
        param($sender, $e)
        try {
            $ex = $e.ExceptionObject
            $msg = if ($ex) { $ex.ToString() } else { "<null exception>" }
            Add-Content -Path $global:ServerToolkitLogPath -Value ("[FATAL] AppDomain.UnhandledException:`r`n$msg`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
} catch {}

try {
    [System.Threading.Tasks.TaskScheduler]::UnobservedTaskException += {
        param($sender, $e)
        try {
            Add-Content -Path $global:ServerToolkitLogPath -Value ("[FATAL] TaskScheduler.UnobservedTaskException:`r`n" + $e.Exception.ToString() + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
        try { $e.SetObserved() } catch {}
    }
} catch {}

try {
    Add-Content -Path $global:ServerToolkitLogPath -Value ("DEBUG: MainWindow type = " + $MainWindow.GetType().FullName + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
} catch {}

if (-not ($MainWindow -is [System.Windows.Window])) {
    throw "XAML did not load a Window. Type is: $($MainWindow.GetType().FullName)"
}

$HostBox        = $MainWindow.FindName("HostBox")
$PortBox        = $MainWindow.FindName("PortBox")
$UserBox        = $MainWindow.FindName("UserBox")
$OpenServerButton = $MainWindow.FindName("OpenServerButton")
$PasswordBox    = $MainWindow.FindName("PasswordBox")

$PasswordRevealBox = $MainWindow.FindName("PasswordRevealBox")
$TogglePassReveal  = $MainWindow.FindName("TogglePassReveal")
$PassRevealIcon    = $MainWindow.FindName("PassRevealIcon")

$NewPasswordRevealBox = $MainWindow.FindName("NewPasswordRevealBox")
$ToggleNewPassReveal  = $MainWindow.FindName("ToggleNewPassReveal")
$NewPassRevealIcon    = $MainWindow.FindName("NewPassRevealIcon")

$P12PasswordRevealBox = $MainWindow.FindName("P12PasswordRevealBox")
$ToggleP12PassReveal  = $MainWindow.FindName("ToggleP12PassReveal")
$P12PassRevealIcon    = $MainWindow.FindName("P12PassRevealIcon")

$KeyBox         = $MainWindow.FindName("KeyBox")
$BrowseKeyButton= $MainWindow.FindName("BrowseKeyButton")
$CreateKeyButton = $MainWindow.FindName("CreateKeyButton")
$NewUserBox     = $MainWindow.FindName("NewUserBox")
$NewPasswordBox = $MainWindow.FindName("NewPasswordBox")
$DisableRootCheck    = $MainWindow.FindName("DisableRootCheck")

$CreateNonRootCheck  = $MainWindow.FindName("CreateNonRootCheck")
$NewUserPanel        = $MainWindow.FindName("NewUserPanel")

$UploadP12Check      = $MainWindow.FindName("UploadP12Check")
$P12Panel            = $MainWindow.FindName("P12Panel")
$P12PathBox          = $MainWindow.FindName("P12PathBox")
$P12PasswordBox      = $MainWindow.FindName("P12PasswordBox")
$BrowseP12Button     = $MainWindow.FindName("BrowseP12Button")
$BackupP12Button     = $MainWindow.FindName("BackupP12Button")

$StartButton    = $MainWindow.FindName("StartButton")
$ResetHostKeysButton = $MainWindow.FindName("ResetHostKeysButton")
$LogBox         = $MainWindow.FindName("LogBox")
$script:UiDispatcher = $MainWindow.Dispatcher
$UiDispatcher = $script:UiDispatcher
$MenuLoadProfile  = $MainWindow.FindName("MenuLoadProfile")
$MenuSaveProfile  = $MainWindow.FindName("MenuSaveProfile")
$MenuExit         = $MainWindow.FindName("MenuExit")

$global:ServerToolkitLogPath = Join-Path $env:USERPROFILE "server-toolkit.log"

function Write-SharedLog {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$Value
    )
    try {
        $dir = Split-Path -Parent $Path
        if (-not [IO.Directory]::Exists($dir)) { [IO.Directory]::CreateDirectory($dir) | Out-Null }

        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
        try {
            $fs.Write($bytes, 0, $bytes.Length)
            $fs.Flush()
        } finally {
            $fs.Dispose()
        }
    } catch {}
}

try {
    if (Test-Path $global:ServerToolkitLogPath) {
        Remove-Item -LiteralPath $global:ServerToolkitLogPath -Force -ErrorAction SilentlyContinue
    }
} catch {}

try {
    Write-SharedLog -Path $global:ServerToolkitLogPath -Value ("===== Server Toolkit Log Started: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " =====`r`n")
} catch {}

$script:ApplyUiEnabledAction = [Action[bool]]{
    param([bool]$En)
    try {
        if ($HostBox)         { $HostBox.IsEnabled = $En }
        if ($PortBox)         { $PortBox.IsEnabled = $En }
        if ($UserBox)         { $UserBox.IsEnabled = $En }
        if ($PasswordBox)     { $PasswordBox.IsEnabled = $En }
        if ($BrowseKeyButton) { $BrowseKeyButton.IsEnabled = $En }

        if ($NewUserBox)      { $NewUserBox.IsEnabled = $En }
        if ($NewPasswordBox)  { $NewPasswordBox.IsEnabled = $En }

        if ($DisableRootCheck){ $DisableRootCheck.IsEnabled = $En }

        if ($CreateNonRootCheck) { $CreateNonRootCheck.IsEnabled = $En }
        if ($UploadP12Check)     { $UploadP12Check.IsEnabled     = $En }
        if ($BrowseP12Button)    { $BrowseP12Button.IsEnabled    = $En }
        if ($BackupP12Button)    { $BackupP12Button.IsEnabled    = $En }
        if ($ResetHostKeysButton){ $ResetHostKeysButton.IsEnabled = $En }

        if ($LogBox)          { $LogBox.IsEnabled = $true }
    } catch {
        try {
            Add-Content -Path $global:ServerToolkitLogPath `
                -Value ("[UIERR] ApplyUiEnabledAction failed: " + $_.Exception.ToString() + "`r`n") `
                -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
}

$script:SetUIEnabledSB = {
    param([bool]$Enabled)
    try {
        if ($script:UiDispatcher -and -not $script:UiDispatcher.CheckAccess()) {
            $null = $script:UiDispatcher.BeginInvoke($script:ApplyUiEnabledAction, [object[]]@($Enabled))
        } else {
            $script:ApplyUiEnabledAction.Invoke($Enabled)
        }
    } catch {
        try {
            Add-Content -Path $global:ServerToolkitLogPath `
                -Value ("[UIERR] SetUIEnabledSB failed: " + $_.Exception.ToString() + "`r`n") `
                -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
}

function Update-StartButtonState {
    try {
        if (-not $StartButton) { return }

        $createChecked = $false
        $p12Checked    = $false

        try { $createChecked = ($CreateNonRootCheck -and ($CreateNonRootCheck.IsChecked -eq $true)) } catch {}
        try { $p12Checked    = ($UploadP12Check   -and ($UploadP12Check.IsChecked   -eq $true)) } catch {}

        $canStart = ($createChecked -or $p12Checked)

        if (-not $script:IsBackgroundRunning) {
            $StartButton.IsEnabled = $canStart
        }
    } catch {}
}

$script:UnlockUIAction = [Action]{
    try { $script:IsBackgroundRunning = $false } catch {}
    try { Stop-LogTail } catch {}
    try { & $script:SetUIEnabledSB $true } catch {}
    try { if ($StartButton) { $StartButton.IsEnabled = $true } } catch {}
    try { if ($StartButton) { $StartButton.Content = "Start Setup" } } catch {}
    try {
        Add-Content -Path $global:ServerToolkitLogPath `
            -Value ("UI_UNLOCK_ACTION_RAN " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") `
            -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

$script:SetUsernameAction = [Action[string]]{
    param([string]$u)

    try {
        if (-not $UserBox) { return }
        if ([string]::IsNullOrWhiteSpace($u)) { return }

        $safe = Sanitize-UbuntuUsername -Text ([string]$u)

        if ($script:UiDispatcher -and -not $script:UiDispatcher.CheckAccess()) {
            $null = $script:UiDispatcher.BeginInvoke([Action]{
                try { if ($UserBox) { $UserBox.Text = $safe } } catch {}
            })
        } else {
            try { $UserBox.Text = $safe } catch {}
        }
    } catch {}
}

$script:ShowMsgAction = [Action[string,string,[System.Windows.MessageBoxImage]]]{
    param([string]$msg, [string]$title, [System.Windows.MessageBoxImage]$icon)
    try {
        [System.Windows.MessageBox]::Show(
            $MainWindow,
            $msg,
            $title,
            [System.Windows.MessageBoxButton]::OK,
            $icon
        ) | Out-Null
    } catch {}
}

$script:UserExistsPromptFunc = [Func[string,string]]{
    param([string]$u)

    $msg = "The user '$u' already exists on this server.`r`n`r`n" +
           "What would you like to do?`r`n`r`n" +
           "YES  = Continue and VERIFY / FIX the existing user`r`n" +
           "NO   = Cancel setup (no changes will be made)"

    $res = [System.Windows.MessageBox]::Show(
        $msg,
        "User Already Exists",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )

    if ($res -eq [System.Windows.MessageBoxResult]::Yes) { "continue" } else { "cancel" }
}

$script:P12OverwritePromptFunc = [Func[string,string,string]]{
    param(
        [string]$RemotePath,
        [string]$LocalPath
    )

    try {
        return (Show-P12OverwriteDialog -RemotePath $RemotePath -LocalPath $LocalPath)
    } catch {
        return "skip"
    }

    param(
        [string]$RemotePath,
        [string]$LocalPath
    )

    $rp = $RemotePath
    if ([string]::IsNullOrWhiteSpace($rp)) { $rp = "(unknown)" }

    $lp = $LocalPath
    if ([string]::IsNullOrWhiteSpace($lp)) { $lp = "(unknown)" }

    $msg = @()
    $msg += "A P12 file already exists on the server at:"
    $msg += "  $rp"
    $msg += ""
    $msg += "Local file selected:"
    $msg += "  $lp"
    $msg += ""
    $msg += "Choose what to do:"
    $msg += "YES    = Overwrite the existing remote file"
    $msg += "NO     = Make a backup, then overwrite"
    $msg += "CANCEL = Skip P12 handling and continue setup"
    $msgText = ($msg -join "`r`n")

    $res = [System.Windows.MessageBox]::Show(
        $MainWindow,
        $msgText,
        "P12 Already Exists",
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Warning
    )

    if ($res -eq [System.Windows.MessageBoxResult]::Yes) { return "overwrite" }
    if ($res -eq [System.Windows.MessageBoxResult]::No)  { return "backup" }
    return "skip"
}

try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue } catch {}

$script:OfferSaveProfileFunc = [Func[string,string,string,string,string,int,string]]{
    param(
        [string]$defaultProfileName,
        [string]$serverHost,
        [string]$sshPort,
        [string]$username,
        [string]$identityFile,
        [int]$PromptFlag = 0
    )

    try {
        $name = $defaultProfileName
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $serverHost }

        foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
            $name = $name.Replace($c, '_')
        }

        $sshDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) ".ssh"
        if (-not (Test-Path $sshDir)) { New-Item -ItemType Directory -Path $sshDir | Out-Null }

        $target = Join-Path $sshDir ("{0}_ssh_config.txt" -f $name)

        $dt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

        $preview = @()
        $preview += "### This ssh_config file can also be used to import this server's settings into Termius. ###"
        $preview += ""
        $preview += "Host $name"
        $preview += "    HostName $serverHost"
        $preview += "    User $username"
        $preview += "    Port $sshPort"
        if (-not [string]::IsNullOrWhiteSpace($identityFile)) {
            $preview += "    IdentityFile $identityFile"
        }
        $preview += ""
        $preview += "# CreatedOn: $dt"
        $keyPart = if (-not [string]::IsNullOrWhiteSpace($identityFile)) { "-i $identityFile " } else { "" }
        $preview += "# SSH login: ssh ${keyPart}${username}@${serverHost}"
        $preview += "# SFTP: sftp ${keyPart}${username}@${serverHost}"

        $msg = @()

        if ($PromptFlag -eq 1) {
            $msg += "Connection verified & setup complete."
        } else {
            $msg += "Connection verified (no changes)."
        }

        $msg += ""
        $msg += "Would you like to save this Connection Profile?"

        $msg += ""
        $msg += "Settings that will be saved:"
        $msg += "HostName $serverHost"
        $msg += "User $username"
        $msg += "Port $sshPort"
        if (-not [string]::IsNullOrWhiteSpace($identityFile)) {
            $msg += "IdentityFile $identityFile"
        } else {
            $msg += "IdentityFile (none)"
        }
        $msg += ""
        $msg += "Save this connection profile?"

        $pickedName = Show-SaveProfileDialog `
            -ServerHost $serverHost `
            -SshPort $sshPort `
            -Username $username `
            -IdentityFile $identityFile `
            -PromptFlag $PromptFlag

        if ([string]::IsNullOrWhiteSpace($pickedName)) {
            return "SKIPPED"
        }

        $name2 = $pickedName

        foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
            $name2 = $name2.Replace($c, '_')
        }

        $target2 = Join-Path $sshDir ("{0}_ssh_config.txt" -f $name2)

        $final = @()
        $final += "### This ssh_config file can also be used to import this server's settings into Termius. ###"
        $final += ""
        $final += "Host $name2"
        $final += "    HostName $serverHost"
        $final += "    User $username"
        $final += "    Port $sshPort"
        if (-not [string]::IsNullOrWhiteSpace($identityFile)) {
            $final += "    IdentityFile $identityFile"
        }
        $final += ""
        $final += "# CreatedOn: $dt"
        $keyPart = if (-not [string]::IsNullOrWhiteSpace($identityFile)) { "-i $identityFile " } else { "" }
        $final += "# SSH login: ssh ${keyPart}${username}@${serverHost}"
        $final += "# SFTP: sftp ${keyPart}${username}@${serverHost}"

        $final | Out-File -FilePath $target2 -Encoding ASCII -Force

        return "SAVED:$target2"
    }
    catch {
        return "ERROR:" + $_.Exception.Message
    }
}

function Get-ShortNodeId {
    param([string]$NodeId)

    if ($NodeId -match '^[0-9a-f]{128}$') {
        return ($NodeId.Substring(0,8) + "-" + $NodeId.Substring(120,8))
    }
    return ""
}

function Get-NodeIdFromP12Local {
    param(
        [Parameter(Mandatory=$true)][string]$P12Path,
        [Parameter(Mandatory=$true)][string]$P12Pass
    )

    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import(
            $P12Path,
            $P12Pass,
            [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
        )

        $ecdsa = $null
        try {
            $ecdsa = [System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPrivateKey($cert)
        } catch {
            $ecdsa = $null
        }

        if (-not $ecdsa) { return "" }

        $pub = $ecdsa.ExportParameters($false)

        if ($pub.Q.X -and $pub.Q.Y -and $pub.Q.X.Length -eq 32 -and $pub.Q.Y.Length -eq 32) {
            $bytes = New-Object byte[] 64
            [Array]::Copy($pub.Q.X, 0, $bytes, 0, 32)
            [Array]::Copy($pub.Q.Y, 0, $bytes, 32, 32)

            return ([BitConverter]::ToString($bytes) -replace '-', '').ToLowerInvariant()
        }

        return ""
    } catch {
        return ""
    }
}

function Protect-LogLine {
    param([string]$Line)

    if ($null -eq $Line) { return "" }

    $s = $Line
    $s = [regex]::Replace($s, '(?i)(\s-pw\s+)(?:"[^"]*"|''[^'']*''|\S+)', '$1"********"')
    $s = [regex]::Replace($s, '(?is)(\b(?:echo|printf)\s+)(["''])(.*?)(\2\s*\|\s*sudo\s+-S\b)', '$1$2********$2 | sudo -S')
    $s = [regex]::Replace($s, '(?is)(\becho\s+)(\\?"?)(.*?)(\\?"?\s*\|\s*sudo\s+-S\b)', '$1"********" | sudo -S')
    $s = [regex]::Replace($s, '(?i)\b(password|passphrase|token|secret)\s*=\s*([^\s;]+)', '$1=********')
    $s = [regex]::Replace($s, '(?i)\b([A-Za-z0-9+/]{40,}={0,2})\b', '********')
    if ($s -match '(?i)\bsudo\s+-S\b' -or $s -match '(?i)\bchpasswd\b' -or $s -match '(?i)\bbase64\s+-d\b') {
        return $s
    }

    return $s
}

$global:GuiLogLevels = @("WARN","ERROR","FATAL")
$global:GuiLogRegex  = "^\s*(?:\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}\s+)?\[(INFO|WARN|ERROR|FATAL)\]"
$global:GuiShowDebug = $false
$global:GuiStripTsRegex = "^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}\s+"

$script:AppendLog = [Action[string]]{
    param([string]$text)

    $guiOnly = $false
        try {
            if ($text -and $text.StartsWith("__GUIONLY__")) {
                $guiOnly = $true

                $text = $text.Substring(11)
            }
        } catch {}

    $uiAlive = $false
    try {
        if ($script:UiDispatcher -and -not $script:UiDispatcher.HasShutdownStarted -and -not $script:UiDispatcher.HasShutdownFinished) {
            $uiAlive = $true
        }
    } catch { $uiAlive = $false }

    if ($text -match '(?i)(\s-pw\s+|sudo\s+-S\b|chpasswd\b|base64\s+-d\b|\btoken\s*=|\bsecret\s*=|\bpassword\s*=|\bpassphrase\s*=)') {
        try {
            $text = Protect-LogLine -Line $text
        } catch {
            $text = "[WARN] Sensitive output suppressed."
        }
    }

    try {
        $guiReady = ($uiAlive -and $GuiLogLevels -and $UiDispatcher -and $LogBox)

        $stamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
        if ($null -eq $text) { $text = "" }

        $line = $text

        try { $line = Protect-LogLine -Line $line } catch {}

        if ($line -notmatch "(\r?\n)$") { $line += "`r`n" }

        if ($line -notmatch "^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}\s+") {
            $line = "$stamp $line"
        }

    try {
        if (-not $script:SuppressFileWrite -and -not $guiOnly) {
            $lp = $global:ServerToolkitLogPath
            if (-not [string]::IsNullOrWhiteSpace($lp)) {
                Write-SharedLog -Path $lp -Value $line
            }
        }
    } catch {}

    try {
        if ($guiReady) {

            $showInGui = $true
            $lvl = $null

            $m = [regex]::Match($line, $GuiLogRegex)
            if ($m.Success -and $m.Groups.Count -gt 1) {
                $lvl = $m.Groups[1].Value
            } else {
                if ($line -match "^\s*\[(WARN|ERROR|FATAL)\]") {
                    $lvl = $matches[1]
                }
            }

            if (-not $GuiShowDebug) {

                if ($lvl -in @("WARN","ERROR","FATAL")) {
                    $showInGui = $true
                }
                elseif ($line -match "(?i)\b(CLEANUP:|OWNED_KEYS:|PROFILE_PROMPT_CHECK)\b") {
                    $showInGui = $false
                }
                elseif ($line -match "TASKWATCH_") {
                    $showInGui = $false
                }
                elseif ($line -match "Node ID extraction|Short ID|Connecting|Testing|switching|Root login failed|permission denied|Creating|Adding|Setting password|Copying|Testing login|Disabling root|Upload P12|Uploading P12|P12 uploaded|P12 backup|P12 installed|P12 successfully|Setup complete|complete|Profile saved") {
                    $showInGui = $true
                }
                else {
                    $showInGui = $false
                }
            }

            if ($showInGui) {
                $guiLine = ($line -replace $GuiStripTsRegex, "")

                if ($guiLine -match '(?i)\bCLEANUP:') {
                    $showInGui = $false
                }

                if (-not $showInGui) { return }

                $level = "INFO"
                try {
                    $m2 = [regex]::Match($guiLine, '^\[(INFO|WARN|ERROR|FATAL)\]')
                    if ($m2.Success -and $m2.Groups.Count -gt 1) {
                        $level = $m2.Groups[1].Value
                    }
                } catch { $level = "INFO" }

                $guiLine = ($guiLine -replace '^\[(INFO|WARN|ERROR|FATAL)\]\s*', '')

                $guiLine = $guiLine `
                    -replace 'Disabling root SSH login and global password auth.*', 'Securing SSH configuration...' `
                    -replace 'Uploading P12 via SCP as .*', 'Uploading P12 file...' `
                    -replace 'Upload P12 enabled\. Starting secure upload.*', 'Preparing P12 upload...' `
                    -replace 'Adding .* to sudo group.*', 'Granting sudo access...' `
                    -replace 'Setting password for .*', 'Setting user password...' `
                    -replace 'Copying authorized_keys.*', 'Configuring SSH access...' `
                    -replace 'Setup complete.*', 'Setup complete.'

                $uiWrite = [Action[object,string,string]]{
                    param($rtb, $txt, $lvl)
                    try {
                        if (-not $rtb) { return }

                            try {
                                if ($rtb.Dispatcher.HasShutdownStarted -or $rtb.Dispatcher.HasShutdownFinished) { return }
                                if (-not $rtb.IsLoaded) { return }
                            } catch { return }

                        $brush = [System.Windows.Media.Brushes]::LightGray
                        switch ($lvl) {
                            "INFO"  { $brush = [System.Windows.Media.Brushes]::LightGray }
                            "WARN"  { $brush = [System.Windows.Media.Brushes]::Khaki }
                            "ERROR" { $brush = [System.Windows.Media.Brushes]::OrangeRed }
                            "FATAL" { $brush = [System.Windows.Media.Brushes]::Red }
                            default { $brush = [System.Windows.Media.Brushes]::LightGray }
                        }

                        $doc = $rtb.Document
                        if (-not $doc.Blocks.FirstBlock) {
                            $p = New-Object System.Windows.Documents.Paragraph
                            $p.Margin = [System.Windows.Thickness]::new(0)
                            $doc.Blocks.Add($p)
                        }

                        $run = New-Object System.Windows.Documents.Run($txt)
                        $run.Foreground = $brush

                        $para = $doc.Blocks.LastBlock
                        if (-not ($para -is [System.Windows.Documents.Paragraph])) {
                            $para = New-Object System.Windows.Documents.Paragraph
                            $para.Margin = [System.Windows.Thickness]::new(0)
                            $doc.Blocks.Add($para)
                        }
                        $para.Inlines.Add($run)

                        $rtb.ScrollToEnd()
                    } catch {
                        try {
                            $s2 = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
                            Add-Content -Path $global:ServerToolkitLogPath -Value ("$s2 [UIERR] UI write failed: " + $_.Exception.Message + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
                        } catch {}
                    }
                }

                if ($UiDispatcher.CheckAccess()) {
                    $uiWrite.Invoke($LogBox, $guiLine, $level)
                } else {
                    $null = $UiDispatcher.BeginInvoke($uiWrite, [object[]]@($LogBox, $guiLine, $level))
                }
            }
        }
    } catch {
        try {
            $s = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
            Add-Content -Path $global:ServerToolkitLogPath -Value ("$s [UIERR] Dispatcher failure: " + $_.Exception.Message + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }

    } catch {
        try {
            Add-Content -Path $global:ServerToolkitLogPath -Value ("[LOGFATAL] AppendLog crashed:`r`n" + $_.Exception.ToString() + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }
}

$AppendLog = $script:AppendLog

if (-not $script:SshSessions) {
    $script:SshSessions = @{}
}

function Update-BackupP12ButtonState {
    try {
        if (-not $BackupP12Button) { return }

        $h = ""
        $p = ""
        $u = ""

        try { $h = [string]$HostBox.Text } catch { $h = "" }
        try { $p = [string]$PortBox.Text } catch { $p = "" }
        try { $u = [string]$UserBox.Text } catch { $u = "" }

        $hostOk = -not [string]::IsNullOrWhiteSpace($h)
        $portOk = ($p.Trim() -match '^\d+$')
        $userOk = -not [string]::IsNullOrWhiteSpace($u)

        $uploadChecked = $false
        try { $uploadChecked = ($UploadP12Check -and ($UploadP12Check.IsChecked -eq $true)) } catch { $uploadChecked = $false }

        $enable = ($hostOk -and $portOk -and $userOk -and (-not $uploadChecked))

        $BackupP12Button.IsEnabled = $enable
        $BackupP12Button.Opacity   = if ($enable) { 1.0 } else { 0.4 }
    } catch {}
}

function Safe-UpdateBackupP12ButtonState {
    try { Update-BackupP12ButtonState } catch {}
}

function Update-OpenServerButtonVisibility {
    if (-not $OpenServerButton) { return }

    $h = ""; $p=""; $u=""; $k=""
    try { $h = ($HostBox.Text).Trim() } catch {}
    try { $p = ($PortBox.Text).Trim() } catch {}
    try { $u = ($UserBox.Text).Trim() } catch {}
    try { $k = ($KeyBox.Text).Trim() } catch {}

    $portOk = ($p -match '^\d+$')
    $hasBasics = (-not [string]::IsNullOrWhiteSpace($h)) -and $portOk -and (-not [string]::IsNullOrWhiteSpace($u))

    $canOpen = $hasBasics -and (
        (-not [string]::IsNullOrWhiteSpace($k)) -or $true
    )

    try {
        $OpenServerButton.IsEnabled = $canOpen
        $OpenServerButton.Opacity   = if ($canOpen) { 1.0 } else { 0.4 }

        if ($canOpen) {
            $OpenServerButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#000000")
        } else {
            $OpenServerButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3b3f46")
        }

    } catch {}
}

function Get-SshExePath {
    $winSsh = Join-Path $env:WINDIR "System32\OpenSSH\ssh.exe"
    if (Test-Path -LiteralPath $winSsh) { return $winSsh }
    try { return (Get-Command ssh -ErrorAction Stop).Source } catch { return $null }
}

function Get-ScpExePath {
    $winScp = Join-Path $env:WINDIR "System32\OpenSSH\scp.exe"
    if (Test-Path -LiteralPath $winScp) { return $winScp }
    try { return (Get-Command scp -ErrorAction Stop).Source } catch { return $null }
}

function Get-KnownHostsPathSafe {
    try {
        $sshDir = Join-Path $env:USERPROFILE ".ssh"
        if (-not (Test-Path -LiteralPath $sshDir)) { New-Item -ItemType Directory -Path $sshDir | Out-Null }
        $kh = Join-Path $sshDir "known_hosts"
        if (-not (Test-Path -LiteralPath $kh)) { New-Item -ItemType File -Path $kh -Force | Out-Null }
        return $kh
    } catch {
        return (Join-Path $env:USERPROFILE ".ssh\known_hosts")
    }
}

if (-not $script:SudoPwCache) { $script:SudoPwCache = @{} }

function Get-RemoteKey {
    param([string]$RemoteHost,[string]$Port,[string]$User)
    return ("{0}:{1}:{2}" -f $RemoteHost,$Port,$User)
}

function Show-RemoteUserPasswordDialog {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteUser,
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port
    )

    $bc = [System.Windows.Media.BrushConverter]::new()

    $w = New-Object System.Windows.Window
    $w.Title = "Password Required"
    $w.Width = 760
    $w.Height = 380
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = $bc.ConvertFromString("#1e1f22")
    $w.Foreground = $bc.ConvertFromString("#f0f0f0")
    $w.FontFamily = "Segoe UI"

    try {
        if ($MainWindow -and $MainWindow.Resources) {
            if (-not $w.Resources) { $w.Resources = New-Object System.Windows.ResourceDictionary }
            try { $w.Resources.MergedDictionaries.Add($MainWindow.Resources) | Out-Null } catch {}
        }
    } catch {}

    function _TryStyle([string]$key) {
        try { return $MainWindow.FindResource($key) } catch {}
        try { return $w.FindResource($key) } catch {}
        return $null
    }

    $PrimaryBtnStyle   = _TryStyle "StartButtonStyle"
    $SecondaryBtnStyle = _TryStyle "SecondaryButtonStyle"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "Admin permissions needed to access protected folders"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Margin = "0,0,0,12"
    $root.Children.Add($hdr) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdr,0)

    $panel = New-Object System.Windows.Controls.Border
    $panel.Background = $bc.ConvertFromString("#25272b")
    $panel.BorderBrush = $bc.ConvertFromString("#3b3f46")
    $panel.BorderThickness = "1"
    $panel.CornerRadius = "6"
    $panel.Padding = "14"
    [System.Windows.Controls.Grid]::SetRow($panel,1)

    $stack = New-Object System.Windows.Controls.StackPanel

    $info = New-Object System.Windows.Controls.TextBlock
    $info.TextWrapping = "Wrap"
    $info.Margin = "0,0,0,12"
    $info.Text = "To search / and to stage protected files safely, the toolkit may need to run sudo as:`r`n`r`n  $RemoteUser@${RemoteHost}:$Port`r`n`r`nEnter the password for that user. This password is NOT saved to disk."
    $stack.Children.Add($info) | Out-Null

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = "Password:"
    $lbl.FontWeight = "SemiBold"
    $lbl.Margin = "0,0,0,6"
    $stack.Children.Add($lbl) | Out-Null

    $row = New-Object System.Windows.Controls.Grid
    $row.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="*" })) | Out-Null
    $row.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto" })) | Out-Null

    $pw = New-Object System.Windows.Controls.PasswordBox
    $pw.Height = 28
    $pw.Padding = "6,3"
    $pw.Width = 540
    $row.Children.Add($pw) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($pw,0)

    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Height = 28
    $tb.Padding = "6,3"
    $tb.Width = 540
    $tb.Visibility = "Collapsed"
    $row.Children.Add($tb) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($tb,0)

    $toggle = New-Object System.Windows.Controls.Primitives.ToggleButton
    $toggle.Width = 34
    $toggle.Height = 28
    $toggle.Margin = "8,0,0,0"
    $toggle.ToolTip = "Show / hide password"

    $eye = New-Object System.Windows.Controls.TextBlock
    $eye.FontFamily = "Segoe MDL2 Assets"
    $eye.Text = [char]0xE890
    $eye.FontSize = 14
    $eye.VerticalAlignment = "Center"
    $eye.HorizontalAlignment = "Center"
    $toggle.Content = $eye

    $toggle.Add_Checked({
        try { $tb.Text = $pw.Password } catch {}
        try { $pw.Visibility = "Collapsed" } catch {}
        try { $tb.Visibility = "Visible" } catch {}
        try { $eye.Text = [char]0xE72E } catch {}
    })
    $toggle.Add_Unchecked({
        try { $pw.Password = $tb.Text } catch {}
        try { $tb.Visibility = "Collapsed" } catch {}
        try { $pw.Visibility = "Visible" } catch {}
        try { $eye.Text = [char]0xE890 } catch {}
    })

    $row.Children.Add($toggle) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($toggle,1)

    $stack.Children.Add($row) | Out-Null
    $panel.Child = $stack
    $root.Children.Add($panel) | Out-Null

    $result = [hashtable]::Synchronized(@{ Ok=$false; Value="" })

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"
    [System.Windows.Controls.Grid]::SetRow($btnRow,2)

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"
    if ($SecondaryBtnStyle) { $btnCancel.Style = $SecondaryBtnStyle }
    $btnCancel.Margin = "0,0,10,0"
    $btnCancel.Add_Click({ $result.Ok=$false; $result.Value=""; $w.Close() })
    $btnRow.Children.Add($btnCancel) | Out-Null

    $btnContinue = New-Object System.Windows.Controls.Button
    $btnContinue.Content = "Continue"
    if ($PrimaryBtnStyle) { $btnContinue.Style = $PrimaryBtnStyle }
    $btnContinue.Add_Click({
        $val = ""
        try { $val = if ($pw.Visibility -eq "Visible") { $pw.Password } else { $tb.Text } } catch { $val = "" }
        if ([string]::IsNullOrWhiteSpace($val)) {
            [System.Windows.MessageBox]::Show($w, "Password cannot be empty.", "Password Required",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $result.Ok=$true
        $result.Value=$val
        $w.Close()
    })
    $btnRow.Children.Add($btnContinue) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    $w.Content = $root
    try { $w.Add_ContentRendered({ try { $pw.Focus() } catch {} }) | Out-Null } catch {}
    $null = $w.ShowDialog()

    if ($result.Ok -eq $true) { return $result.Value }
    return ""
}

function Get-SudoPasswordForRemoteUser {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][string]$RemoteUser
    )

    $key = Get-RemoteKey -RemoteHost $RemoteHost -Port $Port -User $RemoteUser

    try {
        if ($script:SudoPwCache.ContainsKey($key)) {
            $v = [string]$script:SudoPwCache[$key]
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
        }
    } catch {}

    try {
        $nu = ""
        $np = ""
        try { if ($NewUserBox) { $nu = ([string]$NewUserBox.Text).Trim() } } catch { $nu = "" }
        try { if ($NewPasswordBox) { $np = [string]$NewPasswordBox.Password } } catch { $np = "" }

        if (-not [string]::IsNullOrWhiteSpace($nu) -and ($nu -eq $RemoteUser) -and (-not [string]::IsNullOrWhiteSpace($np))) {
            $script:SudoPwCache[$key] = $np
            return $np
        }
    } catch {}

    $pw = Show-RemoteUserPasswordDialog -RemoteUser $RemoteUser -RemoteHost $RemoteHost -Port $Port
    if (-not [string]::IsNullOrWhiteSpace($pw)) {
        try { $script:SudoPwCache[$key] = $pw } catch {}
        return $pw
    }

    return ""
}

function Find-RemoteP12FilesFast {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][string]$User,
        [string]$KeyPath,
        [string]$SudoPassword = ""
    )

    $sshExe = Get-SshExePath
    if (-not $sshExe) {
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] BackupP12: ssh.exe not found.`r`n") } } catch {}
        return @()
    }

    $kh = Get-KnownHostsPathSafe

    $plainCmd = @"
bash -lc 'set +e;
find / -maxdepth 5 \( -type d -name incremental_snapshot -prune \) -o \( -type f -name "*.p12" -print \) 2>/dev/null | sed "s/\r//g"'
"@.Trim()

    $sudoCmd = @"
bash -lc 'set +e;
sudo -S -p """" find / -maxdepth 5 \( -type d -name incremental_snapshot -prune \) -o \( -type f -name "*.p12" -print \) 2>/dev/null | sed "s/\r//g"'
"@.Trim()

    function _RunSSHCapture([string]$remoteCmd,[string]$stdinText,[int]$TimeoutMs=45000) {

        $args = @(
            "-o","BatchMode=yes",
            "-o","NumberOfPasswordPrompts=0",
            "-o","PreferredAuthentications=publickey",
            "-o","PubkeyAuthentication=yes",
            "-o","StrictHostKeyChecking=accept-new",
            "-o","UserKnownHostsFile=$kh",
            "-o","GlobalKnownHostsFile=NUL",
            "-p",$Port
        )


        if (-not [string]::IsNullOrWhiteSpace($KeyPath) -and -not $script:SshAgentTouched) {
            $args += @("-i", "`"$KeyPath`"")
        }

        $args += "$User@$RemoteHost"
        $rc = $remoteCmd.Replace('"','\"')
        $args += "`"$rc`""

        try { if ($script:AppendLog) { $script:AppendLog.Invoke("[INFO] BackupP12: ssh search starting user=$User host=$RemoteHost port=$Port key=" + $(if($KeyPath){"yes"}else{"no"}) + "`r`n") } } catch {}

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $sshExe
        $psi.Arguments = ($args -join " ")
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.RedirectStandardInput  = $true
        $psi.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $null = $p.Start()

        try {
            if ($stdinText -ne $null -and $stdinText.Length -gt 0) {
                if (-not $stdinText.EndsWith("`n")) { $stdinText += "`n" }
                $p.StandardInput.Write($stdinText)
            }
        } catch {}
        try { $p.StandardInput.Close() } catch {}

        $finished = $false
        try { $finished = $p.WaitForExit($TimeoutMs) } catch { $finished = $false }
        if (-not $finished) {
            try { $p.Kill() } catch {}
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] BackupP12: ssh search timed out after ${TimeoutMs}ms.`r`n") } } catch {}
            return [pscustomobject]@{ Exit=124; StdOut=""; StdErr="timeout" }
        }

        $out = ""; $err = ""
        try { $out = $p.StandardOutput.ReadToEnd() } catch { $out = "" }
        try { $err = $p.StandardError.ReadToEnd() } catch { $err = "" }

        $exit = 1
        try { $exit = [int]$p.ExitCode } catch { $exit = 1 }

        if ($err) {
            $last = ($err -split "`r?`n" | Where-Object { $_ -and $_.Trim() } | Select-Object -Last 1)
            if ($last) {
                try { if ($script:AppendLog) { $script:AppendLog.Invoke("[WARN] BackupP12: ssh stderr(last): $last`r`n") } } catch {}
            }
        }
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("[INFO] BackupP12: ssh exit=$exit stdoutLen=$($out.Length) stderrLen=$($err.Length)`r`n") } } catch {}

        return [pscustomobject]@{ Exit=$exit; StdOut=$out; StdErr=$err }
    }

    $r1 = _RunSSHCapture $plainCmd "" 45000

    if ($r1.Exit -ne 0 -and [string]::IsNullOrWhiteSpace($r1.StdOut)) {
        return @("__SSHFAIL__:" + ($r1.StdErr -replace "`r","" -replace "`n"," | ").Trim())
    }

    $paths1 = @()
    foreach ($ln in ($r1.StdOut -split "`r?`n")) {
        $t = $ln.Trim()
        if ($t) { $paths1 += $t }
    }
    $paths1 = $paths1 | Sort-Object -Unique

    if ($paths1.Count -gt 0) {
        return $paths1
    }

    if ([string]::IsNullOrWhiteSpace($SudoPassword)) {
        return @()
    }

    $paths1 = @()
    foreach ($ln in ($r1.StdOut -split "`r?`n")) {
        $t = $ln.Trim()
        if ($t) { $paths1 += $t }
    }
    $paths1 = $paths1 | Sort-Object -Unique
    if ($paths1.Count -gt 0) { return $paths1 }

    $pw = $SudoPassword
    if ([string]::IsNullOrWhiteSpace($pw)) { return @() }

    $r2 = _RunSSHCapture $sudoCmd $pw 60000

    if ($r2.Exit -ne 0 -and [string]::IsNullOrWhiteSpace($r2.StdOut)) {
        return @("__SSHFAIL__:" + ($r2.StdErr -replace "`r","" -replace "`n"," | ").Trim())
    }

    $paths2 = @()
    foreach ($ln in ($r2.StdOut -split "`r?`n")) {
        $t = $ln.Trim()
        if ($t) { $paths2 += $t }
    }
    return ($paths2 | Sort-Object -Unique)
}

function Get-RemoteSha256 {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][string]$User,
        [Parameter(Mandatory=$true)][string]$RemotePath,
        [string]$KeyPath
    )

    $cmd = "bash -lc 'sha256sum " + ([string]$RemotePath).Replace("'", "'\''") + " 2>/dev/null | awk `{print `$1}`''"
    $sshExe = Get-SshExePath
    if (-not $sshExe) { return "" }

    $kh = Get-KnownHostsPathSafe
    $args = @(
        "-o","BatchMode=yes",
        "-o","NumberOfPasswordPrompts=0",
        "-o","StrictHostKeyChecking=accept-new",
        "-o","UserKnownHostsFile=$kh",
        "-o","GlobalKnownHostsFile=NUL",
        "-p",$Port
    )
    if (-not [string]::IsNullOrWhiteSpace($KeyPath)) { $args += @("-i","`"$KeyPath`"") }
    $args += "$User@$RemoteHost"
    $args += "`"$($cmd.Replace('"','\"'))`""

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $sshExe
    $psi.Arguments = ($args -join " ")
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $null = $p.Start()
    $p.WaitForExit(15000) | Out-Null

    $o = ""
    try { $o = $p.StandardOutput.ReadToEnd() } catch { $o = "" }
    return ($o.Trim())
}

function Download-RemoteP12PreserveMetadata {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][string]$User,
        [Parameter(Mandatory=$true)][string]$RemotePath,
        [Parameter(Mandatory=$true)][string]$LocalPath,
        [string]$KeyPath
    )

    if (Test-Path -LiteralPath $LocalPath) {

        $remoteHash = Get-RemoteSha256 -RemoteHost $RemoteHost -Port $Port -User $User -RemotePath $RemotePath -KeyPath $KeyPath
        $localHash  = ""
        try { $localHash = (Get-FileHash -LiteralPath $LocalPath -Algorithm SHA256).Hash } catch { $localHash = "" }

        if ($remoteHash -and $localHash -and ($remoteHash -ieq $localHash)) {
            return "identical"
        }

        $choice = Show-Confirm3ChoiceDialog `
            -Title "File Exists" `
            -Header "P12 already exists locally" `
            -Body ("Local file exists and is different:`r`n`r`n$LocalPath`r`n`r`nOverwrite it?") `
            -YesText "Overwrite" `
            -NoText "Skip" `
            -CancelText "Cancel"

        if ($choice -ne "yes") { return "skipped" }
    }

    $scpExe = Get-ScpExePath
    if (-not $scpExe) { return "error" }

    $sshExe = Get-SshExePath
    if (-not $sshExe) { return "error" }

    $kh = Get-KnownHostsPathSafe

    function _ScpPull([string]$srcRemotePath) {
        $args = @(
            "-p",
            "-P", $Port,
            "-o", "StrictHostKeyChecking=accept-new",
            "-o", "UserKnownHostsFile=$kh",
            "-o", "GlobalKnownHostsFile=NUL"
        )
        if (-not [string]::IsNullOrWhiteSpace($KeyPath)) { $args += @("-i", "`"$KeyPath`"") }

        $remoteSpec = "$User@${RemoteHost}:`"$srcRemotePath`""
        $args += @(
            $remoteSpec,
            "`"$LocalPath`""
        )

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $scpExe
        $psi.Arguments = ($args -join " ")
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $null = $p.Start()
        $p.WaitForExit(60000) | Out-Null

        return $p.ExitCode
    }

    function _RunSudo([string]$remoteCmd,[string]$sudoPw,[int]$TimeoutMs=45000) {
        $args = @(
            "-o","BatchMode=yes",
            "-o","NumberOfPasswordPrompts=0",
            "-o","StrictHostKeyChecking=accept-new",
            "-o","UserKnownHostsFile=$kh",
            "-o","GlobalKnownHostsFile=NUL",
            "-p",$Port
        )
        if (-not [string]::IsNullOrWhiteSpace($KeyPath)) { $args += @("-i","`"$KeyPath`"") }
        $args += "$User@$RemoteHost"
        $rc = $remoteCmd.Replace('"','\"')
        $args += "`"$rc`""

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $sshExe
        $psi.Arguments = ($args -join " ")
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.RedirectStandardInput  = $true
        $psi.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $null = $p.Start()

        try {
            if (-not [string]::IsNullOrWhiteSpace($sudoPw)) {
                if (-not $sudoPw.EndsWith("`n")) { $sudoPw += "`n" }
                $p.StandardInput.Write($sudoPw)
            }
        } catch {}
        try { $p.StandardInput.Close() } catch {}

        $p.WaitForExit($TimeoutMs) | Out-Null
        return $p.ExitCode
    }

    $exit1 = _ScpPull $RemotePath
    if ($exit1 -eq 0) {
        return "copied"
    }

    $sudoPw = Get-SudoPasswordForRemoteUser -RemoteHost $RemoteHost -Port $Port -RemoteUser $User
    if ([string]::IsNullOrWhiteSpace($sudoPw)) {
        return "error"
    }

    $tmpName = ("server-toolkit-p12-" + [guid]::NewGuid().ToString("N") + ".p12")
    $tmpPath = "/tmp/$tmpName"

    $escapedSrc = ([string]$RemotePath).Replace("'", "'\''")
    $escapedTmp = ([string]$tmpPath).Replace("'", "'\''")

    $stageCmd = "bash -lc 'set -e; " +
                "sudo -S -p """" cp -p -- ''$escapedSrc'' ''$escapedTmp''; " +
                "sudo -S -p """" chmod 0644 -- ''$escapedTmp''; " +
                "echo STAGED_OK'"

    $cleanupCmd = "bash -lc 'set +e; sudo -S -p """" rm -f -- ''$escapedTmp'' >/dev/null 2>&1 || true'"

    $staged = $false
    try {
        $sx = _RunSudo $stageCmd $sudoPw 60000
        if ($sx -ne 0) { return "error" }
        $staged = $true

        $exit2 = _ScpPull $tmpPath
        if ($exit2 -ne 0) { return "error" }

        return "copied"
    }
    finally {
        if ($staged) {
            try { _RunSudo $cleanupCmd $sudoPw 20000 | Out-Null } catch {}
        }
    }
}

function Find-P12FilesFast {
    param(
        [Parameter(Mandatory=$true)][string]$Root,
        [int]$MaxDepth = 5
    )

    $results = New-Object System.Collections.Generic.List[string]

    function Walk([string]$path, [int]$depth) {

        if ($depth -gt $MaxDepth) { return }

        try {
            if ($path -match '(?i)[\\\/]incremental_snapshot($|[\\\/])') { return }
        } catch {}

        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($path, "*.p12")) {
                $results.Add($f)
            }

            foreach ($d in [System.IO.Directory]::EnumerateDirectories($path)) {
                Walk $d ($depth + 1)
            }
        } catch {}
    }

    Walk $Root 0
    return $results
}

function Copy-P12PreserveMetadata {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Dest
    )

    if (Test-Path -LiteralPath $Dest) {

        $srcHash = ""
        $dstHash = ""
        try { $srcHash = (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash } catch {}
        try { $dstHash = (Get-FileHash -LiteralPath $Dest   -Algorithm SHA256).Hash } catch {}

        if ($srcHash -and $dstHash -and ($srcHash -eq $dstHash)) {
            return "identical"
        }

        $choice = Show-Confirm3ChoiceDialog `
            -Title "P12 file exists" `
            -Header "Destination already exists" `
            -Body ("A file already exists at:`r`n`r`n$Dest`r`n`r`nIt is different from the source.") `
            -YesText "Overwrite" `
            -NoText "Skip" `
            -CancelText "Cancel"

        if ($choice -ne "yes") { return "skipped" }
    }

    Copy-Item -LiteralPath $Source -Destination $Dest -Force

    try {
        $srcItem = Get-Item -LiteralPath $Source -Force
        $dstItem = Get-Item -LiteralPath $Dest   -Force

        $dstItem.CreationTime  = $srcItem.CreationTime
        $dstItem.LastWriteTime = $srcItem.LastWriteTime
        $dstItem.LastAccessTime= $srcItem.LastAccessTime

        $dstItem.Attributes = $srcItem.Attributes
    } catch {}

    return "copied"
}

function Show-P12BackupDialog {
    param(
        [Parameter(Mandatory=$true)][string[]]$P12Paths
    )

    $bc = [System.Windows.Media.BrushConverter]::new()

    $w = New-Object System.Windows.Window
    $w.Title = "Backup P12"
    $w.Width = 820
    $w.Height = 560
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = $bc.ConvertFromString("#1e1f22")
    $w.Foreground = $bc.ConvertFromString("#f0f0f0")
    $w.FontFamily = "Segoe UI"

    try {
        if ($MainWindow -and $MainWindow.Resources) {
            if (-not $w.Resources) { $w.Resources = New-Object System.Windows.ResourceDictionary }
            try { $w.Resources.MergedDictionaries.Add($MainWindow.Resources) | Out-Null } catch {}
        }
    } catch {}

    function _TryStyle([string]$key) {
        try { return $MainWindow.FindResource($key) } catch {}
        try { return $w.FindResource($key) } catch {}
        return $null
    }

    $PrimaryBtnStyle   = _TryStyle "StartButtonStyle"
    $SecondaryBtnStyle = _TryStyle "SecondaryButtonStyle"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "Select P12 file(s) to back up"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Margin = "0,0,0,12"
    $root.Children.Add($hdr) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdr,0)

    $panel = New-Object System.Windows.Controls.Border
    $panel.Background = $bc.ConvertFromString("#25272b")
    $panel.BorderBrush = $bc.ConvertFromString("#3b3f46")
    $panel.BorderThickness = "1"
    $panel.CornerRadius = "6"
    $panel.Padding = "12"
    [System.Windows.Controls.Grid]::SetRow($panel,1)

    $inner = New-Object System.Windows.Controls.Grid
    $inner.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $inner.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $lv = New-Object System.Windows.Controls.ListView
    $lv.SelectionMode = "Extended"
    $lv.Background = $bc.ConvertFromString("#2f3238")
    $lv.Foreground = $bc.ConvertFromString("#f0f0f0")

    foreach ($p in $P12Paths) { $null = $lv.Items.Add($p) }

    $inner.Children.Add($lv) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($lv,0)

    $destRow = New-Object System.Windows.Controls.Grid
    $destRow.Margin = "0,12,0,0"
    $destRow.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="*" })) | Out-Null
    $destRow.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto" })) | Out-Null

    $destBox = New-Object System.Windows.Controls.TextBox
    $destBox.IsReadOnly = $true
    $destBox.Height = 26
    $destBox.Padding = "6,3"
    $destBox.Text = (Join-Path $env:USERPROFILE "P12-Backups")
    $destRow.Children.Add($destBox) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($destBox,0)

    $btnBrowse = New-Object System.Windows.Controls.Button
    $btnBrowse.Content = "Browse..."
    $btnBrowse.Margin = "10,0,0,0"
    if ($SecondaryBtnStyle) { $btnBrowse.Style = $SecondaryBtnStyle }
    $btnBrowse.Add_Click({
        try {
            $f = New-Object System.Windows.Forms.FolderBrowserDialog
            $f.SelectedPath = $destBox.Text
            if ($f.ShowDialog() -eq "OK") { $destBox.Text = $f.SelectedPath }
        } catch {}
    })
    $destRow.Children.Add($btnBrowse) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($btnBrowse,1)

    $inner.Children.Add($destRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($destRow,1)

    $panel.Child = $inner
    $root.Children.Add($panel) | Out-Null

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"
    [System.Windows.Controls.Grid]::SetRow($btnRow,2)

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Close"
    if ($SecondaryBtnStyle) { $btnCancel.Style = $SecondaryBtnStyle }
    $btnCancel.Margin = "0,0,10,0"
    $btnCancel.Add_Click({ $w.Close() })
    $btnRow.Children.Add($btnCancel) | Out-Null

    $btnBackup = New-Object System.Windows.Controls.Button
    $btnBackup.Content = "Backup Selected"
    if ($PrimaryBtnStyle) { $btnBackup.Style = $PrimaryBtnStyle }
    $btnBackup.Add_Click({
        try {
            $selected = @()
            foreach ($i in $lv.SelectedItems) { $selected += [string]$i }

            if (-not $selected -or $selected.Count -eq 0) {
                [System.Windows.MessageBox]::Show($w, "Select one or more P12 files first.", "No Selection",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            $destRoot = $destBox.Text
            if ([string]::IsNullOrWhiteSpace($destRoot)) { return }

            try { if (-not (Test-Path -LiteralPath $destRoot)) { New-Item -ItemType Directory -Path $destRoot -Force | Out-Null } } catch {}

            $copied=0; $identical=0; $skipped=0

            $remoteHost = ""; $port=""; $user=""; $keyPath=""
            try { $remoteHost = ($HostBox.Text).Trim() } catch {}
            try { $port       = ($PortBox.Text).Trim() } catch {}
            try { $user       = ($UserBox.Text).Trim() } catch {}
            try { $keyPath    = ($KeyBox.Text).Trim() } catch {}

            foreach ($remotePath in $selected) {
                try {
                    $leaf = Split-Path -Leaf $remotePath
                    $dst  = Join-Path $destRoot $leaf

                    $r = Download-RemoteP12PreserveMetadata `
                        -RemoteHost $remoteHost `
                        -Port $port `
                        -User $user `
                        -RemotePath $remotePath `
                        -LocalPath $dst `
                        -KeyPath $keyPath

                    switch ($r) {
                        "copied"    { $copied++ }
                        "identical" { $identical++ }
                        default     { $skipped++ }
                    }
                } catch { $skipped++ }
            }

            [System.Windows.MessageBox]::Show(
                $w,
                ("Backup complete.`r`n`r`nCopied: $copied`r`nAlready present (identical): $identical`r`nSkipped: $skipped"),
                "Backup P12",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null

        } catch {}
    })
    $btnRow.Children.Add($btnBackup) | Out-Null

    $root.Children.Add($btnRow) | Out-Null

    $w.Content = $root
    $null = $w.ShowDialog()
}

function Invoke-P12BackupFlow {
    try {
        if (-not $BackupP12Button -or -not $BackupP12Button.IsEnabled) { return }

        $remoteHost = ""; $port=""; $user=""; $keyPath=""
        try { $remoteHost = ($HostBox.Text).Trim() } catch {}
        try { $port       = ($PortBox.Text).Trim() } catch {}
        try { $user       = ($UserBox.Text).Trim() } catch {}
        try { $keyPath    = ($KeyBox.Text).Trim() } catch {}

        if ([string]::IsNullOrWhiteSpace($remoteHost) -or [string]::IsNullOrWhiteSpace($port) -or [string]::IsNullOrWhiteSpace($user)) {
            [System.Windows.MessageBox]::Show($MainWindow, "Host / Port / Username are required.", "Backup P12",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        $bc = [System.Windows.Media.BrushConverter]::new()
        $pw = New-Object System.Windows.Window
        $pw.Title = "Backup P12"
        $pw.Width = 560
        $pw.Height = 170
        $pw.WindowStartupLocation = "CenterOwner"
        $pw.ResizeMode = "NoResize"
        $pw.Owner = $MainWindow
        $pw.Background = $bc.ConvertFromString("#1e1f22")
        $pw.Foreground = $bc.ConvertFromString("#f0f0f0")
        $pw.FontFamily = "Segoe UI"
        $pw.WindowStyle = "ToolWindow"

        $g = New-Object System.Windows.Controls.Grid
        $g.Margin = "18"
        $g.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
        $g.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = "Please Wait, Searching the server for P12 files..."
        $t.FontSize = 16
        $t.FontWeight = "SemiBold"
        $t.Margin = "0,0,0,12"
        $g.Children.Add($t) | Out-Null
        [System.Windows.Controls.Grid]::SetRow($t,0)

        $bar = New-Object System.Windows.Controls.ProgressBar
        $bar.IsIndeterminate = $true
        $bar.Height = 18
        $g.Children.Add($bar) | Out-Null
        [System.Windows.Controls.Grid]::SetRow($bar,1)

        $pw.Content = $g

        $MainWindow.Cursor = "Wait"

        $found = $null

        function _BkpDbg([string]$m) {
            try {
                Add-Content -Path $global:ServerToolkitLogPath -Value (
                    (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                    " [BKP12] " + $m + "`r`n"
                ) -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
        }

        _BkpDbg ("Invoke-P12BackupFlow entered host=$remoteHost port=$port user=$user key=" + $(if($keyPath){"yes"}else{"no"}))

        $ui_passphrase = ""
        try { $ui_passphrase = $PasswordBox.Password } catch { $ui_passphrase = "" }
        try { _BkpDbg ("Captured ui_passphrase len=" + $ui_passphrase.Length) } catch {}

        $ui_sudoPw = ""

        try { _BkpDbg ("Captured sudoPw len=" + $ui_sudoPw.Length) } catch {}
        try { _BkpDbg "Starting Task.Run for Find-RemoteP12FilesFast..." } catch {}

        $pw.Show()

        try {
            Add-Content -Path $global:ServerToolkitLogPath -Value (
                (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                " [BKP12] Running search synchronously (STA-safe)`r`n"
            ) -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}

        $agentOk = Ensure-SshAgentSessionPreStart -KeyPath $keyPath -Passphrase $ui_passphrase
        if ($agentOk) {
            $agentOk = Ensure-SshAgentWithKey -KeyPath $keyPath -Passphrase $ui_passphrase
        }

        if (-not $agentOk) {
            try { $pw.Close() } catch {}
            [System.Windows.MessageBox]::Show(
                $MainWindow,
                "SSH key is not unlocked or usable. Please unlock the key and retry Backup P12.",
                "Backup P12",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
            return
        }

        $found = Find-RemoteP12FilesFast `
            -RemoteHost $remoteHost `
            -Port $port `
            -User $user `
            -KeyPath $keyPath `
            -SudoPassword ""

        if (-not $found -or $found.Count -eq 0) {
            $sudoPw = Get-SudoPasswordForRemoteUser `
                -RemoteHost $remoteHost `
                -Port $port `
                -RemoteUser $user

            if (-not [string]::IsNullOrWhiteSpace($sudoPw)) {
                $found = Find-RemoteP12FilesFast `
                    -RemoteHost $remoteHost `
                    -Port $port `
                    -User $user `
                    -KeyPath $keyPath `
                    -SudoPassword $sudoPw
            }
        }

        try { $pw.Close() } catch {}
        $MainWindow.Cursor = "Arrow"

        try {
            if ($found -and $found.Count -ge 1 -and ([string]$found[0]) -eq "__NEED_SSHADD__") {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "Your SSH key is not usable yet (ssh-agent / ssh-add is not ready).`r`n`r`n" +
                    "Enter the SSH key passphrase (Password / SSH Passphrase field), then retry Backup P12. " +
                    "If a key-unlock prompt appears, complete it first.",
                    "Backup P12 - SSH Key Not Ready",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                ) | Out-Null
                return
            }
        } catch {}

        try { $pw.Close() } catch {}
        $MainWindow.Cursor = "Arrow"

        try {
            if ($found -and $found.Count -ge 1 -and ([string]$found[0]).StartsWith("__SSHFAIL__:")) {
                $detail = ([string]$found[0]).Substring(11)
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "Backup P12 could not authenticate to the server (ssh failed).`r`n`r`nDetails:`r`n$detail",
                    "Backup P12 - SSH Failed",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
                return
            }
        } catch {}

        if (-not $found -or $found.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                $MainWindow,
                "No .p12 files were found anywhere under / within depth 5 (excluding folders named incremental_snapshot).",
                "Backup P12",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
            return
        }

        Show-P12BackupDialog -P12Paths $found
    }
    catch {
        try { $MainWindow.Cursor = "Arrow" } catch {}
        try {
            [System.Windows.MessageBox]::Show(
                $MainWindow,
                ("Backup P12 crashed:`r`n`r`n" + $_.Exception.Message),
                "Backup P12",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            ) | Out-Null
        } catch {}
    }
}
function Flash-TextBoxWarning {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Controls.TextBox]$Tb,
        [int]$Ms = 450
    )

    try {
        if (-not $Tb) { return }

        $warnBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#f5c542")

        try { $Tb.BorderBrush = $warnBrush } catch {}
        try { $Tb.BorderThickness = [System.Windows.Thickness]::new(2) } catch {}

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds($Ms)
        $timer.Tag = $Tb

        $timer.Add_Tick({
            param($s, $e)
            try {
                $t = $s.Tag
                if ($t) {
                    try { $t.ClearValue([System.Windows.Controls.TextBox]::BorderBrushProperty) } catch {}
                    try { $t.ClearValue([System.Windows.Controls.TextBox]::BorderThicknessProperty) } catch {}
                }
            } catch {}
            try { $s.Stop() } catch {}
        })

        $timer.Start()
    } catch {}
}

function Sanitize-UbuntuUsername {
    param([string]$Text)

    try {
        if ($null -eq $Text) { return "" }

        $t = ([string]$Text).ToLowerInvariant()
        $t = [regex]::Replace($t, '[^a-z0-9_-]', '')
        $t = [regex]::Replace($t, '^[^a-z]+', '')
        if ($t.Length -gt 32) { $t = $t.Substring(0, 32) }

        return $t
    } catch {
        return ""
    }
}

function Test-UbuntuUsernameCandidate {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Controls.TextBox]$Tb,
        [Parameter(Mandatory=$true)][string]$InsertText
    )

    try {
        $before = ""
        try { $before = [string]$Tb.Text } catch { $before = "" }

        $selStart = 0
        $selLen   = 0
        try { $selStart = [int]$Tb.SelectionStart } catch { $selStart = 0 }
        try { $selLen   = [int]$Tb.SelectionLength } catch { $selLen = 0 }

        if ($selStart -lt 0) { $selStart = 0 }
        if ($selStart -gt $before.Length) { $selStart = $before.Length }
        if ($selLen -lt 0) { $selLen = 0 }
        if (($selStart + $selLen) -gt $before.Length) { $selLen = ($before.Length - $selStart) }

        $candidate =
            $before.Substring(0, $selStart) +
            $InsertText +
            $before.Substring($selStart + $selLen)

        $candidate = Sanitize-UbuntuUsername -Text $candidate

        if ([string]::IsNullOrEmpty($candidate)) { return $true }
        return ($candidate -match '^[a-z][a-z0-9_-]{0,31}$')
    } catch {
        return $false
    }
}

function Attach-UbuntuUsernameFilter {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Controls.TextBox]$Tb,
        [string]$Label = "Username"
    )

    if (-not $Tb) { return }

    $Tb.Add_PreviewTextInput({
        param($s, $e)

        try {
            if (-not $e -or [string]::IsNullOrEmpty($e.Text)) { return }

            $ins = $e.Text

            if ($ins -match '\s') {
                $e.Handled = $true
                try { Flash-TextBoxWarning -Tb ([System.Windows.Controls.TextBox]$s) } catch {}
                return
            }

            $lower = $ins.ToLowerInvariant()
            $lowerSan = [regex]::Replace($lower, '[^a-z0-9_-]', '')

            if ([string]::IsNullOrEmpty($lowerSan)) {
                $e.Handled = $true
                try { Flash-TextBoxWarning -Tb ([System.Windows.Controls.TextBox]$s) } catch {}
                return
            }

            $tb = [System.Windows.Controls.TextBox]$s

            if (-not (Test-UbuntuUsernameCandidate -Tb $tb -InsertText $lowerSan)) {
                $e.Handled = $true
                try { Flash-TextBoxWarning -Tb $tb } catch {}
                return
            }

            if ($lowerSan -ne $ins) {
                $e.Handled = $true

                $UiDispatcher.BeginInvoke([Action]{
                    try {
                        $tb2 = $tb
                        if (-not $tb2) { return }
                        $tb2.SelectedText = $lowerSan
                        $tb2.CaretIndex = $tb2.SelectionStart + $lowerSan.Length
                    } catch {}
                }) | Out-Null
                return
            }
        } catch {
            try {
                $e.Handled = $true
                try { Flash-TextBoxWarning -Tb ([System.Windows.Controls.TextBox]$s) } catch {}
            } catch {}
        }
    })

    $Tb.Add_PreviewKeyDown({
        param($s, $e)
        try {
            if (-not $e) { return }
            if ($e.Key -eq [System.Windows.Input.Key]::Space) {
                $e.Handled = $true
                try { Flash-TextBoxWarning -Tb ([System.Windows.Controls.TextBox]$s) } catch {}
                return
            }
        } catch {}
    })

    try {
        [System.Windows.DataObject]::AddPastingHandler($Tb, [System.Windows.DataObjectPastingEventHandler]{
            param($s, $e)
            try {
                if (-not $e) { return }

                $raw = $null
                try { $raw = $e.DataObject.GetData([System.Windows.DataFormats]::UnicodeText) } catch { $raw = $null }
                if ($null -eq $raw) {
                    try { $raw = $e.DataObject.GetData([System.Windows.DataFormats]::Text) } catch { $raw = $null }
                }

                $e.CancelCommand()

                $tb = [System.Windows.Controls.TextBox]$s
                if ($null -eq $raw) {
                    try { Flash-TextBoxWarning -Tb $tb } catch {}
                    return
                }

                $san = Sanitize-UbuntuUsername -Text ([string]$raw)

                if ([string]::IsNullOrEmpty($san)) {
                    try { Flash-TextBoxWarning -Tb $tb } catch {}
                    return
                }

                if (-not (Test-UbuntuUsernameCandidate -Tb $tb -InsertText $san)) {
                    try { Flash-TextBoxWarning -Tb $tb } catch {}
                    return
                }

                $UiDispatcher.BeginInvoke([Action]{
                    try {
                        $tb2 = $tb
                        if (-not $tb2) { return }
                        $tb2.SelectedText = $san
                        $tb2.CaretIndex = $tb2.SelectionStart + $san.Length
                    } catch {}
                }) | Out-Null
            } catch {
                try {
                    try { $e.CancelCommand() } catch {}
                    try { Flash-TextBoxWarning -Tb ([System.Windows.Controls.TextBox]$s) } catch {}
                } catch {}
            }
        })
    } catch {}

    $Tb.Add_TextChanged({
        param($s, $e)
        try {
            $tb = [System.Windows.Controls.TextBox]$s
            if (-not $tb) { return }

            if (-not ($tb.Tag -is [hashtable])) { $tb.Tag = @{} }
            if ($tb.Tag["__UbUserSanitizeGuard"] -eq $true) { return }
            $tb.Tag["__UbUserSanitizeGuard"] = $true

            $orig = [string]$tb.Text
            $san  = Sanitize-UbuntuUsername -Text $orig

            if ($san -ne $orig) {
                $caret = 0
                try { $caret = [int]$tb.CaretIndex } catch { $caret = 0 }

                $tb.Text = $san

                if ($caret -gt $san.Length) { $caret = $san.Length }
                if ($caret -lt 0) { $caret = 0 }
                $tb.CaretIndex = $caret
            }
        } catch {
        } finally {
            try {
                $tb2 = [System.Windows.Controls.TextBox]$s
                if ($tb2 -and ($tb2.Tag -is [hashtable])) {
                    $tb2.Tag["__UbUserSanitizeGuard"] = $false
                }
            } catch {}
        }
    })

    try {
        $Tb.Text = (Sanitize-UbuntuUsername -Text ([string]$Tb.Text))
    } catch {}
}

try {
    if ($UserBox)    { Attach-UbuntuUsernameFilter -Tb $UserBox    -Label "Server Username" }
    if ($NewUserBox) { Attach-UbuntuUsernameFilter -Tb $NewUserBox -Label "New Username" }
} catch {}

try {
    if ($UserBox) {
        $UserBox.Add_TextChanged({
            param($s, $e)
            try {
                $tb = [System.Windows.Controls.TextBox]$s
                if (-not $tb) { return }

                if (-not ($tb.Tag -is [hashtable])) { $tb.Tag = @{} }
                if ($tb.Tag["__ForceLowerGuard"] -eq $true) { return }
                $tb.Tag["__ForceLowerGuard"] = $true

                $t  = [string]$tb.Text
                $lo = $t.ToLowerInvariant()

                if ($lo -ne $t) {
                    $c = 0
                    try { $c = [int]$tb.CaretIndex } catch { $c = 0 }

                    $tb.Text = $lo

                    if ($c -gt $lo.Length) { $c = $lo.Length }
                    if ($c -lt 0) { $c = 0 }
                    $tb.CaretIndex = $c
                }
            } catch {} finally {
                try {
                    $tb2 = [System.Windows.Controls.TextBox]$s
                    if ($tb2 -and ($tb2.Tag -is [hashtable])) { $tb2.Tag["__ForceLowerGuard"] = $false }
                } catch {}
            }
        })
    }

    if ($NewUserBox) {
        $NewUserBox.Add_TextChanged({
            param($s, $e)
            try {
                $tb = [System.Windows.Controls.TextBox]$s
                if (-not $tb) { return }

                if (-not ($tb.Tag -is [hashtable])) { $tb.Tag = @{} }
                if ($tb.Tag["__ForceLowerGuard"] -eq $true) { return }
                $tb.Tag["__ForceLowerGuard"] = $true

                $t  = [string]$tb.Text
                $lo = $t.ToLowerInvariant()

                if ($lo -ne $t) {
                    $c = 0
                    try { $c = [int]$tb.CaretIndex } catch { $c = 0 }

                    $tb.Text = $lo

                    if ($c -gt $lo.Length) { $c = $lo.Length }
                    if ($c -lt 0) { $c = 0 }
                    $tb.CaretIndex = $c
                }
            } catch {} finally {
                try {
                    $tb2 = [System.Windows.Controls.TextBox]$s
                    if ($tb2 -and ($tb2.Tag -is [hashtable])) { $tb2.Tag["__ForceLowerGuard"] = $false }
                } catch {}
            }
        })
    }
} catch {}


try {
    if ($HostBox) {
        $HostBox.Add_TextChanged({
            try { Update-OpenServerButtonVisibility } catch {}
            try { Safe-UpdateBackupP12ButtonState } catch {}
        })
    }
    if ($PortBox) {
        $PortBox.Add_TextChanged({
            try { Update-OpenServerButtonVisibility } catch {}
            try { Safe-UpdateBackupP12ButtonState } catch {}
        })
    }
    if ($UserBox) {
        $UserBox.Add_TextChanged({
            try { Update-OpenServerButtonVisibility } catch {}
            try { Safe-UpdateBackupP12ButtonState } catch {}
        })
    }
    if ($KeyBox) {
        $KeyBox.Add_TextChanged({
            try { Update-OpenServerButtonVisibility } catch {}
        })
    }
} catch {}

try { Update-OpenServerButtonVisibility } catch {}
try { Update-StartButtonState } catch {}
try { Safe-UpdateBackupP12ButtonState } catch {}

$script:PassToggleHover = $false

function Update-PassToggleVisual {
    try {
        if (-not $TogglePassReveal) { return }

        $bgDefault = "#2f3238"
        $bgHover   = "#25272b"
        $bgChecked = "#1e1f22"
        $fgIcon    = "#f0f0f0"

        $hex = $bgDefault
        try {
            if ($TogglePassReveal.IsChecked -eq $true) {
                $hex = $bgChecked
            } elseif ($script:PassToggleHover -eq $true) {
                $hex = $bgHover
            }
        } catch {}

        $b = [System.Windows.Media.BrushConverter]::new().ConvertFromString($hex)
        $TogglePassReveal.Background  = $b
        $TogglePassReveal.BorderBrush = $b

        $TogglePassReveal.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgIcon)

        if ($PassRevealIcon) {
            $PassRevealIcon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgIcon)
        }
    } catch {}
}

try {
    if ($TogglePassReveal) {

        $TogglePassReveal.Add_MouseEnter({
            try { $script:PassToggleHover = $true } catch {}
            try { Update-PassToggleVisual } catch {}
        })

        $TogglePassReveal.Add_MouseLeave({
            try { $script:PassToggleHover = $false } catch {}
            try { Update-PassToggleVisual } catch {}
        })

        $TogglePassReveal.Add_Checked({ try { Update-PassToggleVisual } catch {} })
        $TogglePassReveal.Add_Unchecked({ try { Update-PassToggleVisual } catch {} })

        Update-PassToggleVisual
    }
} catch {}

$script:PassRevealSyncGuard = $false
function Set-PassRevealState {
    param([bool]$Reveal)

    try {
        if (-not $PasswordBox -or -not $PasswordRevealBox) { return }

        if ($Reveal) {
            try { $PasswordRevealBox.Text = $PasswordBox.Password } catch {}
            $PasswordBox.Visibility = "Collapsed"
            $PasswordRevealBox.Visibility = "Visible"
            try { $PasswordRevealBox.Focus() } catch {}

            try { if ($PassRevealIcon) { $PassRevealIcon.Text = [char]0xE72E } } catch {}
            try { Update-PassToggleVisual } catch {}
        } else {
            try { $PasswordBox.Password = $PasswordRevealBox.Text } catch {}
            $PasswordRevealBox.Visibility = "Collapsed"
            $PasswordBox.Visibility = "Visible"
            try { $PasswordBox.Focus() } catch {}

            try { if ($PassRevealIcon) { $PassRevealIcon.Text = [char]0xE890 } } catch {}
            try { Update-PassToggleVisual } catch {}
        }
    } catch {}
}

try {
    if ($PasswordBox -and $PasswordRevealBox) {

        $PasswordBox.Add_PasswordChanged({
            if ($script:PassRevealSyncGuard) { return }
            try {
                $script:PassRevealSyncGuard = $true
                if ($PasswordRevealBox.Visibility -eq "Visible") {
                    $PasswordRevealBox.Text = $PasswordBox.Password
                }
            } catch {} finally { $script:PassRevealSyncGuard = $false }
        })

        $PasswordRevealBox.Add_TextChanged({
            if ($script:PassRevealSyncGuard) { return }
            try {
                $script:PassRevealSyncGuard = $true
                $PasswordBox.Password = $PasswordRevealBox.Text
            } catch {} finally { $script:PassRevealSyncGuard = $false }
        })
    }
} catch {}

try {
    if ($TogglePassReveal) {
        $TogglePassReveal.Add_Checked({ Set-PassRevealState -Reveal $true })
        $TogglePassReveal.Add_Unchecked({ Set-PassRevealState -Reveal $false })
    }
} catch {}

try { Set-PassRevealState -Reveal $false } catch {}

$script:NewPassRevealSyncGuard = $false

function Set-NewPassRevealState {
    param([bool]$Reveal)
    try {
        if (-not $NewPasswordBox -or -not $NewPasswordRevealBox) { return }

        if ($Reveal) {
            try { $NewPasswordRevealBox.Text = $NewPasswordBox.Password } catch {}
            $NewPasswordBox.Visibility = "Collapsed"
            $NewPasswordRevealBox.Visibility = "Visible"
            try { $NewPasswordRevealBox.Focus() } catch {}
            try { if ($NewPassRevealIcon) { $NewPassRevealIcon.Text = [char]0xE72E } } catch {}
        } else {
            try { $NewPasswordBox.Password = $NewPasswordRevealBox.Text } catch {}
            $NewPasswordRevealBox.Visibility = "Collapsed"
            $NewPasswordBox.Visibility = "Visible"
            try { $NewPasswordBox.Focus() } catch {}
            try { if ($NewPassRevealIcon) { $NewPassRevealIcon.Text = [char]0xE890 } } catch {}
        }
    } catch {}
}

try {
    if ($NewPasswordBox -and $NewPasswordRevealBox) {

        $NewPasswordBox.Add_PasswordChanged({
            if ($script:NewPassRevealSyncGuard) { return }
            try {
                $script:NewPassRevealSyncGuard = $true
                if ($NewPasswordRevealBox.Visibility -eq "Visible") {
                    $NewPasswordRevealBox.Text = $NewPasswordBox.Password
                }
            } catch {} finally { $script:NewPassRevealSyncGuard = $false }
        })

        $NewPasswordRevealBox.Add_TextChanged({
            if ($script:NewPassRevealSyncGuard) { return }
            try {
                $script:NewPassRevealSyncGuard = $true
                $NewPasswordBox.Password = $NewPasswordRevealBox.Text
            } catch {} finally { $script:NewPassRevealSyncGuard = $false }
        })
    }
} catch {}

try {
    if ($ToggleNewPassReveal) {
        $ToggleNewPassReveal.Add_Checked({ Set-NewPassRevealState -Reveal $true })
        $ToggleNewPassReveal.Add_Unchecked({ Set-NewPassRevealState -Reveal $false })
    }
} catch {}

try { Set-NewPassRevealState -Reveal $false } catch {}

$script:P12PassRevealSyncGuard = $false

function Set-P12PassRevealState {
    param([bool]$Reveal)
    try {
        if (-not $P12PasswordBox -or -not $P12PasswordRevealBox) { return }

        if ($Reveal) {
            try { $P12PasswordRevealBox.Text = $P12PasswordBox.Password } catch {}
            $P12PasswordBox.Visibility = "Collapsed"
            $P12PasswordRevealBox.Visibility = "Visible"
            try { $P12PasswordRevealBox.Focus() } catch {}
            try { if ($P12PassRevealIcon) { $P12PassRevealIcon.Text = [char]0xE72E } } catch {}
        } else {
            try { $P12PasswordBox.Password = $P12PasswordRevealBox.Text } catch {}
            $P12PasswordRevealBox.Visibility = "Collapsed"
            $P12PasswordBox.Visibility = "Visible"
            try { $P12PasswordBox.Focus() } catch {}
            try { if ($P12PassRevealIcon) { $P12PassRevealIcon.Text = [char]0xE890 } } catch {}
        }
    } catch {}
}

try {
    if ($P12PasswordBox -and $P12PasswordRevealBox) {

        $P12PasswordBox.Add_PasswordChanged({
            if ($script:P12PassRevealSyncGuard) { return }
            try {
                $script:P12PassRevealSyncGuard = $true
                if ($P12PasswordRevealBox.Visibility -eq "Visible") {
                    $P12PasswordRevealBox.Text = $P12PasswordBox.Password
                }
            } catch {} finally { $script:P12PassRevealSyncGuard = $false }
        })

        $P12PasswordRevealBox.Add_TextChanged({
            if ($script:P12PassRevealSyncGuard) { return }
            try {
                $script:P12PassRevealSyncGuard = $true
                $P12PasswordBox.Password = $P12PasswordRevealBox.Text
            } catch {} finally { $script:P12PassRevealSyncGuard = $false }
        })
    }
} catch {}

try {
    if ($ToggleP12PassReveal) {
        $ToggleP12PassReveal.Add_Checked({ Set-P12PassRevealState -Reveal $true })
        $ToggleP12PassReveal.Add_Unchecked({ Set-P12PassRevealState -Reveal $false })
    }
} catch {}

try { Set-P12PassRevealState -Reveal $false } catch {}

$script:PassToggleHover = $false

function Update-PassToggleVisual {
    try {
        if (-not $TogglePassReveal) { return }

        $bg = "#2f3238"
        try {
            if ($TogglePassReveal.IsChecked -eq $true) {
                $bg = "#1e1f22"
            } elseif ($script:PassToggleHover -eq $true) {
                $bg = "#3b3f46"
            }
        } catch {}

        $TogglePassReveal.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bg)
        $TogglePassReveal.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bg)

        $TogglePassReveal.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#f0f0f0")

        if ($PassRevealIcon) {
            $PassRevealIcon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#f0f0f0")
        }
    } catch {}
}

try {
    if ($TogglePassReveal) {

        $TogglePassReveal.Add_MouseEnter({
            try { $script:PassToggleHover = $true } catch {}
            try { Update-PassToggleVisual } catch {}
        })

        $TogglePassReveal.Add_MouseLeave({
            try { $script:PassToggleHover = $false } catch {}
            try { Update-PassToggleVisual } catch {}
        })

        $TogglePassReveal.Add_Checked({ try { Update-PassToggleVisual } catch {} })
        $TogglePassReveal.Add_Unchecked({ try { Update-PassToggleVisual } catch {} })

        Update-PassToggleVisual
    }
} catch {}


if ($OpenServerButton) {
    $OpenServerButton.Add_Click({

        $h  = ""
        $p  = ""
        $u  = ""
        $k  = ""

        try { $h = ($HostBox.Text).Trim() } catch { $h = "" }
        try { $p = ($PortBox.Text).Trim() } catch { $p = "" }
        try { $u = ($UserBox.Text).Trim() } catch { $u = "" }
        try { $k = ($KeyBox.Text).Trim()  } catch { $k = "" }

        if ([string]::IsNullOrWhiteSpace($h) -or
            [string]::IsNullOrWhiteSpace($u) -or
            ($p -notmatch '^\d+$')) {

            try {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "Host, Port, and Username are required.",
                    "Open Server",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                ) | Out-Null
            } catch {}
            return
        }

        $winSsh = Join-Path $env:WINDIR "System32\OpenSSH\ssh.exe"
        if (-not (Test-Path $winSsh)) {
            try {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "OpenSSH client (ssh.exe) not found.",
                    "Open Server",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
            } catch {}
            return
        }

        $knownHostsPath = $null
        try {
            $sshDir = Join-Path $env:USERPROFILE ".ssh"
            if (-not (Test-Path $sshDir)) { New-Item -ItemType Directory -Path $sshDir | Out-Null }
            $knownHostsPath = Join-Path $sshDir "known_hosts"
            if (-not (Test-Path $knownHostsPath)) { New-Item -ItemType File -Path $knownHostsPath -Force | Out-Null }
        } catch {
            $knownHostsPath = (Join-Path $env:USERPROFILE ".ssh\known_hosts")
        }

        try {
            $probeArgs = @(
                "-p", $p,
                "-o", "BatchMode=yes",
                "-o", "NumberOfPasswordPrompts=0",
                "-o", "ConnectTimeout=4",
                "-o", "StrictHostKeyChecking=yes",
                "-o", "UserKnownHostsFile=$knownHostsPath",
                "-o", "GlobalKnownHostsFile=NUL"
            )

            if (-not [string]::IsNullOrWhiteSpace($k)) { $probeArgs += @("-i", "`"$k`"") }

            $probeArgs += "$u@$h"
            $probeArgs += "exit"

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $winSsh
            $psi.Arguments = ($probeArgs -join " ")
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.CreateNoWindow = $true

            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $psi
            $null = $proc.Start()

            if ($proc.WaitForExit(6000)) {
                $stderr = ""
                try { $stderr = $proc.StandardError.ReadToEnd() } catch { $stderr = "" }

                if ($stderr -match '(?i)REMOTE HOST IDENTIFICATION HAS CHANGED' -or
                    $stderr -match '(?i)Host key verification failed' -or
                    $stderr -match '(?i)Offending .* key in') {

                    try { if ($script:AppendLog) { $script:AppendLog.Invoke("[WARN] Host key mismatch detected. Resetting known_hosts entries for ${h}:${p}...`r`n") } } catch {}
                    try { Reset-HostKeysForServer -HostName $h -Port $p } catch {}
                }
            } else {
                try { $proc.Kill() } catch {}
            }
        } catch {

        }

        $args = @(
            "-p", $p,
            "-o", "StrictHostKeyChecking=accept-new",
            "-o", "UserKnownHostsFile=$knownHostsPath",
            "-o", "GlobalKnownHostsFile=NUL"
        )

        if (-not [string]::IsNullOrWhiteSpace($k)) { $args += @("-i", "`"$k`"") }
        $args += "$u@$h"

        $title = "ServerToolkit SSH ${u}@${h}:${p}"
        $cmd   = "title $title & `"$winSsh`" " + ($args -join " ")

        try {
            Start-Process -FilePath "cmd.exe" -ArgumentList "/k", $cmd | Out-Null
        } catch {
            try { $script:AppendLog.Invoke("[ERROR] Failed to launch ssh window: $($_.Exception.Message)`r`n") } catch {}
        }
    })
}

try {
    if ($NewUserPanel) { $NewUserPanel.IsEnabled = ($CreateNonRootCheck -and ($CreateNonRootCheck.IsChecked -eq $true)) }
    if ($P12Panel)     { $P12Panel.IsEnabled     = ($UploadP12Check   -and ($UploadP12Check.IsChecked -eq $true)) }
} catch {}

if ($CreateNonRootCheck -and $NewUserPanel) {

    $CreateNonRootCheck.Add_Checked({
        try { $NewUserPanel.IsEnabled = $true } catch {}
        try { Update-StartButtonState } catch {}
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("DEBUG: CreateNonRootCheck Checked. NewUserPanel enabled.`r`n") } } catch {}
    })

    $CreateNonRootCheck.Add_Unchecked({
        try { $NewUserPanel.IsEnabled = $false } catch {}
        try { Update-StartButtonState } catch {}
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("DEBUG: CreateNonRootCheck Unchecked. NewUserPanel disabled.`r`n") } } catch {}
    })
}

if ($UploadP12Check -and $P12Panel) {

    $UploadP12Check.Add_Checked({
        try { $P12Panel.IsEnabled = $true } catch {}
        try { Update-StartButtonState } catch {}
        try { Safe-UpdateBackupP12ButtonState } catch {}
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("DEBUG: UploadP12Check Checked. P12Panel enabled.`r`n") } } catch {}
    })

    $UploadP12Check.Add_Unchecked({
        try { $P12Panel.IsEnabled = $false } catch {}
        try { Update-StartButtonState } catch {}
        try { Safe-UpdateBackupP12ButtonState } catch {}

        try { if ($P12PathBox) { $P12PathBox.Text = "" } } catch {}
        try { if ($P12PasswordBox) { $P12PasswordBox.Password = "" } } catch {}

        try { $script:LastValidatedP12Alias = "" } catch {}
        try {
            if ($script:TaskResult) {
                $script:TaskResult.P12Alias      = ""
                $script:TaskResult.P12RemotePath = ""
                $script:TaskResult.P12LocalPath  = ""
                $script:TaskResult.ShowP12Alias  = 0
            }
        } catch {}

        try { if ($script:AppendLog) { $script:AppendLog.Invoke("DEBUG: UploadP12Check Unchecked. P12Panel disabled + cleared P12 fields + cleared cached alias.`r`n") } } catch {}
    })
}

try {
    if ($P12PathBox) {

        try { $script:LastP12Path = ($P12PathBox.Text).Trim() } catch { $script:LastP12Path = "" }

        $P12PathBox.Add_TextChanged({
            try {
                $cur = ""
                try { $cur = ($P12PathBox.Text).Trim() } catch { $cur = "" }

                if ($cur -ne $script:LastP12Path) {
                    try { if ($P12PasswordBox) { $P12PasswordBox.Password = "" } } catch {}

                    try { $script:LastValidatedP12Alias = "" } catch {}
                    try {
                        if ($script:TaskResult) {
                            $script:TaskResult.P12Alias     = ""
                            $script:TaskResult.ShowP12Alias = 0
                        }
                    } catch {}

                    try { if ($script:AppendLog) { $script:AppendLog.Invoke("DEBUG: P12 path changed (strict). Cleared P12 passphrase field + cleared cached alias.`r`n") } } catch {}
                    $script:LastP12Path = $cur
                }
            } catch {}
        })
    }
} catch {}

$script:SuppressFileWrite = $false

$script:LogTailTimer      = $null
$script:LogTailLastPos    = 0

function Start-LogTail {
    try {
        if (-not $global:ServerToolkitLogPath) { return }
        if (-not (Test-Path $global:ServerToolkitLogPath)) { return }

        try {
            $script:LogTailLastPos = (Get-Item -LiteralPath $global:ServerToolkitLogPath).Length
        } catch {
            $script:LogTailLastPos = 0
        }

        if ($script:LogTailTimer) {
            try { $script:LogTailTimer.Stop() } catch {}
            $script:LogTailTimer = $null
        }

        $t = New-Object System.Windows.Threading.DispatcherTimer
        $t.Interval = [TimeSpan]::FromMilliseconds(250)

        $t.Add_Tick({
            try {
                if (-not $global:ServerToolkitLogPath) { return }
                if (-not (Test-Path $global:ServerToolkitLogPath)) { return }

                $len = 0
                try { $len = (Get-Item -LiteralPath $global:ServerToolkitLogPath).Length } catch { return }
                if ($len -le $script:LogTailLastPos) { return }

                $fs = $null
                try {
                    $fs = [System.IO.File]::Open($global:ServerToolkitLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $null = $fs.Seek($script:LogTailLastPos, [System.IO.SeekOrigin]::Begin)

                    $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true, 4096, $true)
                    $newText = $sr.ReadToEnd()
                    $sr.Dispose()

                    $script:LogTailLastPos = $fs.Position
                } finally {
                    try { if ($fs) { $fs.Dispose() } } catch {}
                }

                if ([string]::IsNullOrWhiteSpace($newText)) { return }

                $script:SuppressFileWrite = $true
                try {
                    foreach ($ln in ($newText -split "`r?`n")) {
                        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                        $script:AppendLog.Invoke($ln + "`r`n")
                    }
                } finally {
                    $script:SuppressFileWrite = $false
                }
            } catch {
                try {
                    Add-Content -Path $global:ServerToolkitLogPath -Value ("[UIERR] LogTail failed: " + $_.Exception.ToString() + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
                } catch {}
            }
        })

        $script:LogTailTimer = $t
        $t.Start()
    } catch {}
}

function Stop-LogTail {
    try {
        if ($script:LogTailTimer) {
            try { $script:LogTailTimer.Stop() } catch {}
            $script:LogTailTimer = $null
        }
    } catch {}
}

$script:AppendLog.Invoke("DEBUG: DisableRootCheck initial IsChecked = '$($DisableRootCheck.IsChecked)'`r`n")
$script:useKeyAuth       = $false
$script:usePasswordAuth  = $false
$script:sudoPassword     = $null
$global:PlinkPath        = ""
$global:PscpPath         = ""

if (-not $script:OwnedAgentKeyPaths) {
    $script:OwnedAgentKeyPaths = New-Object System.Collections.Generic.List[string]
}

$script:SshAddUnlockProc        = $null
$script:SshAddAutoResumeRunning = $false

$script:SshAgentPrevStartType = $null
$script:SshAgentPrevStatus    = $null
$script:SshAgentSessionDir    = $null
$script:SshAgentReadyFile     = $null
$script:SshAgentCloseFile     = $null
$script:SshAgentHelperProc    = $null
$script:SshAgentTouched       = $false

$script:KeyCleanupWatchdogProc = $null

$script:TaskWatchTimer = $null
$script:IsBackgroundRunning = $false
$script:ActiveTask = $null

$script:TaskResult = [hashtable]::Synchronized(@{
    ServerHost = ""
    SshPort    = ""
    User       = ""
    Key        = ""
    Prompt     = 0

    Success    = 0

    RootDisabled = 0

    StatusTitle   = ""
    StatusMessage = ""
    StatusIcon    = "Information"
    ShowStatus    = 0

    P12Alias      = ""
    P12RemotePath = ""
    P12LocalPath  = ""
    P12NodeId     = ""
    ShowP12Alias  = 0
})

function Start-TaskWatchTimer {
    param([Parameter(Mandatory=$true)]$Task)

    $script:WatchedTask = $Task

    if ($script:TaskWatchTimer) {
        try { $script:TaskWatchTimer.Stop() } catch {}
        $script:TaskWatchTimer = $null
    }

    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(250)

    try {
        Add-Content -Path $global:ServerToolkitLogPath `
            -Value ("TASKWATCH_STARTED " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") `
            -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}

    $t.Add_Tick({
        param($sender, $e)

        $Task = $script:WatchedTask

        try {
            if (-not $Task -or -not $Task.PowerShell) { return }

            $state = $null
            try { $state = $Task.PowerShell.InvocationStateInfo.State } catch { $state = $null }

            $isDone = $false
            try {
                if ($state -in @("Completed","Failed","Stopped")) { $isDone = $true }
            } catch {}

            try {
                if (-not $isDone -and $Task.AsyncResult -and $Task.AsyncResult.IsCompleted) { $isDone = $true }
            } catch {}

            if (-not $isDone) { return }

            try {
                Add-Content -Path $global:ServerToolkitLogPath `
                    -Value ("TASKWATCH_DONE State=" + $state + " " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") `
                    -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}

            try { $sender.Stop() } catch { try { $t.Stop() } catch {} }

            try {
                if ($Task.AsyncResult) {
                    $null = $Task.PowerShell.EndInvoke($Task.AsyncResult)
                }
                $Task.Status = "Completed"
            }
            catch {
                $Task.Status    = "Faulted"
                $Task.IsFaulted = $true
                $Task.Exception = $_.Exception
                try {
                    Add-Content -Path $global:ServerToolkitLogPath `
                        -Value ("BG_ENDINVOKE_EXCEPTION`r`n" + $_.Exception.ToString() + "`r`n") `
                        -Encoding UTF8 -ErrorAction SilentlyContinue
                } catch {}
            }
            finally {
                try { $Task.PowerShell.Dispose() } catch {}
                try { $Task.Runspace.Close(); $Task.Runspace.Dispose() } catch {}

                try {
                    if ($script:TaskResult -and $script:TaskResult.ShowStatus -eq 1 -and
                        -not [string]::IsNullOrWhiteSpace($script:TaskResult.StatusMessage)) {

                        $icon = [System.Windows.MessageBoxImage]::Information
                        switch -Regex ($script:TaskResult.StatusIcon) {
                            '^Error$'   { $icon = [System.Windows.MessageBoxImage]::Error }
                            '^Warning$' { $icon = [System.Windows.MessageBoxImage]::Warning }
                            default     { $icon = [System.Windows.MessageBoxImage]::Information }
                        }

                        [System.Windows.MessageBox]::Show(
                            $MainWindow,
                            $script:TaskResult.StatusMessage,
                            $(if ($script:TaskResult.StatusTitle) { $script:TaskResult.StatusTitle } else { "Info" }),
                            [System.Windows.MessageBoxButton]::OK,
                            $icon
                        ) | Out-Null
                    }
                } catch {
                    try {
                        Add-Content -Path $global:ServerToolkitLogPath `
                            -Value ("[UIERR] Status dialog failed: " + $_.Exception.ToString() + "`r`n") `
                            -Encoding UTF8 -ErrorAction SilentlyContinue
                    } catch {}
                }

                try {
                    if ($script:TaskResult -and $script:TaskResult.ShowP12Alias -eq 1) {
                        try {
                            Show-P12AliasDialog `
                                -Alias $script:TaskResult.P12Alias `
                                -RemotePath $script:TaskResult.P12RemotePath `
                                -LocalPath $script:TaskResult.P12LocalPath `
                                -NodeId $script:TaskResult.P12NodeId
                        }
                        catch {
                            try {
                                Add-Content -Path $global:ServerToolkitLogPath -Value (
                                    (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                                    " [UIERR] Show-P12AliasDialog failed: " +
                                    $_.Exception.ToString() + "`r`n"
                                ) -Encoding UTF8 -ErrorAction SilentlyContinue
                            } catch {}
                        }
                    }
                }
                catch {
                    try {
                        Add-Content -Path $global:ServerToolkitLogPath -Value (
                            (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                            " [UIERR] P12 dialog wrapper failed: " +
                            $_.Exception.ToString() + "`r`n"
                        ) -Encoding UTF8 -ErrorAction SilentlyContinue
                    } catch {}
                }

                try {
                    if ($script:OfferSaveProfileFunc -and $script:TaskResult -and
                        -not [string]::IsNullOrWhiteSpace($script:TaskResult.ServerHost) -and
                        ([int]$script:TaskResult.Success -eq 1) -and
                        ([int]$script:TaskResult.Prompt -eq 1)) {

                        $pf = 0
                        try { $pf = [int]$script:TaskResult.Prompt } catch { $pf = 0 }

                        try {
                            Add-Content -Path $global:ServerToolkitLogPath `
                                -Value ("PROFILE_PROMPT_CHECK pf=" + $pf + " host=" + $script:TaskResult.ServerHost + " user=" + $script:TaskResult.User + "`r`n") `
                                -Encoding UTF8 -ErrorAction SilentlyContinue
                        } catch {}

                        $result = $script:OfferSaveProfileFunc.Invoke(
                            "",
                            $script:TaskResult.ServerHost,
                            $script:TaskResult.SshPort,
                            $script:TaskResult.User,
                            $script:TaskResult.Key,
                            $pf
                        )

                        if ($result -like "SAVED:*") {
                            try { $script:AppendLog.Invoke("Profile saved: $($result.Substring(6))`r`n") } catch {}
                        } elseif ($result -like "ERROR:*") {
                            try { $script:AppendLog.Invoke("[WARN] Profile save failed: $result`r`n") } catch {}
                        }
                    }
                } catch {
                    try {
                        Add-Content -Path $global:ServerToolkitLogPath `
                            -Value ("[UIERR] OfferSaveProfileFunc failed: " + $_.Exception.ToString() + "`r`n") `
                            -Encoding UTF8 -ErrorAction SilentlyContinue
                    } catch {}
                }

                try {
                    $rd = 0
                    $newUiUser = ""

                    try { if ($script:TaskResult) { $rd = [int]$script:TaskResult.RootDisabled } } catch { $rd = 0 }
                    try { if ($script:TaskResult) { $newUiUser = [string]$script:TaskResult.User } } catch { $newUiUser = "" }

                    if ($rd -eq 1 -and -not [string]::IsNullOrWhiteSpace($newUiUser) -and $newUiUser -ne "root") {
                        try { if ($script:SetUsernameAction) { $script:SetUsernameAction.Invoke($newUiUser) } } catch {}
                    }
                } catch {}

                try {
                    Cleanup-AgentOwnedKeys
                } catch {
                    try {
                        Add-Content -Path $global:ServerToolkitLogPath `
                            -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " [WARN] Immediate agent cleanup failed: " + $_.Exception.Message + "`r`n") `
                            -Encoding UTF8 -ErrorAction SilentlyContinue
                    } catch {}
                }

                try { Signal-SshAgentSessionEnd } catch {}
                try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
                try { $script:IsBackgroundRunning = $false } catch {}

                try {
                    Add-Content -Path $global:ServerToolkitLogPath `
                        -Value ("[INFO] TASK_UI_UNLOCKED State=" + $state + " " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") `
                        -Encoding UTF8 -ErrorAction SilentlyContinue
                } catch {}

                try { $script:WatchedTask = $null } catch {}
                try { $script:TaskWatchTimer = $null } catch {}
            }
        }
        catch {
            try {
                Add-Content -Path $global:ServerToolkitLogPath `
                    -Value ("[UIERR] TaskWatchTimer tick failed:`r`n" + $_.Exception.ToString() + "`r`n") `
                    -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
        }
    })

    $script:TaskWatchTimer = $t
    $t.Start()
}

function Stop-ActiveBackgroundTask {
    param([string]$Reason = "user-close")

    try { Add-Content -Path $global:ServerToolkitLogPath -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " [WARN] Stop-ActiveBackgroundTask reason=$Reason`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}

    try { Stop-LogTail } catch {}
    try {
        if ($script:TaskWatchTimer) {
            try { $script:TaskWatchTimer.Stop() } catch {}
            $script:TaskWatchTimer = $null
        }
    } catch {}

    try {
        $t = $script:ActiveTask
        if ($t) {

            try {
                if ($t.PowerShell) {
                    try { $t.PowerShell.Stop() } catch {}
                }
            } catch {}

            try {
                if ($t.Runspace) {
                    try { $t.Runspace.Close() } catch {}
                    try { $t.Runspace.Dispose() } catch {}
                }
            } catch {}

            try { if ($t.PowerShell) { $t.PowerShell.Dispose() } } catch {}
        }
    } catch {}

    try { $script:ActiveTask = $null } catch {}
    try { $script:IsBackgroundRunning = $false } catch {}
    try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
}

function Invoke-BackgroundUI {
    param(
        [Parameter(Mandatory=$true)]
        [ScriptBlock]$Work,

        [hashtable]$Context = $null
    )

    try {
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions  = 'ReuseThread'
        $rs.Open()

        $rs.SessionStateProxy.SetVariable("ServerToolkitLogPath", $global:ServerToolkitLogPath)
        $rs.SessionStateProxy.SetVariable("UiDispatcher", $script:UiDispatcher)

        $rs.SessionStateProxy.SetVariable("UnlockUIAction",        $script:UnlockUIAction)
        $rs.SessionStateProxy.SetVariable("SetUsernameAction",     $script:SetUsernameAction)
        $rs.SessionStateProxy.SetVariable("ShowMsgAction",         $script:ShowMsgAction)
        $rs.SessionStateProxy.SetVariable("UserExistsPromptFunc",  $script:UserExistsPromptFunc)
        $rs.SessionStateProxy.SetVariable("P12OverwritePromptFunc",$script:P12OverwritePromptFunc)

        $rs.SessionStateProxy.SetVariable("LogBox", $LogBox)
        $rs.SessionStateProxy.SetVariable("GuiLogLevels", $global:GuiLogLevels)
        $rs.SessionStateProxy.SetVariable("GuiLogRegex", $global:GuiLogRegex)
        $rs.SessionStateProxy.SetVariable("GuiShowDebug", $global:GuiShowDebug)
        $rs.SessionStateProxy.SetVariable("GuiStripTsRegex", $global:GuiStripTsRegex)

        if ($Context) {
            foreach ($k in $Context.Keys) {
                try { $rs.SessionStateProxy.SetVariable($k, $Context[$k]) } catch {}
            }
        }

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs

        $ps.AddScript({
            try {
                $AppendLog = [Action[string]]{
                    param([string]$t)
                    try {
                        if (-not $ServerToolkitLogPath) { return }

                        $stamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
                        $line = $t
                        if ($null -eq $line) { $line = "" }

                        if ($line -notmatch "(\r?\n)$") { $line += "`r`n" }
                        if ($line -notmatch "^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}\s+") {
                            $line = "$stamp $line"
                        }

                        try {
                            $bytes = [System.Text.Encoding]::UTF8.GetBytes($line)
                            $fs = [System.IO.File]::Open(
                                $ServerToolkitLogPath,
                                [System.IO.FileMode]::Append,
                                [System.IO.FileAccess]::Write,
                                [System.IO.FileShare]::ReadWrite
                            )
                            try {
                                $fs.Write($bytes, 0, $bytes.Length)
                                $fs.Flush()
                            } finally {
                                $fs.Dispose()
                            }
                        } catch {}
                    } catch {}
                }

                Add-Content -Path $ServerToolkitLogPath `
                    -Value ("BG_RUNSPACE_ENTERED " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") `
                    -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
        }) | Out-Null

        $funcImport = @()
        $funcImport += "function Reset-HostKeysForServer { $(${function:Reset-HostKeysForServer}.ToString()) }"
        $funcImport += "function Ensure-PuttyTools { $(${function:Ensure-PuttyTools}.ToString()) }"
        $funcImport += "function Run-SshCommand { $(${function:Run-SshCommand}.ToString()) }"

        $funcImport += "function Get-ServerToolkitAgentSessionDir { $(${function:Get-ServerToolkitAgentSessionDir}.ToString()) }"
        $funcImport += "function Get-ServerToolkitOwnedPubKeysFile { $(${function:Get-ServerToolkitOwnedPubKeysFile}.ToString()) }"
        $funcImport += "function Get-AgentPubKeyLines { $(${function:Get-AgentPubKeyLines}.ToString()) }"
        $funcImport += "function Write-OwnedAgentPubKeys { $(${function:Write-OwnedAgentPubKeys}.ToString()) }"
        $funcImport += "function Read-OwnedAgentPubKeys { $(${function:Read-OwnedAgentPubKeys}.ToString()) }"
        $funcImport += "function Remove-OwnedAgentPubKeyLine { $(${function:Remove-OwnedAgentPubKeyLine}.ToString()) }"

        $funcImport += "function Ensure-SshAgentWithKey { $(${function:Ensure-SshAgentWithKey}.ToString()) }"
        $funcImport += "function Cleanup-AgentOwnedKeys { $(${function:Cleanup-AgentOwnedKeys}.ToString()) }"

        $ps.AddScript(($funcImport -join "`r`n")) | Out-Null

        $ps.AddScript($Work) | Out-Null

        $async = $ps.BeginInvoke()

        return [pscustomobject]@{
            Id          = Get-Random -Minimum 1000 -Maximum 9999
            Status      = "Running"
            IsFaulted   = $false
            Exception   = $null
            PowerShell  = $ps
            Runspace    = $rs
            AsyncResult = $async
        }
    }
    catch {
        try {
            Add-Content -Path $global:ServerToolkitLogPath `
                -Value ("INVOKE_BACKGROUNDUI_FAILED`r`n" + $_.Exception.ToString()) `
                -Encoding UTF8
        } catch {}
        return $null
    }
}

function Ensure-PuttyTools {
    $global:PlinkPath = ""
    $global:PscpPath  = ""

    try {
        $plinkCmd = Get-Command plink.exe -ErrorAction Stop
        $global:PlinkPath = $plinkCmd.Source
    } catch {}

    try {
        $pscpCmd = Get-Command pscp.exe -ErrorAction Stop
        $global:PscpPath = $pscpCmd.Source
    } catch {}

    if ($global:PlinkPath -and (Test-Path -LiteralPath $global:PlinkPath) -and
        $global:PscpPath -and (Test-Path -LiteralPath $global:PscpPath)) {
        return $true
    }

    $AppendLog.Invoke("[INFO] PuTTY tools not found (plink/pscp). Downloading for password-only mode...`r`n")

    $arch = if ([IntPtr]::Size -eq 8) { "w64" } else { "w32" }
    $baseUrl = "https://the.earth.li/~sgtatham/putty/latest/$arch"

    $tempPath = [IO.Path]::GetTempPath()
    $plinkTarget = Join-Path $tempPath "plink.exe"
    $pscpTarget  = Join-Path $tempPath "pscp.exe"

    try {
        if (-not ($global:PlinkPath -and (Test-Path -LiteralPath $global:PlinkPath))) {
            $plinkUrl = "$baseUrl/plink.exe"
            (New-Object Net.WebClient).DownloadFile($plinkUrl, $plinkTarget)
            if (-not (Test-Path -LiteralPath $plinkTarget)) {
                $AppendLog.Invoke("[ERROR] plink.exe download completed but file not found at: $plinkTarget`r`n")
                return $false
            }
            $global:PlinkPath = $plinkTarget
            $AppendLog.Invoke("Downloaded plink.exe to $plinkTarget`r`n")
        }

        if (-not ($global:PscpPath -and (Test-Path -LiteralPath $global:PscpPath))) {
            $pscpUrl = "$baseUrl/pscp.exe"
            (New-Object Net.WebClient).DownloadFile($pscpUrl, $pscpTarget)
            if (-not (Test-Path -LiteralPath $pscpTarget)) {
                $AppendLog.Invoke("[ERROR] pscp.exe download completed but file not found at: $pscpTarget`r`n")
                return $false
            }
            $global:PscpPath = $pscpTarget
            $AppendLog.Invoke("Downloaded pscp.exe to $pscpTarget`r`n")
        }

        return $true
    }
    catch {
        $AppendLog.Invoke("[ERROR] Failed to download PuTTY tools (plink/pscp).`r`n")
        $AppendLog.Invoke("[ERROR] Exception: $($_.Exception.Message)`r`n")
        return $false
    }
}

function Reset-HostKeysForServer {
    param(
        [Parameter(Mandatory=$true)][string]$HostName,
        [Parameter(Mandatory=$true)][string]$Port
    )

    try {
        $regPath = "HKCU:\Software\SimonTatham\PuTTY\SshHostKeys"
        if (Test-Path $regPath) {
            $regItem       = Get-Item $regPath
            $targetPrefix  = "*@${Port}:${HostName}"
            $propsToRemove = @($regItem.Property | Where-Object { $_ -like $targetPrefix })

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
        $sshDir = Join-Path $env:USERPROFILE ".ssh"
        if (-not (Test-Path $sshDir)) {
            New-Item -ItemType Directory -Path $sshDir | Out-Null
        }

        $knownHostsPath = Join-Path $sshDir "known_hosts"
        if (-not (Test-Path $knownHostsPath)) {
            return
        }

        $sshKeygen = Join-Path $env:WINDIR "System32\OpenSSH\ssh-keygen.exe"
        if (-not (Test-Path $sshKeygen)) {
            try { $sshKeygen = (Get-Command ssh-keygen -ErrorAction Stop).Source } catch { $sshKeygen = $null }
        }

        if ($sshKeygen) {
            $AppendLog.Invoke("Resetting OpenSSH known_hosts entries for ${HostName}:${Port} via ssh-keygen -R...`r`n")

            try { & $sshKeygen -R "$HostName" -f "$knownHostsPath" 2>$null | Out-Null } catch {}
            try { & $sshKeygen -R "[$HostName]:$Port" -f "$knownHostsPath" 2>$null | Out-Null } catch {}

            $AppendLog.Invoke("Reset OpenSSH known_hosts entries for $HostName (port $Port).`r`n")
        }
        else {
            $AppendLog.Invoke("[WARN] ssh-keygen not found; using regex cleanup (hashed known_hosts entries may remain).`r`n")

            $allLines    = Get-Content -Path $knownHostsPath
            $escapedHost = [regex]::Escape($HostName)
            $escapedPort = [regex]::Escape($Port)
            $pattern1 = "^$escapedHost[\s,]"
            $pattern2 = "^\[$escapedHost\]:$escapedPort[\s,]"

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

function Try-CollectRemoteSnapshot {
    param(
        [Parameter(Mandatory=$true)][string]$SshExe,
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][string]$User,
        [string]$KeyPath,
        [string]$Reason = "timeout"
    )

    if ($script:__SnapshotInProgress) { return }
    $script:__SnapshotInProgress = $true

    try {
        $sshDir = Join-Path $env:USERPROFILE ".ssh"
        $knownHostsPath = Join-Path $sshDir "known_hosts"

        $snapCmd = @"
bash -lc '
echo "=== SNAPSHOT_START reason=`$Reason ts=`$(date -Is) ==="
echo "whoami=`$(whoami 2>/dev/null || true)"
echo "id=`$(id 2>/dev/null || true)"
echo "--- uname ---"
uname -a 2>/dev/null || true
echo "--- uptime ---"
uptime 2>/dev/null || true
echo "--- nsswitch ---"
cat /etc/nsswitch.conf 2>/dev/null || true
echo "--- resolv ---"
( command -v resolvectl >/dev/null 2>&1 && resolvectl status 2>/dev/null ) || true
cat /etc/resolv.conf 2>/dev/null || true
echo "--- processes (top 25 by elapsed) ---"
ps -eo pid,ppid,stat,etime,cmd --sort=-etime 2>/dev/null | head -n 25 || true
echo "--- last auth/syslog (tail) ---"
( test -f /var/log/auth.log && tail -n 40 /var/log/auth.log ) 2>/dev/null || true
( test -f /var/log/syslog && tail -n 40 /var/log/syslog ) 2>/dev/null || true
echo "--- journalctl (last 30) ---"
journalctl -n 30 --no-pager 2>/dev/null || true
echo "=== SNAPSHOT_END ==="
'
"@

        $args = @(
            "-o","BatchMode=yes",
            "-o","NumberOfPasswordPrompts=0",
            "-o","PreferredAuthentications=publickey",
            "-o","PubkeyAuthentication=yes",
            "-o","PasswordAuthentication=no",
            "-o","KbdInteractiveAuthentication=no",
            "-o","ChallengeResponseAuthentication=no",
            "-o","StrictHostKeyChecking=accept-new",
            "-o","ConnectTimeout=5",
            "-o","ServerAliveInterval=2",
            "-o","ServerAliveCountMax=1",
            "-o","UserKnownHostsFile=$knownHostsPath",
            "-o","GlobalKnownHostsFile=NUL",
            "-p",$Port
        )

        if (-not [string]::IsNullOrWhiteSpace($KeyPath)) {
            $args += @("-i", "`"$KeyPath`"")
        }

        $args += "$User@$RemoteHost"
        $snapCmd = $snapCmd.Replace('"','\"')
        $args += "`"$snapCmd`""

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $SshExe
        $psi.Arguments = ($args -join " ")
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $null = $p.Start()

        if (-not $p.WaitForExit(8000)) {
            try { $p.Kill() } catch {}
            try { $p.WaitForExit(1000) | Out-Null } catch {}
            try { $AppendLog.Invoke("[WARN] Snapshot ssh also timed out (8s).`r`n") } catch {}
            return
        }

        $o = ""
        $e = ""
        try { $o = $p.StandardOutput.ReadToEnd() } catch {}
        try { $e = $p.StandardError.ReadToEnd() } catch {}

        if ($o) { try { $AppendLog.Invoke("__GUIONLY__" + $o + "`r`n") } catch {} }
        if ($e) { try { $AppendLog.Invoke("__GUIONLY__" + "[SNAPSHOT_STDERR] " + $e + "`r`n") } catch {} }

    } catch {
        try { $AppendLog.Invoke("[WARN] Snapshot collection failed: $($_.Exception.Message)`r`n") } catch {}
    } finally {
        $script:__SnapshotInProgress = $false
    }
}

function Run-SshCommand {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteCommand,
        [Parameter(Mandatory=$true)][string]$RemoteHost,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][string]$User,
        [string]$KeyPath,
        [int]$TimeoutMs = 60000,
        [string]$StdinText = $null
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

    $sshDir = Join-Path $env:USERPROFILE ".ssh"
    if (-not (Test-Path $sshDir)) {
        New-Item -ItemType Directory -Path $sshDir | Out-Null
    }

    $knownHostsPath = Join-Path $sshDir "known_hosts"

    $args = @(
        "-o", "BatchMode=yes",
        "-o", "NumberOfPasswordPrompts=0",
        "-o", "PreferredAuthentications=publickey",
        "-o", "PubkeyAuthentication=yes",
        "-o", "PasswordAuthentication=no",
        "-o", "KbdInteractiveAuthentication=no",
        "-o", "ChallengeResponseAuthentication=no",

        "-o", "StrictHostKeyChecking=accept-new",
        "-o", "ConnectTimeout=10",
        "-o", "ServerAliveInterval=10",
        "-o", "ServerAliveCountMax=3",
        "-o", "UserKnownHostsFile=$knownHostsPath",
        "-o", "GlobalKnownHostsFile=NUL",
        "-p", $Port
    )

    if (-not [string]::IsNullOrWhiteSpace($KeyPath)) {
        $args += @("-i", "`"$KeyPath`"")
    }

    $args += "$User@$RemoteHost"

    $rc = $RemoteCommand
    if ($null -eq $rc) { $rc = "" }

    $rc = $rc.Replace('"', '\"')
    $args += "`"$rc`""

    $fullArgs = ($args -join ' ')
    if ($fullArgs -like "*chpasswd*" -or $fullArgs -like "*base64 -d*") {
        $AppendLog.Invoke("DEBUG: Running ssh ($sshExe) with password update command (details redacted).`r`n")
    } else {
        $AppendLog.Invoke("DEBUG: Running ssh command (arguments suppressed).`r`n")
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $sshExe
        $psi.Arguments              = ($args -join " ")
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.RedirectStandardInput  = $true
        $psi.CreateNoWindow         = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi

        $null = $proc.Start()

        try {
            if ($StdinText -ne $null) {
                $txt = [string]$StdinText
                if (-not $txt.EndsWith("`n")) { $txt += "`n" }
                $proc.StandardInput.Write($txt)
            }
        } catch {}
        try { $proc.StandardInput.Close() } catch {}

        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()

        $timeoutMs = $TimeoutMs
        if ($RemoteCommand -like "*chpasswd*" -or $RemoteCommand -like "*base64 -d*") {
            $timeoutMs = [Math]::Max($timeoutMs, 180000)
        }

        $finished = $proc.WaitForExit($timeoutMs)
        if (-not $finished) {
            $AppendLog.Invoke("[ERROR] ssh command timed out after ${timeoutMs}ms. Killing ssh.exe...`r`n")
            try { $proc.Kill() } catch {}
            try { $proc.WaitForExit(2000) | Out-Null } catch {}

            try {
                $AppendLog.Invoke("[WARN] Collecting remote snapshot after ssh timeout...`r`n")
                Try-CollectRemoteSnapshot -SshExe $sshExe -RemoteHost $RemoteHost -Port $Port -User $User -KeyPath $KeyPath -Reason "ssh-timeout"
            } catch {}

            return 1
        }

        $stdout = ""
        $stderr = ""
        try { $null = [System.Threading.Tasks.Task]::WaitAll(@($outTask, $errTask), 2000) } catch {}

        try { $stdout = $outTask.Result } catch { $stdout = "" }
        try { $stderr = $errTask.Result } catch { $stderr = "" }

        if ($stderr) {
            foreach ($line in $stderr -split "`r?`n") {
                if ($line -ne "") {
                    if ($line -like "Warning: Permanently added*") {
                        $AppendLog.Invoke("First-time connection to $RemoteHost confirmed. Host key saved to known_hosts.`r`n")
                    } else {
                        $AppendLog.Invoke("[ERROR] $line`r`n")
                    }
                }
            }
        }

        if ($stdout) {
            foreach ($line in $stdout -split "`r?`n") {
                if ($line -ne "") {
                    if ($RemoteCommand -like "*echo OK*" -and $line -eq "OK") { continue }
                    $AppendLog.Invoke("$line`r`n")
                }
            }
        }

        try {
            $script:LastSshExitCode  = $proc.ExitCode
            $script:LastSshStdErr    = $stderr
            $script:LastSshStdOut    = $stdout
            $script:LastSshErrorLine = $null
            if ($stderr) {
                $lines = $stderr -split "`r?`n" | Where-Object { $_ -and $_.Trim() -ne "" }
                if ($lines.Count -gt 0) { $script:LastSshErrorLine = $lines[-1].Trim() }
            }
        } catch {}

        return $proc.ExitCode
    }
    catch {
        $AppendLog.Invoke("[FATAL] Exception while running ssh: $($_.Exception.Message)`r`n")
        return 1
    }
}

function Normalize-FullPath {
    param([string]$p)
    try {
        if ([string]::IsNullOrWhiteSpace($p)) { return "" }
        return ([IO.Path]::GetFullPath($p)).Trim()
    } catch {
        try { return ($p.Trim()) } catch { return "" }
    }
}

function Convert-AgentCommentToPath {
    param($CommentValue)

    try {
        if ($null -eq $CommentValue) { return "" }

        if ($CommentValue -is [string]) { return $CommentValue }

        if ($CommentValue -is [byte[]]) {
            $s = [System.Text.Encoding]::UTF8.GetString($CommentValue)
            return ($s -replace "`0","").Trim()
        }

        return ([string]$CommentValue).Trim()
    } catch {
        return ""
    }
}

function Get-AgentRegistryKeyForPath {
    param([string]$KeyPath)

    $base = "HKCU:\Software\OpenSSH\Agent\Keys"
    if (-not (Test-Path $base)) { return $null }
    if ([string]::IsNullOrWhiteSpace($KeyPath)) { return $null }

    $normalized = $null
    try { $normalized = ([IO.Path]::GetFullPath($KeyPath)).Trim() } catch { $normalized = $KeyPath.Trim() }
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }

    foreach ($sub in Get-ChildItem $base -ErrorAction SilentlyContinue) {
        try {
            $props = Get-ItemProperty -Path $sub.PSPath -ErrorAction SilentlyContinue
            $c = Convert-AgentCommentToPath $props.comment
            if ([string]::IsNullOrWhiteSpace($c)) { continue }

            $cFull = $null
            try { $cFull = ([IO.Path]::GetFullPath($c)).Trim() } catch { $cFull = $c.Trim() }

            if ($cFull -and ($cFull -ieq $normalized)) {
                return $sub.PSPath
            }
        } catch {}
    }

    return $null
}

function Get-ServerToolkitAgentSessionDir {
    try {
        if ($script:SshAgentSessionDir -and (Test-Path -LiteralPath $script:SshAgentSessionDir)) {
            return $script:SshAgentSessionDir
        }
    } catch {}
    try {
        if ($env:SERVER_TOOLKIT_SSHAGENT_SESSIONDIR -and (Test-Path -LiteralPath $env:SERVER_TOOLKIT_SSHAGENT_SESSIONDIR)) {
            return $env:SERVER_TOOLKIT_SSHAGENT_SESSIONDIR
        }
    } catch {}
    return $null
}

function Get-ServerToolkitOwnedPubKeysFile {
    try {
        $base = Join-Path $env:LOCALAPPDATA "ServerToolkit"
        if (-not (Test-Path -LiteralPath $base)) {
            New-Item -ItemType Directory -Path $base -Force | Out-Null
        }
        return (Join-Path $base "owned_agent_pubkeys.txt")
    } catch {
        try { return (Join-Path $env:TEMP "ServerToolkit-owned_agent_pubkeys.txt") } catch {}
    }
    return $null
}

function Get-AgentPubKeyLines {
    $sshAddExe = $null
    $win = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
    if (Test-Path -LiteralPath $win) { $sshAddExe = $win } else {
        try { $sshAddExe = (Get-Command ssh-add -ErrorAction Stop).Source } catch { return @() }
    }

    try {
        $out = & $sshAddExe -L 2>&1
        if (-not $out) { return @() }

        $lines = @()
        foreach ($ln in ($out -split "`r?`n")) {
            $t = ($ln | ForEach-Object { $_.Trim() })
            if ([string]::IsNullOrWhiteSpace($t)) { continue }
            if ($t -match '^(ssh-(rsa|ed25519)|ecdsa-sha2-nistp\d+)\s+') {
                $lines += $t
            }
        }
        return $lines
    } catch {
        return @()
    }
}

function Write-OwnedAgentPubKeys {
    param([string[]]$PubKeyLines)

    if (-not $PubKeyLines -or $PubKeyLines.Count -eq 0) { return }

    $f = Get-ServerToolkitOwnedPubKeysFile
    if (-not $f) { return }

    try {
        $dir = Split-Path -Parent $f
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $existing = @()
        if (Test-Path -LiteralPath $f) {
            try { $existing = Get-Content -LiteralPath $f -ErrorAction SilentlyContinue } catch { $existing = @() }
        }

        $added = 0
        foreach ($ln in $PubKeyLines) {
            if ([string]::IsNullOrWhiteSpace($ln)) { continue }
            if ($existing -notcontains $ln) {
                Add-Content -LiteralPath $f -Value $ln -Encoding ASCII
                $added++
            }
        }

        try { $AppendLog.Invoke("[INFO] OWNED_KEYS: wrote $added new key(s) to: $f`r`n") } catch {}
    } catch {
        try { $AppendLog.Invoke("[WARN] OWNED_KEYS: failed writing owned key file: $($_.Exception.Message)`r`n") } catch {}
    }
}

function Read-OwnedAgentPubKeys {
    $f = Get-ServerToolkitOwnedPubKeysFile
    if (-not $f -or -not (Test-Path -LiteralPath $f)) { return @() }

    try {
        return (Get-Content -LiteralPath $f -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() })
    } catch {
        return @()
    }
}

function Remove-OwnedAgentPubKeyLine {
    param(
        [Parameter(Mandatory=$true)][string]$PubKeyLine
    )

    $sshAddExe = $null
    $win = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
    if (Test-Path -LiteralPath $win) { $sshAddExe = $win } else {
        try { $sshAddExe = (Get-Command ssh-add -ErrorAction Stop).Source } catch { $sshAddExe = $null }
    }
    if (-not $sshAddExe) { return $false }

    $tmpPub = Join-Path $env:TEMP ("server-toolkit-agent-remove-" + [guid]::NewGuid().ToString("N") + ".pub")
    try {
        Set-Content -LiteralPath $tmpPub -Value $PubKeyLine -Encoding ASCII -Force
        $out = & $sshAddExe -d $tmpPub 2>&1
        if ($out) {
            foreach ($ln in ($out -split "`r?`n")) {
                if ($ln) { try { $AppendLog.Invoke("[INFO] CLEANUP: ssh-add -d: $ln`r`n") } catch {} }
            }
        }
        return $true
    } catch {
        return $false
    } finally {
        try { Remove-Item -LiteralPath $tmpPub -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Cleanup-AgentOwnedKeys {

    try { $AppendLog.Invoke("[INFO] CLEANUP: Starting ssh-agent key cleanup (pubkey-file method)...`r`n") } catch {}

    $sshAddExe = $null
    $win = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
    if (Test-Path -LiteralPath $win) { $sshAddExe = $win } else {
        try { $sshAddExe = (Get-Command ssh-add -ErrorAction Stop).Source } catch { $sshAddExe = $null }
    }

    if (-not $sshAddExe) {
        try { $AppendLog.Invoke("[WARN] CLEANUP: ssh-add not found. Skipping agent cleanup.`r`n") } catch {}
        return
    }

    $fOwned = Get-ServerToolkitOwnedPubKeysFile
    try { $AppendLog.Invoke("[INFO] CLEANUP: reading owned key file: $fOwned`r`n") } catch {}
    $owned = Read-OwnedAgentPubKeys

    if (-not $owned -or $owned.Count -eq 0) {
        try { $AppendLog.Invoke("[INFO] CLEANUP: No owned pubkeys recorded. Nothing to remove.`r`n") } catch {}
        try {
            $f = Get-ServerToolkitOwnedPubKeysFile
            if ($f -and (Test-Path -LiteralPath $f)) {
                Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
                try { $AppendLog.Invoke("[INFO] CLEANUP: removed stale owned key file: $f`r`n") } catch {}
            }
        } catch {}
        return
    }

    $removed = 0
    $failed  = 0

    foreach ($pubLine in $owned) {
        if ([string]::IsNullOrWhiteSpace($pubLine)) { continue }

        $ok = $false
        try { $ok = Remove-OwnedAgentPubKeyLine -PubKeyLine $pubLine } catch { $ok = $false }

        if ($ok) {
            $removed++
        } else {
            $failed++
            try { $AppendLog.Invoke("[WARN] CLEANUP: Failed removing one key via pubfile.`r`n") } catch {}
        }
    }

    try { $AppendLog.Invoke("[INFO] CLEANUP: Completed. removed=$removed failed=$failed`r`n") } catch {}

    try {
        $f = Get-ServerToolkitOwnedPubKeysFile
        if ($f -and (Test-Path -LiteralPath $f)) {
            Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Start-SshAddAutoResumeWaiter {
    param(
        [Parameter(Mandatory=$true)][System.Diagnostics.Process]$Proc,
        [Parameter(Mandatory=$true)][string]$KeyPath
    )

    if ($script:SshAddAutoResumeRunning) { return }
    $script:SshAddAutoResumeRunning = $true
    $script:SshAddUnlockProc = $Proc

    [System.Threading.Tasks.Task]::Run([Action]{
        try {
            try { $Proc.WaitForExit() } catch {}
            Start-Sleep -Milliseconds 300

            $loaded = $false
            try {
                $sshAddExe = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
                if (-not (Test-Path $sshAddExe)) {
                    try { $sshAddExe = (Get-Command ssh-add -ErrorAction Stop).Source } catch { $sshAddExe = $null }
                }

                if ($sshAddExe) {
                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $sshAddExe
                    $psi.Arguments = "-L"
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError  = $true
                    $psi.CreateNoWindow         = $true

                    $p2 = New-Object System.Diagnostics.Process
                    $p2.StartInfo = $psi
                    $p2.Start() | Out-Null
                    if ($p2.WaitForExit(3000)) {
                        $out = $p2.StandardOutput.ReadToEnd()
                        if ($p2.ExitCode -eq 0 -and $out) {
                            $esc = [regex]::Escape($KeyPath)
                            foreach ($ln in ($out -split "`r?`n")) {
                                if ($ln -match $esc) { $loaded = $true; break }
                            }
                        }
                    } else {
                        try { $p2.Kill() } catch {}
                    }
                }
            } catch { $loaded = $false }

            if (-not $loaded) {
                try { $script:AppendLog.Invoke("[WARN] ssh-add window closed, but the key still wasn't detected as loaded. Please try again.`r`n") } catch {}
                return
            }

            try {
                $script:UiDispatcher.BeginInvoke([Action]{
                    try {
                        Start-KeyCleanupWatchdog
                    } catch {}
                }) | Out-Null
            } catch {}

            try { $script:AppendLog.Invoke("[INFO] SSH key unlocked and detected. Continuing automatically...`r`n") } catch {}

            for ($i=0; $i -lt 200; $i++) {
                $okToClick = $false
                try {
                    $okToClick = $script:UiDispatcher.Invoke([Func[bool]]{
                        try { return ($StartButton -and $StartButton.IsEnabled -and (-not $script:IsBackgroundRunning)) } catch { return $false }
                    })
                } catch { $okToClick = $false }

                if ($okToClick) {
                    try {
                        $script:UiDispatcher.BeginInvoke([Action]{
                            try {
                                $StartButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
                            } catch {}
                        }) | Out-Null
                    } catch {}
                    break
                }

                Start-Sleep -Milliseconds 150
            }
        }
        finally {
            $script:SshAddAutoResumeRunning = $false
            $script:SshAddUnlockProc = $null
        }
    }) | Out-Null
}

function Start-KeyCleanupWatchdog {
    if ($script:KeyCleanupWatchdogProc -and -not $script:KeyCleanupWatchdogProc.HasExited) { return }

    $parentPid = $PID
    $log = $global:ServerToolkitLogPath

    $payload = @"
`$ErrorActionPreference='SilentlyContinue'
`$parentPid=$parentPid
`$logPath='$log'

function WL([string]`$m){
  try{
    if(`$logPath){
      Add-Content -Path `$logPath -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " [INFO] " + `$m) -Encoding UTF8 -ErrorAction SilentlyContinue
    }
  }catch{}
}

while(`$true){
  try{
    `$p=Get-Process -Id `$parentPid -ErrorAction SilentlyContinue
    if(-not `$p){ break }
  }catch{ break }
  Start-Sleep -Milliseconds 500
}

WL "KEY_CLEANUP_WATCHDOG_TRIGGERED parentPid=`$parentPid"

`$sshAddExe=Join-Path `$env:WINDIR "System32\OpenSSH\ssh-add.exe"
if(-not (Test-Path `$sshAddExe)){
  try{ `$sshAddExe=(Get-Command ssh-add -ErrorAction Stop).Source }catch{ `$sshAddExe=$null }
}
if(-not `$sshAddExe){
  WL "KEY_CLEANUP_WATCHDOG_SKIP ssh-add not found"
  exit 0
}

`$sess = `$env:SERVER_TOOLKIT_SSHAGENT_SESSIONDIR
if(-not `$sess -or -not (Test-Path -LiteralPath `$sess)){
  WL "KEY_CLEANUP_WATCHDOG_SKIP no session dir env/dir missing"
  exit 0
}

`$ownedFile = Join-Path `$sess "owned_agent_pubkeys.txt"
if(-not (Test-Path -LiteralPath `$ownedFile)){
  WL "KEY_CLEANUP_WATCHDOG_NONE no owned_agent_pubkeys.txt"
  exit 0
}

`$lines = @()
try{ `$lines = Get-Content -LiteralPath `$ownedFile -ErrorAction SilentlyContinue }catch{ `$lines=@() }

`$removed=0
`$failed=0

foreach(`$ln in `$lines){
  try{
    `$t = ([string]`$ln).Trim()
    if(-not `$t){ continue }

    `$tmpPub = Join-Path `$env:TEMP ("server-toolkit-agent-remove-" + [guid]::NewGuid().ToString("N") + ".pub")
    try{
      Set-Content -LiteralPath `$tmpPub -Value `$t -Encoding ASCII -Force
      `$out = & `$sshAddExe -d `$tmpPub 2>&1
      if(`$out){ WL ("WATCHDOG ssh-add -d: " + (`$out -join " | ")) }
      `$removed++
    }catch{
      `$failed++
      WL ("WATCHDOG remove failed: " + `$_.Exception.Message)
    }finally{
      try{ Remove-Item -LiteralPath `$tmpPub -Force -ErrorAction SilentlyContinue }catch{}
    }
  }catch{
    `$failed++
  }
}

try{ Remove-Item -LiteralPath `$ownedFile -Force -ErrorAction SilentlyContinue }catch{}

WL ("KEY_CLEANUP_WATCHDOG_DONE removed=" + `$removed + " failed=" + `$failed)
"@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($payload))

    try {
        $p = Start-Process powershell.exe -WindowStyle Hidden -PassThru `
            -ArgumentList "-NoProfile -NoLogo -ExecutionPolicy Bypass -EncodedCommand $encoded"
        $script:KeyCleanupWatchdogProc = $p
        $script:AppendLog.Invoke("[INFO] Key cleanup watchdog started (PID=$($p.Id)).`r`n")
    } catch {
        $script:AppendLog.Invoke("[WARN] Failed to start key cleanup watchdog: $($_.Exception.Message)`r`n")
    }
}

function Start-SshAgentSessionHelper {
    param(
        [Parameter(Mandatory=$true)][string]$PrevStartType,
        [Parameter(Mandatory=$true)][string]$PrevStatus,
        [int]$ParentPid = $PID,
        [int]$FailsafeMinutes = 30,
        [string]$LogPath = $global:ServerToolkitLogPath
    )

    $dir = Join-Path ([IO.Path]::GetTempPath()) ("server-toolkit-sshagent-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null

    $ready = Join-Path $dir "ready.txt"
    $close = Join-Path $dir "close.txt"

    $script:SshAgentSessionDir = $dir
    $script:SshAgentReadyFile  = $ready
    $script:SshAgentCloseFile  = $close

    if ($null -eq $LogPath) { $LogPath = "" }
    $LogPath = $LogPath.Trim()

    $helper = @"
`$ErrorActionPreference = 'SilentlyContinue'

`$prevStartType = '$PrevStartType'
`$prevStatus    = '$PrevStatus'
`$readyFile     = '$ready'
`$closeFile     = '$close'
`$parentPid     = $ParentPid
`$failsafeSec   = [int]($FailsafeMinutes * 60)
`$logPath       = '$LogPath'

function Write-HelperLog([string]`$msg) {
    try {
        if (-not [string]::IsNullOrWhiteSpace(`$logPath)) {
            `$ts = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
            Add-Content -Path `$logPath -Value ("`$ts [INFO] `$msg") -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Get-SshAddPath {
    try {
        `$p = Join-Path `$env:WINDIR "System32\OpenSSH\ssh-add.exe"
        if (Test-Path -LiteralPath `$p) { return `$p }
    } catch {}
    try {
        return (Get-Command ssh-add -ErrorAction Stop).Source
    } catch {
        return `$null
    }
}

function Convert-CommentToPath([object]`$commentValue) {
    try {
        if (`$null -eq `$commentValue) { return "" }
        if (`$commentValue -is [string]) { return (`$commentValue).Trim() }
        if (`$commentValue -is [byte[]]) {
            `$s = [System.Text.Encoding]::UTF8.GetString(`$commentValue)
            return (`$s -replace "`0","").Trim()
        }
        return ([string]`$commentValue).Trim()
    } catch {
        return ""
    }
}

function Cleanup-ServerToolkitOwnedAgentKeys {
    try {
        Write-HelperLog "HELPER_CLEANUP_BEGIN scanning ServerToolkitOwned agent keys..."
    } catch {}

    `$sshAddExe = Get-SshAddPath
    if (-not `$sshAddExe) {
        try { Write-HelperLog "HELPER_CLEANUP_SKIP ssh-add not found." } catch {}
        return
    }

    `$base = "HKCU:\Software\OpenSSH\Agent\Keys"
    if (-not (Test-Path -LiteralPath `$base)) {
        try { Write-HelperLog "HELPER_CLEANUP_NONE no HKCU:\Software\OpenSSH\Agent\Keys present." } catch {}
        return
    }

    `$removed = 0
    `$seen    = 0

    foreach (`$sub in Get-ChildItem -LiteralPath `$base -ErrorAction SilentlyContinue) {
        try {
            `$props = Get-ItemProperty -Path `$sub.PSPath -ErrorAction SilentlyContinue
            if (`$props.ServerToolkitOwned -ne 1) { continue }

            `$seen++
            `$commentPath = Convert-CommentToPath `$props.comment

            if (-not [string]::IsNullOrWhiteSpace(`$commentPath)) {
                try { & `$sshAddExe -d "`$commentPath" 1>`$null 2>`$null } catch {}
            }

            try { Remove-Item -LiteralPath `$sub.PSPath -Recurse -Force -ErrorAction SilentlyContinue } catch {}

            `$removed++
            try { Write-HelperLog ("HELPER_CLEANUP_REMOVED commentPath=" + `$commentPath) } catch {}
        } catch {}
    }

    try {
        Write-HelperLog ("HELPER_CLEANUP_DONE seenOwned=" + `$seen + " removed=" + `$removed)
    } catch {}
}

Write-HelperLog "HELPER_STARTED ssh-agent prevStartType=`$prevStartType prevStatus=`$prevStatus parentPid=`$parentPid failsafeSec=`$failsafeSec"

`$start = `$null

try {
    try { Set-Service ssh-agent -StartupType Manual } catch {}
    try { Start-Service ssh-agent } catch {}

    try {
        `$svc = Get-Service ssh-agent -ErrorAction SilentlyContinue
        Write-HelperLog ("HELPER_AGENT_STATE status=" + `$svc.Status + " startType=" + `$svc.StartType)
    } catch {}

    try { 'READY' | Out-File -FilePath `$readyFile -Encoding ASCII -Force } catch {}
    `$start = [DateTime]::UtcNow
    Write-HelperLog "HELPER_READY ssh-agent enabled/running (ready file written)"

    while (`$true) {

        try {
            if (`$null -eq `$start) { `$start = [DateTime]::UtcNow }
            `$elapsed = ([DateTime]::UtcNow - `$start).TotalSeconds
            if (`$elapsed -ge `$failsafeSec) {
                Write-HelperLog ("HELPER_TIMEOUT elapsedSec=" + [int]`$elapsed + " >= failsafeSec=" + `$failsafeSec)
                break
            }
        } catch {}

        try {
            if (Test-Path -LiteralPath `$closeFile) {
                Write-HelperLog "HELPER_CLOSE_SIGNAL file detected"
                break
            }
        } catch {}

        try {
            `$p = Get-Process -Id `$parentPid -ErrorAction SilentlyContinue
            if (-not `$p) {
                Write-HelperLog ("HELPER_PARENT_GONE parentPid=" + `$parentPid + " (GUI likely closed/crashed)")
                break
            }
        } catch {
            Write-HelperLog "HELPER_PARENT_CHECK_ERROR (treating as closed)"
            break
        }

        Start-Sleep -Milliseconds 500
    }

} finally {

    try {
        Cleanup-ServerToolkitOwnedAgentKeys
    } catch {
        try { Write-HelperLog ("HELPER_CLEANUP_FAILED " + `$_.Exception.Message) } catch {}
    }

    Write-HelperLog "HELPER_REVERTING restoring ssh-agent to prevStartType=`$prevStartType prevStatus=`$prevStatus"

    try { Set-Service ssh-agent -StartupType `$prevStartType } catch {}

    if (`$prevStatus -eq 'Running') {
        try { Start-Service ssh-agent -ErrorAction SilentlyContinue } catch {}
    } else {
        try { Stop-Service ssh-agent -ErrorAction SilentlyContinue } catch {}
    }

    try {
        `$svc2 = Get-Service ssh-agent -ErrorAction SilentlyContinue
        Write-HelperLog ("HELPER_REVERTED ssh-agent restored status=" + `$svc2.Status + " startType=" + `$svc2.StartType)
    } catch {
        Write-HelperLog "HELPER_REVERTED ssh-agent restored (state unknown)"
    }
}
"@

    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($helper))

    $p = Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -PassThru `
        -ArgumentList "-NoProfile -NoLogo -ExecutionPolicy Bypass -EncodedCommand $encoded"

    $script:SshAgentHelperProc = $p
    return $true
}

function Ensure-SshAgentSessionPreStart {
    param(
        [string]$KeyPath,
        [string]$Passphrase
    )

    if ([string]::IsNullOrWhiteSpace($KeyPath)) { return $true }
    if ([string]::IsNullOrWhiteSpace($Passphrase)) { return $true }
    $svc = $null
    try { $svc = Get-Service ssh-agent -ErrorAction SilentlyContinue } catch { $svc = $null }
    if (-not $svc) {
        $script:AppendLog.Invoke("[WARN] ssh-agent service not found on this system. Skipping agent enable step.`r`n")
        return $true
    }

    if (-not $script:SshAgentTouched) {
        $script:SshAgentPrevStartType = $svc.StartType.ToString()
        $script:SshAgentPrevStatus    = $svc.Status.ToString()
    }

    $needsEnable = ($svc.StartType -eq 'Disabled' -or $svc.Status -ne 'Running')

    if (-not $needsEnable) { return $true }

    $msg = @()
    $msg += "This tool needs to temporarily turn on Windows 'ssh-agent' so it can unlock your SSH key."
    $msg += ""
    $msg += "In plain terms: it stores your key securely for this session so the toolkit can connect as needed."
    $msg += ""
    $msg += "When you close the Server Toolkit, it will automatically restore ssh-agent back to how it was before."
    $msg += ""
    $msg += "Allow enabling ssh-agent now?"

    $res = [System.Windows.MessageBox]::Show(
        $MainWindow,
        ($msg -join "`r`n"),
        "Enable SSH Key Helper?",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($res -ne [System.Windows.MessageBoxResult]::Yes) {
        $script:AppendLog.Invoke("[ERROR] User declined enabling ssh-agent. Cannot continue with encrypted key automation.`r`n")
        return $false
    }

    $script:AppendLog.Invoke("[INFO] Enabling ssh-agent (requires UAC) for this session...`r`n")

    $ok = Start-SshAgentSessionHelper `
        -PrevStartType $script:SshAgentPrevStartType `
        -PrevStatus $script:SshAgentPrevStatus `
        -ParentPid $PID `
        -FailsafeMinutes 30 `
        -LogPath $global:ServerToolkitLogPath

    if (-not $ok) {
        $script:AppendLog.Invoke("[FATAL] Failed to start elevated ssh-agent helper.`r`n")
        return $false
    }

    $waitMs = 0
    while ($waitMs -lt 30000) {
        if (Test-Path -LiteralPath $script:SshAgentReadyFile) {
            $script:SshAgentTouched = $true

            $hp = $null
            try { if ($script:SshAgentHelperProc) { $hp = $script:SshAgentHelperProc.Id } } catch { $hp = $null }

            if ($hp) {
                $script:AppendLog.Invoke("[INFO] ssh-agent enabled and running (helper PID=$hp).`r`n")
            } else {
                $script:AppendLog.Invoke("[INFO] ssh-agent enabled and running.`r`n")
            }

            return $true
        }
        Start-Sleep -Milliseconds 250
        $waitMs += 250
    }

    $script:AppendLog.Invoke("[FATAL] Timed out waiting for ssh-agent helper to start.`r`n")
    return $false
}

function Signal-SshAgentSessionEnd {
    try {
        if ($script:SshAgentTouched -and $script:SshAgentCloseFile) {
            try { 'CLOSE' | Out-File -FilePath $script:SshAgentCloseFile -Encoding ASCII -Force } catch {}
        }
    } catch {}
}

function Ensure-SshAgentWithKey {
    param(
        [string]$KeyPath,
        [string]$Passphrase
    )

    if ([string]::IsNullOrWhiteSpace($KeyPath)) { return $true }

    $agentService = $null
    try { $agentService = Get-Service ssh-agent -ErrorAction SilentlyContinue } catch { $agentService = $null }
    if (-not $agentService) {
        try { $AppendLog.Invoke("[WARN] ssh-agent service not found on this system. Skipping agent integration.`r`n") } catch {}
        return $true
    }

    $sshAddExe = $null
    $winSshAdd = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
    if (Test-Path -LiteralPath $winSshAdd) {
        $sshAddExe = $winSshAdd
    } else {
        try { $sshAddExe = (Get-Command ssh-add -ErrorAction Stop).Source } catch { $sshAddExe = $null }
    }

    if (-not $sshAddExe) {
        try { $AppendLog.Invoke("[WARN] ssh-add not found. Cannot unlock key automatically.`r`n") } catch {}
        return $true
    }

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $sshAddExe
        $psi.Arguments = "-l"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.CreateNoWindow         = $true

        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        $p.Start() | Out-Null

        if ($p.WaitForExit(3000)) {
            $out = ""
            try { $out = $p.StandardOutput.ReadToEnd() } catch {}

            if ($p.ExitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($out)) {
                foreach ($ln in ($out -split "`r?`n")) {
                    if (-not $ln) { continue }
                    if ($ln -like "*$KeyPath*") {
                        try { $AppendLog.Invoke("ssh-agent already has this key loaded: $KeyPath`r`n") } catch {}
                        return $true
                    }
                }
            }
        } else {
            try { $p.Kill() } catch {}
        }
    } catch {}

    if ([string]::IsNullOrWhiteSpace($Passphrase)) {
        try { $AppendLog.Invoke("[WARN] Key is not detected as loaded and no passphrase was provided.`r`n") } catch {}
        try { $AppendLog.Invoke("[WARN] Provide the passphrase in the Password / SSH Key Passphrase field and click Start again.`r`n") } catch {}
        return $false
    }

    $beforePubLines = @()
    try { $beforePubLines = Get-AgentPubKeyLines } catch { $beforePubLines = @() }
    try { $AppendLog.Invoke("[INFO] OWNED_KEYS: baseline captured (count=$($beforePubLines.Count)).`r`n") } catch {}

    $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ("server-toolkit-askpass-" + [guid]::NewGuid().ToString("N"))
    try { New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null } catch {}

    $askpassPs1 = Join-Path $tmpDir "askpass.ps1"
    $askpassCmd = Join-Path $tmpDir "askpass.cmd"

    $pp = $Passphrase.Replace("'", "''")

    $psContent = @"
`$ErrorActionPreference = 'SilentlyContinue'
[Console]::Write(@'
$pp
'@)
"@

    $cmdContent = "@echo off`r`n" +
                  "`"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`" -NoProfile -ExecutionPolicy Bypass -File `"$askpassPs1`"`r`n"

    try { Set-Content -LiteralPath $askpassPs1 -Value $psContent -Encoding UTF8 -Force } catch {}
    try { Set-Content -LiteralPath $askpassCmd -Value $cmdContent -Encoding ASCII -Force } catch {}

    try { $AppendLog.Invoke("[INFO] Unlocking SSH key via ssh-agent...`r`n") } catch {}

    $exit = 1
    $stderr = ""

    try {
        $psi2 = New-Object System.Diagnostics.ProcessStartInfo
        $psi2.FileName = $sshAddExe
        $psi2.Arguments = "`"$KeyPath`""
        $psi2.UseShellExecute = $false
        $psi2.RedirectStandardInput  = $true
        $psi2.RedirectStandardOutput = $true
        $psi2.RedirectStandardError  = $true
        $psi2.CreateNoWindow = $true

        $psi2.EnvironmentVariables["SSH_ASKPASS"] = $askpassCmd
        $psi2.EnvironmentVariables["SSH_ASKPASS_REQUIRE"] = "force"
        $psi2.EnvironmentVariables["DISPLAY"] = "1"

        $p2 = New-Object System.Diagnostics.Process
        $p2.StartInfo = $psi2
        $p2.Start() | Out-Null

        try { $p2.StandardInput.Close() } catch {}

        if ($p2.WaitForExit(15000)) {
            $exit = $p2.ExitCode
            try { $stderr = $p2.StandardError.ReadToEnd() } catch { $stderr = "" }
        } else {
            try { $p2.Kill() } catch {}
            $exit = 1
            $stderr = "ssh-add timed out."
        }
    }
    catch {
        $exit = 1
        $stderr = $_.Exception.Message
    }
    finally {
        try { Remove-Item -LiteralPath $askpassCmd -Force -ErrorAction SilentlyContinue } catch {}
        try { Remove-Item -LiteralPath $askpassPs1 -Force -ErrorAction SilentlyContinue } catch {}
        try { Remove-Item -LiteralPath $tmpDir -Force -Recurse -ErrorAction SilentlyContinue } catch {}
    }

    if ($exit -ne 0) {
        if ($stderr) {
            $last = ($stderr -split "`r?`n" | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Last 1)
            if ($last) {
                try { $AppendLog.Invoke("[ERROR] ssh-add failed: $last`r`n") } catch {}
            } else {
                try { $AppendLog.Invoke("[ERROR] ssh-add failed (exit=$exit).`r`n") } catch {}
            }
        } else {
            try { $AppendLog.Invoke("[ERROR] ssh-add failed (exit=$exit).`r`n") } catch {}
        }
        return $false
    }

    try {
        if (-not $beforePubLines) { $beforePubLines = @() }

        Start-Sleep -Milliseconds 200

        $after = @()
        try { $after = Get-AgentPubKeyLines } catch { $after = @() }

        $delta = @()
        foreach ($ln in $after) {
            if ($beforePubLines -notcontains $ln) { $delta += $ln }
        }

        if ($delta.Count -gt 0) {
            Write-OwnedAgentPubKeys -PubKeyLines $delta
            try { $AppendLog.Invoke("[INFO] Recorded owned agent pubkey lines (delta count=$($delta.Count)).`r`n") } catch {}
        } else {
            $fallback = @()
            foreach ($ln in $after) {
                if ($ln -like "*$KeyPath") { $fallback += $ln }
            }

            if ($fallback.Count -gt 0) {
                Write-OwnedAgentPubKeys -PubKeyLines $fallback
                try { $AppendLog.Invoke("[WARN] No pubkey delta detected; recorded fallback match by path (count=$($fallback.Count)).`r`n") } catch {}
            } else {
                try { $AppendLog.Invoke("[WARN] No pubkey delta detected and no fallback match found. Cleanup may not remove this key.`r`n") } catch {}
            }
        }
    } catch {
        try { $AppendLog.Invoke("[WARN] Failed recording owned pubkey lines: $($_.Exception.Message)`r`n") } catch {}
    }

    try { Start-KeyCleanupWatchdog } catch {}

    try { $AppendLog.Invoke("[INFO] SSH key unlocked and loaded into ssh-agent.`r`n") } catch {}
    return $true
}

function New-ThemedButtonStyleSafe {
    param(
        [Parameter(Mandatory=$true)][string]$Normal,
        [Parameter(Mandatory=$true)][string]$Hover,
        [Parameter(Mandatory=$true)][string]$Pressed,
        [string]$ForegroundHex = "#FFFFFF"
    )

    $bc = [System.Windows.Media.BrushConverter]::new()

    $style = New-Object System.Windows.Style([System.Windows.Controls.Button])

    $style.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty,
        $bc.ConvertFromString($Normal)
    ))) | Out-Null

    $style.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::ForegroundProperty,
        $bc.ConvertFromString($ForegroundHex)
    ))) | Out-Null

    $style.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BorderThicknessProperty,
        (New-Object System.Windows.Thickness(0))
    ))) | Out-Null

    $style.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::PaddingProperty,
        (New-Object System.Windows.Thickness(14,6,14,6))
    ))) | Out-Null

    $style.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::CursorProperty,
        [System.Windows.Input.Cursors]::Hand
    ))) | Out-Null

    $tpl = New-Object System.Windows.Controls.ControlTemplate([System.Windows.Controls.Button])

    $border = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.Border])
    $border.SetValue(
        [System.Windows.Controls.Border]::CornerRadiusProperty,
        (New-Object System.Windows.CornerRadius(4))
    )
    $border.SetValue(
        [System.Windows.Controls.Border]::BackgroundProperty,
        (New-Object System.Windows.TemplateBindingExtension(
            [System.Windows.Controls.Control]::BackgroundProperty
        ))
    )

    $cp = New-Object System.Windows.FrameworkElementFactory([System.Windows.Controls.ContentPresenter])
    $cp.SetValue(
        [System.Windows.Controls.ContentPresenter]::HorizontalAlignmentProperty,
        [System.Windows.HorizontalAlignment]::Center
    )
    $cp.SetValue(
        [System.Windows.Controls.ContentPresenter]::VerticalAlignmentProperty,
        [System.Windows.VerticalAlignment]::Center
    )
    $cp.SetValue(
        [System.Windows.Controls.ContentPresenter]::MarginProperty,
        (New-Object System.Windows.Thickness(4,1,4,1))
    )

    $border.AppendChild($cp) | Out-Null
    $tpl.VisualTree = $border

    $tOver = New-Object System.Windows.Trigger
    $tOver.Property = [System.Windows.Controls.Control]::IsMouseOverProperty
    $tOver.Value = $true
    $tOver.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty,
        $bc.ConvertFromString($Hover)
    ))) | Out-Null

    $tPress = New-Object System.Windows.Trigger
    $tPress.Property = [System.Windows.Controls.Primitives.ButtonBase]::IsPressedProperty
    $tPress.Value = $true
    $tPress.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::BackgroundProperty,
        $bc.ConvertFromString($Pressed)
    ))) | Out-Null

    $tDis = New-Object System.Windows.Trigger
    $tDis.Property = [System.Windows.Controls.Control]::IsEnabledProperty
    $tDis.Value = $false
    $tDis.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.UIElement]::OpacityProperty,
        0.4
    ))) | Out-Null

    $tpl.Triggers.Add($tOver)  | Out-Null
    $tpl.Triggers.Add($tPress) | Out-Null
    $tpl.Triggers.Add($tDis)   | Out-Null

    $style.Setters.Add((New-Object System.Windows.Setter(
        [System.Windows.Controls.Control]::TemplateProperty,
        $tpl
    ))) | Out-Null

    return $style
}

function Show-SshKeyDetailsDialog {
    param(
        [Parameter(Mandatory=$true)][string]$PrivatePath,
        [Parameter(Mandatory=$true)][string]$PublicPath
    )

    try {
        Add-Content -Path $global:ServerToolkitLogPath -Value (
            (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
            " [INFO] Show-SshKeyDetailsDialog called priv='$PrivatePath' pub='$PublicPath'`r`n"
        ) -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}

    if (-not (Test-Path -LiteralPath $PrivatePath)) { return }
    if (-not (Test-Path -LiteralPath $PublicPath))  { return }

    $pubText = ""
    try { $pubText = (Get-Content -LiteralPath $PublicPath -Raw).Trim() } catch {}

    $bgDark     = "#1e1f22"
    $panelBg    = "#25272b"
    $fgMain     = "#f0f0f0"
    $fgMuted    = "#b0b0b0"
    $fgAccent   = "#00a8ff"
    $fgOk       = "#47d16c"
    $borderDark = "#3b3f46"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"
    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    function New-ThemedButtonStyle {
        param(
            [string]$Normal,
            [string]$Hover,
            [string]$Pressed,
            [string]$ForegroundHex = "#FFFFFF"
        )

        return (New-ThemedButtonStyleSafe `
            -Normal $Normal `
            -Hover $Hover `
            -Pressed $Pressed `
            -ForegroundHex $ForegroundHex)
    }

    $PrimaryBtnStyle   = New-ThemedButtonStyle $btnPrimary   $btnPrimaryH   $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyle $btnSecondary $btnSecondaryH $btnSecondaryP

    function New-CopyRow {
        param(
            [Parameter(Mandatory=$true)][System.Windows.Controls.StackPanel]$Parent,
            [Parameter(Mandatory=$true)][string]$Label,
            [Parameter(Mandatory=$true)][string]$Value,
            [Parameter(Mandatory=$true)][string]$Tooltip
        )

        $safeValue = if ([string]::IsNullOrWhiteSpace($Value)) { "(not available)" } else { $Value }

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $Label
        $lbl.Foreground = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent))
        $lbl.FontWeight = "SemiBold"
        $lbl.Margin = "0,0,0,4"
        $Parent.Children.Add($lbl) | Out-Null

        $row = New-Object System.Windows.Controls.Border
        $row.Tag = $safeValue
        $row.ToolTip = $Tooltip
        $row.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#00000000")
        $row.CornerRadius = "4"
        $row.Padding = "2,2"
        $row.Margin = "0,0,0,12"
        $row.Cursor = [System.Windows.Input.Cursors]::Hand

        $inner = New-Object System.Windows.Controls.StackPanel
        $inner.Orientation = "Horizontal"

        $txt = New-Object System.Windows.Controls.TextBlock
        $txt.Text = $safeValue
        $txt.FontFamily = "Consolas"
        $txt.TextWrapping = "Wrap"
        $txt.MaxWidth = 560
        $txt.Foreground = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted))
        $txt.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        $ico = New-Object System.Windows.Controls.TextBlock
        $ico.FontFamily = "Segoe MDL2 Assets"
        $ico.Text = [char]0xE8C8
        $ico.FontSize = 16
        $ico.Margin = "10,0,0,0"
        $ico.Foreground = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent))
        $ico.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        $inner.Children.Add($txt) | Out-Null
        $inner.Children.Add($ico) | Out-Null
        $row.Child = $inner

        try { $row.Resources["__CopyIcon"] = $ico } catch { try { $row.Resources.Add("__CopyIcon",$ico) } catch {} }

        $Parent.Children.Add($row) | Out-Null

        $hoverBg  = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#22333842")
        $normalBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#00000000")

        $row.Add_MouseEnter({ param($s,$e) try { $s.Background = $hoverBg } catch {} })
        $row.Add_MouseLeave({ param($s,$e) try { $s.Background = $normalBg } catch {} })

        $row.Add_MouseLeftButtonUp({
            param($s,$e)
            try { if ($e) { $e.Handled = $true } } catch {}

            try {
                $val = ""
                try { $val = [string]$s.Tag } catch { $val = "" }
                if ([string]::IsNullOrWhiteSpace($val)) { return }

                try {
                    [System.Windows.Clipboard]::SetText($val)
                } catch {
                    try { [System.Windows.Clipboard]::SetDataObject($val, $true) } catch {}
                }

                $icon2 = $null
                try { $icon2 = $s.Resources["__CopyIcon"] } catch { $icon2 = $null }

                if ($icon2 -and ($icon2 -is [System.Windows.Controls.TextBlock])) {

                    $icon2.Text = [char]0xE73E
                    $icon2.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgOk)

                    $timer = New-Object System.Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]::FromMilliseconds(650)
                    $timer.Tag = $icon2

                    $timer.Add_Tick({
                        param($sender,$evt)
                        try {
                            $sender.Tag.Text = [char]0xE8C8
                            $sender.Tag.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
                        } catch {}
                        try { $sender.Stop() } catch {}
                    })

                    $timer.Start()
                }

            } catch {}
        })
    }

    $w = New-Object System.Windows.Window
    $w.Title = "SSH Key Pair Details"
    $w.Width = 760
    $w.Height = 520
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"

    $w.Owner = $MainWindow

    try { $w.Topmost = $true } catch {}
    try { $w.Activate() | Out-Null } catch {}
    try { $w.Topmost = $false } catch {}

    $w.Background = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark))
    $w.Foreground = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain))
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*" }))    | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "SSH Key Pair Created Successfully"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Margin = "0,0,0,12"
    $root.Children.Add($hdr) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdr,0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg))
    $border.BorderBrush = ([System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark))
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $stack = New-Object System.Windows.Controls.StackPanel
    New-CopyRow $stack "Private Key Path:" $PrivatePath "Click to copy private key path"
    New-CopyRow $stack "Public Key Path:"  $PublicPath  "Click to copy public key path"
    New-CopyRow $stack "Public Key:"       $pubText     "Click to copy public key (paste into Hetzner / VPS)"

    $border.Child = $stack
    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border,1)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnCopyAll = New-Object System.Windows.Controls.Button
    $btnCopyAll.Content = "Copy Details"
    $btnCopyAll.Style = $SecondaryBtnStyle
    $btnCopyAll.Margin = "0,0,10,0"
    $btnCopyAll.Add_Click({
        try {
            $txt = @(
                "Private Key: $PrivatePath"
                "Public Key:  $PublicPath"
                ""
                $pubText
            ) -join "`r`n"
            [System.Windows.Clipboard]::SetText($txt)
        } catch {}
    })

    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "OK"
    $btnOk.Style = $PrimaryBtnStyle
    $btnOk.Add_Click({ $w.Close() })

    $btnRow.Children.Add($btnCopyAll) | Out-Null
    $btnRow.Children.Add($btnOk) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow,2)

    $w.Content = $root
    $null = $w.ShowDialog()
}

function Show-CreateSshKeyDialog {

    $sshDir = Join-Path $env:USERPROFILE ".ssh"
    if (-not (Test-Path $sshDir)) {
        New-Item -ItemType Directory -Path $sshDir | Out-Null
    }

    $bc = [System.Windows.Media.BrushConverter]::new()

    $w = New-Object System.Windows.Window
    $w.Title = "Create SSH Key Pair"
    $w.Width = 740
    $w.Height = 360
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = $bc.ConvertFromString("#1e1f22")
    $w.Foreground = $bc.ConvertFromString("#f0f0f0")
    $w.FontFamily = "Segoe UI"

    try {
        if ($MainWindow -and $MainWindow.Resources) {
            if (-not $w.Resources) { $w.Resources = New-Object System.Windows.ResourceDictionary }
            try { $w.Resources.MergedDictionaries.Add($MainWindow.Resources) | Out-Null } catch {}
        }
    } catch {}

    function _TryStyle([string]$key) {
        try { return $MainWindow.FindResource($key) } catch {}
        try { return $w.FindResource($key) } catch {}
        return $null
    }

    $SecondaryBtnStyle = _TryStyle "SecondaryButtonStyle"
    $StartBtnStyle     = _TryStyle "StartButtonStyle"

    # Root grid with rows that MATCH our SetRow() calls
    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"

    # Row 0: Header
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 1: Name label
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 2: Name box
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 3: Pass label
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 4: Pass row
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 5: Path label
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 6: Path row
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    # Row 7: Buttons row
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null

    # Header
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "Create SSH Key Pair"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Margin = "0,0,0,14"
    $root.Children.Add($hdr) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdr, 0)

    # Name label
    $lblName = New-Object System.Windows.Controls.TextBlock
    $lblName.Text = "SSH Key Name:"
    $lblName.FontWeight = "SemiBold"
    $lblName.Margin = "0,0,0,6"
    $root.Children.Add($lblName) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($lblName, 1)

    # Name box
    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.ToolTip  = "Base name for the key files (example: myserver)"
    $nameBox.MinWidth = 520
    $nameBox.Height   = 28
    $nameBox.Padding  = "6,3"
    $nameBox.Margin   = "0,0,0,14"
    $root.Children.Add($nameBox) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($nameBox, 2)

    # Pass label
    $lblPass = New-Object System.Windows.Controls.TextBlock
    $lblPass.Text = "Passphrase:"
    $lblPass.Margin = "0,0,0,6"
    $root.Children.Add($lblPass) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($lblPass, 3)

    # Pass row: PasswordBox + Reveal TextBox + Toggle
    $passGrid = New-Object System.Windows.Controls.Grid
    $passGrid.Margin = "0,0,0,14"
    $passGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="*"   })) | Out-Null
    $passGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto"})) | Out-Null

    $pwBox = New-Object System.Windows.Controls.PasswordBox
    $pwBox.MinWidth = 520
    $pwBox.Height   = 28
    $pwBox.Padding  = "6,3"
    $pwBox.Margin   = "0,0,10,0"
    $passGrid.Children.Add($pwBox) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($pwBox, 0)

    $pwReveal = New-Object System.Windows.Controls.TextBox
    $pwReveal.MinWidth  = 520
    $pwReveal.Height    = 28
    $pwReveal.Padding   = "6,3"
    $pwReveal.Margin    = "0,0,10,0"
    $pwReveal.Visibility = "Collapsed"
    $passGrid.Children.Add($pwReveal) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($pwReveal, 0)

    $toggle = New-Object System.Windows.Controls.Primitives.ToggleButton
    $toggle.Width = 34
    $toggle.Height = 28
    $toggle.VerticalAlignment = "Center"
    $toggle.ToolTip = "Show / hide passphrase"

    try {
        $toggle.Background  = $bc.ConvertFromString("#2f3238")
        $toggle.BorderBrush = $bc.ConvertFromString("#2f3238")
        $toggle.BorderThickness = "0"
        $toggle.Padding = "4"
    } catch {}

    $eyeIcon = New-Object System.Windows.Controls.TextBlock
    $eyeIcon.FontFamily = "Segoe MDL2 Assets"
    $eyeIcon.Text = [char]0xE890   # eye
    $eyeIcon.FontSize = 14
    $eyeIcon.VerticalAlignment = "Center"
    $eyeIcon.HorizontalAlignment = "Center"
    $eyeIcon.Foreground = $bc.ConvertFromString("#f0f0f0")

    $toggle.Content = $eyeIcon

    $toggle.Add_Checked({
        try { $pwReveal.Text = $pwBox.Password } catch {}
        try { $pwBox.Visibility = "Collapsed" } catch {}
        try { $pwReveal.Visibility = "Visible" } catch {}
        try { $eyeIcon.Text = [char]0xE72E } catch {}  # eye-off
    })

    $toggle.Add_Unchecked({
        try { $pwBox.Password = $pwReveal.Text } catch {}
        try { $pwReveal.Visibility = "Collapsed" } catch {}
        try { $pwBox.Visibility = "Visible" } catch {}
        try { $eyeIcon.Text = [char]0xE890 } catch {}  # eye
    })

    $passGrid.Children.Add($toggle) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($toggle, 1)

    $root.Children.Add($passGrid) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($passGrid, 4)

    # Path label
    $lblPath = New-Object System.Windows.Controls.TextBlock
    $lblPath.Text = "Save Location:"
    $lblPath.Margin = "0,0,0,6"
    $root.Children.Add($lblPath) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($lblPath, 5)

    # Path row
    $pathRow = New-Object System.Windows.Controls.Grid
    $pathRow.Margin = "0,0,0,0"
    $pathRow.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="*"   })) | Out-Null
    $pathRow.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto"})) | Out-Null

    $pathBox = New-Object System.Windows.Controls.TextBox
    $pathBox.Text = $sshDir
    $pathBox.IsReadOnly = $true
    $pathBox.MinWidth = 520
    $pathBox.Height   = 28
    $pathBox.Padding  = "6,3"
    $pathBox.Margin   = "0,0,10,0"
    $pathRow.Children.Add($pathBox) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($pathBox, 0)

    $browseBtn = New-Object System.Windows.Controls.Button
    $browseBtn.Content = "Browse..."
    $browseBtn.Padding = "10,4"
    if ($SecondaryBtnStyle) { $browseBtn.Style = $SecondaryBtnStyle }

    $browseBtn.Add_Click({
        try {
            $f = New-Object System.Windows.Forms.FolderBrowserDialog
            $f.SelectedPath = $pathBox.Text
            if ($f.ShowDialog() -eq "OK") {
                $pathBox.Text = $f.SelectedPath
            }
        } catch {}
    })

    $pathRow.Children.Add($browseBtn) | Out-Null
    [System.Windows.Controls.Grid]::SetColumn($browseBtn, 1)

    $root.Children.Add($pathRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($pathRow, 6)

    # Buttons row
    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.VerticalAlignment = "Bottom"
    $btnRow.Margin = "0,16,0,0"

    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "Cancel"
    $cancel.Margin = "0,0,10,0"
    if ($SecondaryBtnStyle) { $cancel.Style = $SecondaryBtnStyle }
    $cancel.Add_Click({ try { $w.Close() } catch {} })

    $create = New-Object System.Windows.Controls.Button
    $create.Content = "Create SSH"
    if ($StartBtnStyle) { $create.Style = $StartBtnStyle }

    $create.Add_Click({
        try {
            $name = ""
            try { $name = $nameBox.Text.Trim() } catch { $name = "" }

            if (-not $name) {
                [System.Windows.MessageBox]::Show($w, "Key name is required.", "Missing Name",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            $pp = ""
            try {
                if ($pwBox.Visibility -eq "Visible") { $pp = $pwBox.Password } else { $pp = $pwReveal.Text }
            } catch { $pp = "" }

            if ([string]::IsNullOrWhiteSpace($pp)) {
                [System.Windows.MessageBox]::Show($w, "Passphrase is required. Please enter a passphrase.", "Passphrase Required",
                    [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            $base = Join-Path $pathBox.Text "$name-ssh"
            $priv = $base
            $pub  = "$base.pub"

            if ((Test-Path $priv) -or (Test-Path $pub)) {

                $res = [System.Windows.MessageBox]::Show(
                    $w,
                    "One or more files already exist:`r`n$priv`r`n$pub`r`n`r`nOverwrite?",
                    "Overwrite SSH Key?",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning
                )

                if ($res -ne [System.Windows.MessageBoxResult]::Yes) {
                    return
                }

                try { Remove-Item -LiteralPath $priv -Force -ErrorAction SilentlyContinue } catch {}
                try { Remove-Item -LiteralPath $pub  -Force -ErrorAction SilentlyContinue } catch {}
            }


            $sshKeygen = "ssh-keygen"
            $comment = $name
            & $sshKeygen -t ed25519 -f "$priv" -N "$pp" -C "$comment" -q | Out-Null


            try { $KeyBox.Text = $priv } catch {}
            try { $PasswordBox.Password = $pp } catch {}
            try { $w.Close() } catch {}

            try {
                $priv2 = [string]$priv
                $pub2  = [string]$pub

                try {
                    Add-Content -Path $global:ServerToolkitLogPath -Value (
                        (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                        " [INFO] Scheduling SSH details dialog priv='$priv2' pub='$pub2'`r`n"
                    ) -Encoding UTF8 -ErrorAction SilentlyContinue
                } catch {}

                $actionShow = [Action[object,object]]{
                    param($pPriv, $pPub)
                    try {
                        $p1 = [string]$pPriv
                        $p2 = [string]$pPub

                        if ([string]::IsNullOrWhiteSpace($p1) -or [string]::IsNullOrWhiteSpace($p2)) {
                            throw "Details dialog received empty path(s). priv='$p1' pub='$p2'"
                        }

                        Show-SshKeyDetailsDialog -PrivatePath $p1 -PublicPath $p2
                    } catch {
                        try {
                            Add-Content -Path $global:ServerToolkitLogPath -Value (
                                (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                                " [UIERR] Show-SshKeyDetailsDialog failed: " + $_.Exception.ToString() + "`r`n"
                            ) -Encoding UTF8 -ErrorAction SilentlyContinue
                        } catch {}
                        try {
                            [System.Windows.MessageBox]::Show(
                                $MainWindow,
                                ("SSH Key was created, but details window failed to open.`r`n`r`n" + $_.Exception.Message),
                                "SSH Key Created",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Warning
                            ) | Out-Null
                        } catch {}
                    }
                }

                $null = $script:UiDispatcher.BeginInvoke($actionShow, [object[]]@($priv2, $pub2))
            } catch {}
        } catch {
            try {
                Add-Content -Path $global:ServerToolkitLogPath -Value (
                    (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                    " [UIERR] Create SSH Key dialog inner create failed: " +
                    $_.Exception.ToString() + "`r`n"
                ) -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
            try {
                [System.Windows.MessageBox]::Show($w, ("Failed to create key:`r`n`r`n" + $_.Exception.Message),
                    "Create SSH Key", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            } catch {}
        }
    })

    $btnRow.Children.Add($cancel) | Out-Null
    $btnRow.Children.Add($create) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 7)

    $w.Content = $root

    try { $w.Add_ContentRendered({ try { $nameBox.Focus() } catch {} }) | Out-Null } catch {}

    $null = $w.ShowDialog()
}

function Show-P12PasswordDialog {
    param(
        [Parameter(Mandatory=$true)][string]$P12Path
    )

    $bgDark        = "#1e1f22"
    $panelBg       = "#25272b"
    $fgMain        = "#f0f0f0"
    $fgMuted       = "#b0b0b0"
    $fgAccent      = "#00a8ff"
    $fgWarning     = "#f5c542"
    $borderDark    = "#3b3f46"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"

    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    $PrimaryBtnStyle   = New-ThemedButtonStyleSafe -Normal $btnPrimary   -Hover $btnPrimaryH   -Pressed $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyleSafe -Normal $btnSecondary -Hover $btnSecondaryH -Pressed $btnSecondaryP

    $w = New-Object System.Windows.Window
    $w.Title = "P12 Password (Local Only)"
    $w.Width = 720
    $w.Height = 420
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark)
    $w.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdrRow = New-Object System.Windows.Controls.StackPanel
    $hdrRow.Orientation = "Horizontal"
    $hdrRow.Margin = "0,0,0,12"

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.FontFamily = "Segoe MDL2 Assets"
    $icon.Text = [char]0xE72E
    $icon.FontSize = 22
    $icon.Margin = "0,1,10,0"
    $icon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "Enter your P12 password (local only)"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdrRow.Children.Add($icon) | Out-Null
    $hdrRow.Children.Add($hdr)  | Out-Null
    $root.Children.Add($hdrRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdrRow, 0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark)
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $body = New-Object System.Windows.Controls.StackPanel

    $t1 = New-Object System.Windows.Controls.TextBlock
    $t1.TextWrapping = "Wrap"
    $t1.Margin = "0,0,0,10"
    $t1.Text = "This password is used only on your PC to read the alias (friendlyName). It is NOT uploaded or sent to the server."
    $body.Children.Add($t1) | Out-Null

    $lblSel = New-Object System.Windows.Controls.TextBlock
    $lblSel.Text = "Selected P12 file:"
    $lblSel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lblSel.FontWeight = "SemiBold"
    $lblSel.Margin = "0,0,0,4"
    $body.Children.Add($lblSel) | Out-Null

    $sel = New-Object System.Windows.Controls.TextBlock
    $sel.TextWrapping = "Wrap"
    $sel.Margin = "0,0,0,12"
    $sel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $sel.FontFamily = "Consolas"
    $sel.Text = $P12Path
    $body.Children.Add($sel) | Out-Null

    $lblPw = New-Object System.Windows.Controls.TextBlock
    $lblPw.Text = "P12 Password:"
    $lblPw.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lblPw.FontWeight = "SemiBold"
    $lblPw.Margin = "0,0,0,6"
    $body.Children.Add($lblPw) | Out-Null

    $pwBox = New-Object System.Windows.Controls.PasswordBox
    $pwBox.Width = 420
    $pwBox.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2f3238")
    $pwBox.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $pwBox.Padding = "6,4"
    $body.Children.Add($pwBox) | Out-Null

    $border.Child = $body
    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border, 1)

    $result = [hashtable]::Synchronized(@{ Ok=$false; Value="" })

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Cancel"
    $btnCancel.Style = $SecondaryBtnStyle
    $btnCancel.Margin = "0,0,10,0"
    $btnCancel.Add_Click({ $result.Ok=$false; $result.Value=""; $w.Close() })
    $btnRow.Children.Add($btnCancel) | Out-Null

    $btnContinue = New-Object System.Windows.Controls.Button
    $btnContinue.Content = "Continue"
    $btnContinue.Style = $PrimaryBtnStyle
    $btnContinue.Add_Click({
        $val = $pwBox.Password
        if ([string]::IsNullOrWhiteSpace($val)) {
            [System.Windows.MessageBox]::Show($w, "Password cannot be empty.", "P12 Password Required",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }
        $result.Ok=$true
        $result.Value=$val
        $w.Close()
    })
    $btnRow.Children.Add($btnContinue) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 2)

    $w.Content = $root
    try { $w.Add_ContentRendered({ try { $pwBox.Focus() } catch {} }) | Out-Null } catch {}
    $null = $w.ShowDialog()

    if ($result.Ok -eq $true) { return $result.Value }
    return ""
}

function Show-PpkHelpDialog {
    param(
        [Parameter(Mandatory=$true)][string]$PpkPath
    )

    $url = "https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html"
    $sshDir = Join-Path $HOME ".ssh"

    $baseName = [IO.Path]::GetFileNameWithoutExtension($PpkPath)
    $recommendedKey = Join-Path $sshDir ($baseName + "-ssh")

    $bgDark        = "#1e1f22"
    $panelBg       = "#25272b"
    $fgMain        = "#f0f0f0"
    $fgMuted       = "#b0b0b0"
    $fgAccent      = "#00a8ff"
    $fgWarning     = "#f5c542"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"

    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    $borderDark    = "#3b3f46"

    $PrimaryBtnStyle   = New-ThemedButtonStyleSafe -Normal $btnPrimary   -Hover $btnPrimaryH   -Pressed $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyleSafe -Normal $btnSecondary -Hover $btnSecondaryH -Pressed $btnSecondaryP

    $w = New-Object System.Windows.Window
    $w.Title = "PuTTY Key Not Supported"
    $w.Width = 720
    $w.Height = 600
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark)
    $w.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdrRow = New-Object System.Windows.Controls.StackPanel
    $hdrRow.Orientation = "Horizontal"
    $hdrRow.Margin = "0,0,0,12"

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.FontFamily = "Segoe MDL2 Assets"
    $icon.Text = [char]0xE72E
    $icon.FontSize = 22
    $icon.Margin = "0,1,10,0"
    $icon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "PuTTY Private Key Detected (.ppk)"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdrRow.Children.Add($icon) | Out-Null
    $hdrRow.Children.Add($hdr)  | Out-Null
    $root.Children.Add($hdrRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdrRow, 0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark)
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $scroll = New-Object System.Windows.Controls.ScrollViewer
    $scroll.VerticalScrollBarVisibility = "Auto"
    $scroll.HorizontalScrollBarVisibility = "Disabled"

    $body = New-Object System.Windows.Controls.StackPanel

    $t1 = New-Object System.Windows.Controls.TextBlock
    $t1.TextWrapping = "Wrap"
    $t1.Margin = "0,0,0,10"
    $t1.Text = "This toolkit requires a standard OpenSSH private key. PuTTY (.ppk) keys cannot be used directly."
    $body.Children.Add($t1) | Out-Null

    $lblSel = New-Object System.Windows.Controls.TextBlock
    $lblSel.Text = "Selected file:"
    $lblSel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lblSel.FontWeight = "SemiBold"
    $lblSel.Margin = "0,0,0,4"
    $body.Children.Add($lblSel) | Out-Null

    $sel = New-Object System.Windows.Controls.TextBlock
    $sel.TextWrapping = "Wrap"
    $sel.Margin = "0,0,0,12"
    $sel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $sel.FontFamily = "Consolas"
    $sel.Text = $PpkPath
    $body.Children.Add($sel) | Out-Null

    $tStepsHdr = New-Object System.Windows.Controls.TextBlock
    $tStepsHdr.Text = "How to convert your key:"
    $tStepsHdr.FontWeight = "SemiBold"
    $tStepsHdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $tStepsHdr.Margin = "0,0,0,6"
    $body.Children.Add($tStepsHdr) | Out-Null

    $steps = New-Object System.Windows.Controls.TextBlock
    $steps.TextWrapping = "Wrap"
    $steps.Margin = "0,0,0,10"
    $steps.Text = @"
1) Open PuTTYgen

2) Conversions -> Import Key
   Select your .ppk file

3) If prompted, enter your PPK passphrase

4) Conversions -> Export OpenSSH key (force new file format)
"@
    $body.Children.Add($steps) | Out-Null

    $lblRec = New-Object System.Windows.Controls.TextBlock
    $lblRec.Text = "Recommended filename:"
    $lblRec.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lblRec.FontWeight = "SemiBold"
    $lblRec.Margin = "0,0,0,4"
    $body.Children.Add($lblRec) | Out-Null

    $rec = New-Object System.Windows.Controls.TextBlock
    $rec.TextWrapping = "Wrap"
    $rec.Margin = "0,0,0,14"
    $rec.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $rec.FontFamily = "Consolas"
    $rec.Text = $recommendedKey
    $body.Children.Add($rec) | Out-Null

    $linkLine = New-Object System.Windows.Controls.TextBlock
    $linkLine.Margin = "0,0,0,0"
    $linkLine.Inlines.Add((New-Object System.Windows.Documents.Run("Download PuTTY: "))) | Out-Null

    $hl = New-Object System.Windows.Documents.Hyperlink
    $hl.Inlines.Add($url) | Out-Null
    $hl.NavigateUri = [Uri]$url
    $hl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $hl.Add_Click({ try { Start-Process $url } catch {} })
    $linkLine.Inlines.Add($hl) | Out-Null
    $body.Children.Add($linkLine) | Out-Null

    $scroll.Content = $body
    $border.Child = $scroll

    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border, 1)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnOpen = New-Object System.Windows.Controls.Button
    $btnOpen.Content = "Open PuTTY Download"
    $btnOpen.Style = $PrimaryBtnStyle
    $btnOpen.Margin = "0,0,10,0"
    $btnOpen.Add_Click({ try { Start-Process $url } catch {} })
    $btnRow.Children.Add($btnOpen) | Out-Null

    $btnCopyLink = New-Object System.Windows.Controls.Button
    $btnCopyLink.Content = "Copy Download Link"
    $btnCopyLink.Style = $SecondaryBtnStyle
    $btnCopyLink.Margin = "0,0,10,0"
    $btnCopyLink.Add_Click({ try { [System.Windows.Clipboard]::SetText($url) } catch {} })
    $btnRow.Children.Add($btnCopyLink) | Out-Null

    $btnCopyRec = New-Object System.Windows.Controls.Button
    $btnCopyRec.Content = "Copy Recommended Filename"
    $btnCopyRec.Style = $SecondaryBtnStyle
    $btnCopyRec.Margin = "0,0,10,0"
    $btnCopyRec.Add_Click({ try { [System.Windows.Clipboard]::SetText($recommendedKey) } catch {} })
    $btnRow.Children.Add($btnCopyRec) | Out-Null

    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "OK"
    $btnOk.Style = $SecondaryBtnStyle
    $btnOk.Add_Click({ $w.Close() })
    $btnRow.Children.Add($btnOk) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 2)

    $w.Content = $root
    $null = $w.ShowDialog()
}

function Show-P12ValidatedDialog {
    param(
        [Parameter(Mandatory=$true)][string]$LocalPath,
        [Parameter(Mandatory=$false)][string]$Alias = ""
    )

    $bgDark        = "#1e1f22"
    $panelBg       = "#25272b"
    $fgMain        = "#f0f0f0"
    $fgMuted       = "#b0b0b0"
    $fgAccent      = "#00a8ff"
    $fgOk          = "#47d16c"
    $borderDark    = "#3b3f46"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"

    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    $PrimaryBtnStyle   = New-ThemedButtonStyleSafe -Normal $btnPrimary   -Hover $btnPrimaryH   -Pressed $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyleSafe -Normal $btnSecondary -Hover $btnSecondaryH -Pressed $btnSecondaryP

    $shownAlias = $Alias
    if ([string]::IsNullOrWhiteSpace($shownAlias)) { $shownAlias = "(none)" }

    $shownLocal = $LocalPath
    if ([string]::IsNullOrWhiteSpace($shownLocal)) { $shownLocal = "(unknown)" }

    $w = New-Object System.Windows.Window
    $w.Title = "P12 Validated (Local Only)"
    $w.Width = 720
    $w.Height = 440
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark)
    $w.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdrRow = New-Object System.Windows.Controls.StackPanel
    $hdrRow.Orientation = "Horizontal"
    $hdrRow.Margin = "0,0,0,12"

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.FontFamily = "Segoe MDL2 Assets"
    $icon.Text = [char]0xE73E
    $icon.FontSize = 22
    $icon.Margin = "0,1,10,0"
    $icon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgOk)

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "P12 validated locally (upload will proceed)"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgOk)

    $hdrRow.Children.Add($icon) | Out-Null
    $hdrRow.Children.Add($hdr)  | Out-Null
    $root.Children.Add($hdrRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdrRow, 0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark)
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $body = New-Object System.Windows.Controls.StackPanel

    $t1 = New-Object System.Windows.Controls.TextBlock
    $t1.TextWrapping = "Wrap"
    $t1.Margin = "0,0,0,10"
    $t1.Text = "Your password successfully unlocked the P12 on this PC. This does NOT upload your password; it is used locally only."
    $body.Children.Add($t1) | Out-Null

    $lblA = New-Object System.Windows.Controls.TextBlock
    $lblA.Text = "Alias (friendlyName):"
    $lblA.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lblA.FontWeight = "SemiBold"
    $lblA.Margin = "0,0,0,4"
    $body.Children.Add($lblA) | Out-Null

    $tbA = New-Object System.Windows.Controls.TextBlock
    $tbA.TextWrapping = "Wrap"
    $tbA.Margin = "0,0,0,12"
    $tbA.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $tbA.FontFamily = "Consolas"
    $tbA.Text = $shownAlias
    $body.Children.Add($tbA) | Out-Null

    $lblL = New-Object System.Windows.Controls.TextBlock
    $lblL.Text = "Local file path:"
    $lblL.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lblL.FontWeight = "SemiBold"
    $lblL.Margin = "0,0,0,4"
    $body.Children.Add($lblL) | Out-Null

    $tbL = New-Object System.Windows.Controls.TextBlock
    $tbL.TextWrapping = "Wrap"
    $tbL.Margin = "0,0,0,0"
    $tbL.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $tbL.FontFamily = "Consolas"
    $tbL.Text = $shownLocal
    $body.Children.Add($tbL) | Out-Null

    $border.Child = $body
    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border, 1)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnCopyAlias = New-Object System.Windows.Controls.Button
    $btnCopyAlias.Content = "Copy Alias"
    $btnCopyAlias.Style = $SecondaryBtnStyle
    $btnCopyAlias.Margin = "0,0,10,0"
    $btnCopyAlias.Add_Click({ try { [System.Windows.Clipboard]::SetText($shownAlias) } catch {} })
    $btnRow.Children.Add($btnCopyAlias) | Out-Null

    $btnCopyLocal = New-Object System.Windows.Controls.Button
    $btnCopyLocal.Content = "Copy Local Path"
    $btnCopyLocal.Style = $SecondaryBtnStyle
    $btnCopyLocal.Margin = "0,0,10,0"
    $btnCopyLocal.Add_Click({ try { [System.Windows.Clipboard]::SetText($shownLocal) } catch {} })
    $btnRow.Children.Add($btnCopyLocal) | Out-Null

    $btnCopyAll = New-Object System.Windows.Controls.Button
    $btnCopyAll.Content = "Copy All"
    $btnCopyAll.Style = $PrimaryBtnStyle
    $btnCopyAll.Margin = "0,0,10,0"
    $btnCopyAll.Add_Click({
        try {
            [System.Windows.Clipboard]::SetText(("P12 Alias: $shownAlias`r`nLocal Path: $shownLocal"))
        } catch {}
    })
    $btnRow.Children.Add($btnCopyAll) | Out-Null

    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "OK"
    $btnOk.Style = $SecondaryBtnStyle
    $btnOk.Add_Click({ $w.Close() })
    $btnRow.Children.Add($btnOk) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 2)

    $w.Content = $root
    $null = $w.ShowDialog()
}
function Show-Confirm3ChoiceDialog {
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string]$Header,
        [Parameter(Mandatory=$true)][string]$Body,
        [Parameter(Mandatory=$true)][string]$YesText,
        [Parameter(Mandatory=$true)][string]$NoText,
        [Parameter(Mandatory=$true)][string]$CancelText
    )

    $bgDark     = "#1e1f22"
    $panelBg    = "#25272b"
    $fgMain     = "#f0f0f0"
    $fgMuted    = "#b0b0b0"
    $fgAccent   = "#00a8ff"
    $fgWarning  = "#f5c542"
    $borderDark = "#3b3f46"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"
    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    $PrimaryBtnStyle   = New-ThemedButtonStyleSafe -Normal $btnPrimary   -Hover $btnPrimaryH   -Pressed $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyleSafe -Normal $btnSecondary -Hover $btnSecondaryH -Pressed $btnSecondaryP

    $result = [hashtable]::Synchronized(@{ Choice = "cancel" })

    $w = New-Object System.Windows.Window
    $w.Title = $Title
    $w.Width = 760
    $w.Height = 460
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark)
    $w.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdrRow = New-Object System.Windows.Controls.StackPanel
    $hdrRow.Orientation = "Horizontal"
    $hdrRow.Margin = "0,0,0,12"

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.FontFamily = "Segoe MDL2 Assets"
    $icon.Text = [char]0xE7BA
    $icon.FontSize = 22
    $icon.Margin = "0,1,10,0"
    $icon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = $Header
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdrRow.Children.Add($icon) | Out-Null
    $hdrRow.Children.Add($hdr)  | Out-Null
    $root.Children.Add($hdrRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdrRow, 0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark)
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $bodyTb = New-Object System.Windows.Controls.TextBlock
    $bodyTb.TextWrapping = "Wrap"
    $bodyTb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $bodyTb.FontFamily = "Consolas"
    $bodyTb.Text = $Body

    $border.Child = $bodyTb
    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border, 1)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnYes = New-Object System.Windows.Controls.Button
    $btnYes.Content = $YesText
    $btnYes.Style = $PrimaryBtnStyle
    $btnYes.Margin = "0,0,10,0"
    $btnYes.Add_Click({ $result.Choice = "yes"; $w.Close() })
    $btnRow.Children.Add($btnYes) | Out-Null

    $btnNo = New-Object System.Windows.Controls.Button
    $btnNo.Content = $NoText
    $btnNo.Style = $SecondaryBtnStyle
    $btnNo.Margin = "0,0,10,0"
    $btnNo.Add_Click({ $result.Choice = "no"; $w.Close() })
    $btnRow.Children.Add($btnNo) | Out-Null

    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = $CancelText
    $btnCancel.Style = $SecondaryBtnStyle
    $btnCancel.Add_Click({ $result.Choice = "cancel"; $w.Close() })
    $btnRow.Children.Add($btnCancel) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 2)

    $w.Content = $root
    $null = $w.ShowDialog()

    return $result.Choice
}

function Show-P12OverwriteDialog {
    param(
        [Parameter(Mandatory=$true)][string]$RemotePath,
        [Parameter(Mandatory=$true)][string]$LocalPath
    )

    $rp = if ([string]::IsNullOrWhiteSpace($RemotePath)) { "(unknown)" } else { $RemotePath }
    $lp = if ([string]::IsNullOrWhiteSpace($LocalPath))  { "(unknown)" } else { $LocalPath }

    $body = @()
    $body += "A P12 file already exists on the server:"
    $body += ""
    $body += "  $rp"
    $body += ""
    $body += "Local file selected:"
    $body += ""
    $body += "  $lp"
    $body += ""
    $body += "Choose what to do:"
    $bodyText = ($body -join "`r`n")

    $choice = Show-Confirm3ChoiceDialog `
        -Title "P12 Already Exists" `
        -Header "Remote P12 already exists" `
        -Body $bodyText `
        -YesText "Overwrite" `
        -NoText "Backup + Overwrite" `
        -CancelText "Skip"

    if ($choice -eq "yes")    { return "overwrite" }
    if ($choice -eq "no")     { return "backup" }
    return "skip"
}

function Show-SaveProfileDialog {
    param(
        [Parameter(Mandatory=$true)][string]$ServerHost,
        [Parameter(Mandatory=$true)][string]$SshPort,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$false)][string]$IdentityFile,
        [Parameter(Mandatory=$true)][int]$PromptFlag
    )

    $bgDark     = "#1e1f22"
    $panelBg    = "#25272b"
    $fgMain     = "#f0f0f0"
    $fgMuted    = "#b0b0b0"
    $fgAccent   = "#00a8ff"
    $fgWarning  = "#f5c542"
    $borderDark = "#3b3f46"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"
    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    $PrimaryBtnStyle   = New-ThemedButtonStyleSafe -Normal $btnPrimary   -Hover $btnPrimaryH   -Pressed $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyleSafe -Normal $btnSecondary -Hover $btnSecondaryH -Pressed $btnSecondaryP

    $result = [hashtable]::Synchronized(@{ Ok=$false; Name="" })

    $defaultName = $ServerHost
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $defaultName = $defaultName.Replace($c, '_') }

    $statusLine = if ($PromptFlag -eq 1) { "Connection verified & setup complete." } else { "Connection verified (no changes)." }
    $idLine = if ([string]::IsNullOrWhiteSpace($IdentityFile)) { "(none)" } else { $IdentityFile }

    $w = New-Object System.Windows.Window
    $w.Title = "Save Connection Profile"
    $w.Width = 760
    $w.Height = 520
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark)
    $w.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = $statusLine
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Margin = "0,0,0,12"
    $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $root.Children.Add($hdr) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdr, 0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark)
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $stack = New-Object System.Windows.Controls.StackPanel

    $sum = @()
    $sum += "Settings that will be saved:"
    $sum += ""
    $sum += "HostName: $ServerHost"
    $sum += "User:     $Username"
    $sum += "Port:     $SshPort"
    $sum += "Identity: $idLine"
    $sumTxt = ($sum -join "`r`n")

    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.TextWrapping = "Wrap"
    $tb.FontFamily = "Consolas"
    $tb.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $tb.Text = $sumTxt
    $tb.Margin = "0,0,0,14"
    $stack.Children.Add($tb) | Out-Null

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = "Profile name (Host alias + filename):"
    $lbl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $lbl.FontWeight = "SemiBold"
    $lbl.Margin = "0,0,0,6"
    $stack.Children.Add($lbl) | Out-Null

    $nameBox = New-Object System.Windows.Controls.TextBox
    $nameBox.Text = $defaultName
    $nameBox.Width = 520
    $nameBox.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2f3238")
    $nameBox.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $nameBox.Padding = "6,4"
    $stack.Children.Add($nameBox) | Out-Null

    $sshDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) ".ssh"

    $pathHdr = New-Object System.Windows.Controls.TextBlock
    $pathHdr.Text = "Will save to:"
    $pathHdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
    $pathHdr.FontWeight = "SemiBold"
    $pathHdr.Margin = "0,14,0,6"
    $stack.Children.Add($pathHdr) | Out-Null

    $pathBox = New-Object System.Windows.Controls.TextBlock
    $pathBox.TextWrapping = "Wrap"
    $pathBox.FontFamily = "Consolas"
    $pathBox.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
    $pathBox.Margin = "0,0,0,0"
    $stack.Children.Add($pathBox) | Out-Null

    $existsWarn = New-Object System.Windows.Controls.TextBlock
    $existsWarn.TextWrapping = "Wrap"
    $existsWarn.FontWeight = "SemiBold"
    $existsWarn.Margin = "0,10,0,6"
    $existsWarn.Visibility = "Collapsed"
    $existsWarn.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)
    $stack.Children.Add($existsWarn) | Out-Null

    $script:__SaveBtnRef = $null

    $updatePath = {
        try {
            $n = $nameBox.Text
            if ([string]::IsNullOrWhiteSpace($n)) {
                $pathBox.Text = "(enter a profile name)"
                $existsWarn.Visibility = "Collapsed"
                if ($script:__SaveBtnRef) { $script:__SaveBtnRef.IsEnabled = $false }
                return
            }

            $safe = $n
            foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $safe = $safe.Replace($c, '_') }

            $candidate = (Join-Path $sshDir ("{0}_ssh_config.txt" -f $safe))
            $pathBox.Text = $candidate

            $exists = $false
            try { $exists = (Test-Path -LiteralPath $candidate) } catch { $exists = $false }

            if ($exists) {
                $existsWarn.Text = "Warning: This profile file already exists. If you continue, you will be asked to confirm overwrite."
                $existsWarn.Visibility = "Visible"
            } else {
                $existsWarn.Visibility = "Collapsed"
            }

            if ($script:__SaveBtnRef) { $script:__SaveBtnRef.IsEnabled = $true }
        } catch {
            $pathBox.Text = "(unknown)"
            $existsWarn.Visibility = "Collapsed"
            if ($script:__SaveBtnRef) { $script:__SaveBtnRef.IsEnabled = $false }
        }
    }

    $null = $nameBox.Add_TextChanged($updatePath)
    & $updatePath

    $border.Child = $stack
    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border, 1)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnSave = New-Object System.Windows.Controls.Button
    $script:__SaveBtnRef = $btnSave
    $btnSave.Content = "Save Profile"
    $btnSave.Style = $PrimaryBtnStyle
    $btnSave.Margin = "0,0,10,0"
    $btnSave.Add_Click({
        $n = $nameBox.Text
        if ([string]::IsNullOrWhiteSpace($n)) { return }

        $targetPath = $pathBox.Text
        $exists = $false
        try { $exists = (Test-Path -LiteralPath $targetPath) } catch { $exists = $false }

        if ($exists) {
            $res = [System.Windows.MessageBox]::Show(
                $w,
                "This profile already exists:`r`n`r`n$targetPath`r`n`r`nOverwrite it?",
                "Overwrite Profile?",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            if ($res -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }

        $result.Ok = $true
        $result.Name = $n
        $w.Close()
    })
    $btnRow.Children.Add($btnSave) | Out-Null

    $btnSkip = New-Object System.Windows.Controls.Button
    $btnSkip.Content = "Skip"
    $btnSkip.Style = $SecondaryBtnStyle
    $btnSkip.Margin = "0,0,10,0"
    $btnSkip.Add_Click({ $result.Ok = $false; $result.Name=""; $w.Close() })
    $btnRow.Children.Add($btnSkip) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 2)

    $w.Content = $root
    try { $w.Add_ContentRendered({ try { $nameBox.SelectAll(); $nameBox.Focus() } catch {} }) | Out-Null } catch {}
    $null = $w.ShowDialog()

    if ($result.Ok -eq $true) { return $result.Name }
    return ""
}

function Show-P12AliasDialog {
    param(
        [Parameter(Mandatory=$true)][string]$Alias,
        [Parameter(Mandatory=$false)][string]$RemotePath = "",
        [Parameter(Mandatory=$false)][string]$LocalPath = "",
        [Parameter(Mandatory=$false)][string]$NodeId = ""
    )

    $bgDark        = "#1e1f22"
    $panelBg       = "#25272b"
    $fgMain        = "#f0f0f0"
    $fgMuted       = "#b0b0b0"
    $fgAccent      = "#00a8ff"
    $fgWarning     = "#f5c542"
    $borderDark    = "#3b3f46"

    $btnPrimary    = "#00a8ff"
    $btnPrimaryH   = "#14b5ff"
    $btnPrimaryP   = "#0090d0"

    $btnSecondary  = "#3b3f46"
    $btnSecondaryH = "#4a4f5a"
    $btnSecondaryP = "#2f3238"

    $PrimaryBtnStyle   = New-ThemedButtonStyleSafe -Normal $btnPrimary   -Hover $btnPrimaryH   -Pressed $btnPrimaryP
    $SecondaryBtnStyle = New-ThemedButtonStyleSafe -Normal $btnSecondary -Hover $btnSecondaryH -Pressed $btnSecondaryP

    function New-CopyRow {
        param(
            [Parameter(Mandatory=$true)][System.Windows.Controls.StackPanel]$Parent,
            [Parameter(Mandatory=$true)][string]$Label,
            [Parameter(Mandatory=$true)][string]$Value,
            [Parameter(Mandatory=$true)][string]$Tooltip
        )

        $safeValue = if ([string]::IsNullOrWhiteSpace($Value)) { "(not available)" } else { $Value }

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $Label
        $lbl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
        $lbl.FontWeight = "SemiBold"
        $lbl.Margin = "0,0,0,4"
        $Parent.Children.Add($lbl) | Out-Null

        $row = New-Object System.Windows.Controls.Border
        $row.Tag = $safeValue
        $row.ToolTip = $Tooltip
        $row.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#00000000")
        $row.CornerRadius = "4"
        $row.Padding = "2,2"
        $row.Margin = "0,0,0,12"
        $row.Cursor = [System.Windows.Input.Cursors]::Hand

        $inner = New-Object System.Windows.Controls.StackPanel
        $inner.Orientation = "Horizontal"

        $txt = New-Object System.Windows.Controls.TextBlock
        $txt.Text = $safeValue
        $txt.FontFamily = "Consolas"
        $txt.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMuted)
        $txt.TextWrapping = "Wrap"
        $txt.MaxWidth = 560
        $txt.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        $ico = New-Object System.Windows.Controls.TextBlock
        $ico.FontFamily = "Segoe MDL2 Assets"
        $ico.Text = [char]0xE8C8
        $ico.FontSize = 16
        $ico.Margin = "10,0,0,0"
        $ico.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
        $ico.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

        $inner.Children.Add($txt) | Out-Null
        $inner.Children.Add($ico) | Out-Null
        $row.Child = $inner
        $Parent.Children.Add($row) | Out-Null

        $hoverBg  = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#22333842")
        $normalBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#00000000")

        $row.Add_MouseEnter({ try { $row.Background = $hoverBg } catch {} })
        $row.Add_MouseLeave({ try { $row.Background = $normalBg } catch {} })

        $row.Add_MouseLeftButtonUp({
            param($s,$e)
            try { if ($e) { $e.Handled = $true } } catch {}
            try {
                $val = [string]$s.Tag
                if ([string]::IsNullOrWhiteSpace($val)) { return }

                try { [System.Windows.Clipboard]::SetText($val) } catch { try { [System.Windows.Clipboard]::SetDataObject($val, $true) } catch {} }

                $ico.Text = [char]0xE73E
                $ico.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#47d16c")

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(650)
                $timer.Add_Tick({
                    try {
                        $ico.Text = [char]0xE8C8
                        $ico.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgAccent)
                    } catch {}
                    try { $timer.Stop() } catch {}
                })
                $timer.Start()
            } catch {}
        })
    }

    $w = New-Object System.Windows.Window
    $w.Title = "P12 file details"
    $w.Width = 720
    $w.Height = 420
    $w.WindowStartupLocation = "CenterOwner"
    $w.ResizeMode = "NoResize"
    $w.Owner = $MainWindow
    $w.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($bgDark)
    $w.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgMain)
    $w.FontFamily = "Segoe UI"

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = "18"
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="*"    })) | Out-Null
    $root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height="Auto" })) | Out-Null

    $hdrRow = New-Object System.Windows.Controls.StackPanel
    $hdrRow.Orientation = "Horizontal"
    $hdrRow.Margin = "0,0,0,12"

    $icon = New-Object System.Windows.Controls.TextBlock
    $icon.FontFamily = "Segoe MDL2 Assets"
    $icon.Text = [char]0xE7BA
    $icon.FontSize = 22
    $icon.Margin = "0,1,10,0"
    $icon.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = "Document your P12 file details"
    $hdr.FontSize = 20
    $hdr.FontWeight = "SemiBold"
    $hdr.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgWarning)

    $hdrRow.Children.Add($icon) | Out-Null
    $hdrRow.Children.Add($hdr)  | Out-Null
    $root.Children.Add($hdrRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($hdrRow, 0)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($panelBg)
    $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($borderDark)
    $border.BorderThickness = "1"
    $border.CornerRadius = "6"
    $border.Padding = "14"

    $body = New-Object System.Windows.Controls.StackPanel

    $t1 = New-Object System.Windows.Controls.TextBlock
    $t1.TextWrapping = "Wrap"
    $t1.Margin = "0,0,0,10"
    $t1.Text = "This alias is commonly required later when importing or referencing the wallet. Copy it now and save it somewhere safe."
    $body.Children.Add($t1) | Out-Null

    New-CopyRow -Parent $body -Label "P12 Alias:" -Value $Alias -Tooltip "Click to copy Alias"

    $nid = $NodeId
    if ([string]::IsNullOrWhiteSpace($nid)) { $nid = "(not available)" }
    New-CopyRow -Parent $body -Label "Node ID:" -Value $nid -Tooltip "Click to copy Node ID"

    $shownRemote = if ([string]::IsNullOrWhiteSpace($RemotePath)) { "(unknown)" } else { $RemotePath }
    $shownLocal  = if ([string]::IsNullOrWhiteSpace($LocalPath))  { "(unknown)" } else { $LocalPath }

    New-CopyRow -Parent $body -Label "Uploaded to (remote path):" -Value $shownRemote -Tooltip "Click to copy Remote Path"
    New-CopyRow -Parent $body -Label "Local file path:" -Value $shownLocal -Tooltip "Click to copy Local Path"

    $border.Child = $body
    $root.Children.Add($border) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($border, 1)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = "Horizontal"
    $btnRow.HorizontalAlignment = "Right"
    $btnRow.Margin = "0,14,0,0"

    $btnCopyDetails = New-Object System.Windows.Controls.Button
    $btnCopyDetails.Content = "Copy Details"
    $btnCopyDetails.Style = $SecondaryBtnStyle
    $btnCopyDetails.Margin = "0,0,10,0"
    $btnCopyDetails.Add_Click({
        try {
            $a = if ([string]::IsNullOrWhiteSpace($Alias)) { "(none)" } else { $Alias }
            $n = if ([string]::IsNullOrWhiteSpace($NodeId)) { "(not available)" } else { $NodeId }
            $rp = if ([string]::IsNullOrWhiteSpace($RemotePath)) { "(unknown)" } else { $RemotePath }
            $lp = if ([string]::IsNullOrWhiteSpace($LocalPath)) { "(unknown)" } else { $LocalPath }
            [System.Windows.Clipboard]::SetText(("P12 Alias: $a`r`nNode ID:   $n`r`nRemote Path: $rp`r`nLocal Path:  $lp"))
        } catch {}
    })
    $btnRow.Children.Add($btnCopyDetails) | Out-Null

    $btnOk = New-Object System.Windows.Controls.Button
    $btnOk.Content = "OK"
    $btnOk.Style = $PrimaryBtnStyle
    $btnOk.Add_Click({ $w.Close() })
    $btnRow.Children.Add($btnOk) | Out-Null

    $root.Children.Add($btnRow) | Out-Null
    [System.Windows.Controls.Grid]::SetRow($btnRow, 2)

    $w.Content = $root
    $null = $w.ShowDialog()
}

$BrowseKeyButton.Add_Click({

    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Title  = "Select SSH Private Key File"
    $dlg.Filter = "All Files|*.*"

    $sshDir = $null
    try { $sshDir = Join-Path ([Environment]::GetFolderPath('UserProfile')) ".ssh" } catch {}
    if (-not $sshDir) { try { $sshDir = Join-Path $env:USERPROFILE ".ssh" } catch {} }

    try {
        if ($sshDir -and -not (Test-Path -LiteralPath $sshDir)) {
            New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
        }
        if ($sshDir -and (Test-Path -LiteralPath $sshDir)) {
            $dlg.InitialDirectory = $sshDir
        }
    } catch {}

    if (-not $dlg.ShowDialog()) { return }

    $chosen = $dlg.FileName
    if (-not (Test-Path -LiteralPath $chosen)) {
        try { $script:AppendLog.Invoke("[ERROR] Selected key file does not exist: $chosen`r`n") } catch {}
        return
    }

    $firstLine = $null
    try { $firstLine = (Get-Content -LiteralPath $chosen -TotalCount 1 -ErrorAction Stop) } catch {
        try { $script:AppendLog.Invoke("[ERROR] Failed to read key file header: $($_.Exception.Message)`r`n") } catch {}
        return
    }

    $firstLine = $firstLine.Trim()

    $isPuttyPPK = ($firstLine -like "PuTTY-User-Key-File-*")
    $isOpenSSH  = ($firstLine -eq "-----BEGIN OPENSSH PRIVATE KEY-----")
    $isPEM      = ($firstLine -match "^-----BEGIN .*PRIVATE KEY-----$")

    if ($isPuttyPPK) { try { Show-PpkHelpDialog -PpkPath $chosen } catch {}; return }

    if (-not ($isOpenSSH -or $isPEM)) {
        [System.Windows.MessageBox]::Show(
            $MainWindow,
            "The selected file does not appear to be a valid SSH private key.`r`n`r`n" +
            "Expected formats:`r`n" +
            "* OpenSSH private key`r`n" +
            "* PEM private key`r`n`r`n" +
            "First line read:`r`n$firstLine",
            "Invalid Key File",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        return
    }

    $KeyBox.Text = $chosen
    try { $script:AppendLog.Invoke("Selected SSH private key: $chosen`r`n") } catch {}
})

if ($CreateKeyButton) {
    $CreateKeyButton.Add_Click({

        function _LogCreateKey([string]$msg) {
            try {
                Add-Content -Path $global:ServerToolkitLogPath -Value (
                    (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " " + $msg + "`r`n"
                ) -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
        }

        function _FuncFingerprint([string]$name) {
            try {
                $cmd = $ExecutionContext.SessionState.InvokeCommand.GetCommand($name, "Function")
                if (-not $cmd -or -not $cmd.ScriptBlock) { return "<missing>" }

                $src = $cmd.ScriptBlock.ToString()
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($src)
                $sha = [System.Security.Cryptography.SHA256]::Create()
                try {
                    $hash = ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join ""
                } finally {
                    $sha.Dispose()
                }

                $head = $src
                if ($head.Length -gt 80) { $head = $head.Substring(0,80) + "..." }
                return ("sha256=" + $hash + " head=" + ($head -replace "\r|\n"," "))
            } catch {
                return "<fingerprint-failed: " + $_.Exception.Message + ">"
            }
        }

        try {
            _LogCreateKey ("[INFO] CreateKeyButton clicked. Show-CreateSshKeyDialog " + (_FuncFingerprint "Show-CreateSshKeyDialog"))

            _LogCreateKey ("[INFO] WPF types loaded: Grid=" + [System.Windows.Controls.Grid].FullName + " TextBlock=" + [System.Windows.Controls.TextBlock].FullName)

            Show-CreateSshKeyDialog
        }
        catch {
            $ex = $_.Exception

            _LogCreateKey ("[UIERR] Create SSH Key dialog failed (top): " + $ex.GetType().FullName + ": " + $ex.Message)

            try {
                $cur = $ex
                $depth = 0
                while ($cur -and $depth -lt 12) {
                    _LogCreateKey ("[UIERR] ex[" + $depth + "]=" + $cur.GetType().FullName + " msg=" + $cur.Message)
                    $cur = $cur.InnerException
                    $depth++
                }
            } catch {}

            try {
                _LogCreateKey ("[UIERR] Create SSH Key dialog full exception:`r`n" + $_.Exception.ToString())
            } catch {}

            try {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    ("Failed to open SSH key creation window.`r`n`r`n" + $ex.Message),
                    "Create SSH Key",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
            } catch {}
        }

    })
}

if ($BrowseP12Button -and $P12PathBox) {
    $BrowseP12Button.Add_Click({
        try {
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Title  = "Select P12 File"
            $dlg.Filter = "P12 File (*.p12)|*.p12|All Files|*.*"

            if ($dlg.ShowDialog()) {

                $chosen = $dlg.FileName
                $ext = [IO.Path]::GetExtension($chosen).ToLower()

                if ($ext -ne ".p12") {
                    try {
                        [System.Windows.MessageBox]::Show(
                            $MainWindow,
                            "Only .p12 files are supported for wallet upload.`r`n`r`nSelected:`r`n$chosen",
                            "Invalid File Type",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        ) | Out-Null
                    } catch {}
                    return
                }

                $prevP12 = ""
                try { if ($P12PathBox) { $prevP12 = ($P12PathBox.Text).Trim() } } catch { $prevP12 = "" }

                $P12PathBox.Text = $chosen

                try {
                    if ($P12PasswordBox -and -not [string]::IsNullOrWhiteSpace($prevP12) -and ($prevP12 -ne $chosen)) {
                        $P12PasswordBox.Password = ""
                        if ($script:AppendLog) { $script:AppendLog.Invoke("DEBUG: P12 file changed. Cleared P12 passphrase field.`r`n") }
                    }
                } catch {}

                try { if ($script:AppendLog) { $script:AppendLog.Invoke("Selected P12: $chosen`r`n") } } catch {}
            }
        } catch {
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] P12 browse failed: $($_.Exception.Message)`r`n") } } catch {}
        }
    })
}

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
            $script:AppendLog.Invoke("Profile loaded: $($dlg.FileName)`r`n")
        } catch {
            $script:AppendLog.Invoke("[ERROR] Failed to load profile: $_`r`n")
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

    if (-not $dlg.ShowDialog()) { return }

    try {
        $profilePath = $dlg.FileName
        $alias = [IO.Path]::GetFileNameWithoutExtension($profilePath)
        $alias = $alias -replace "_ssh_config$", ""

        $hostName = $HostBox.Text
        $port     = $PortBox.Text
        $rootUser = $UserBox.Text
        $newUser  = $NewUserBox.Text

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
        try { $script:AppendLog.Invoke("Profile saved to: $profilePath`r`n") } catch {}
    }
    catch {
        try { $script:AppendLog.Invoke("[ERROR] Failed to save profile: $_`r`n") } catch {}
    }
})

if ($BackupP12Button) {
    $BackupP12Button.Add_Click({
        Invoke-P12BackupFlow
    })
}

$MenuExit.Add_Click({
    $MainWindow.Close()
})

if ($ResetHostKeysButton) {
    $ResetHostKeysButton.Add_Click({

        $h = ""
        $p = ""

        try { $h = ($HostBox.Text).Trim() } catch { $h = "" }
        try { $p = ($PortBox.Text).Trim() } catch { $p = "" }

        if ([string]::IsNullOrWhiteSpace($h) -or [string]::IsNullOrWhiteSpace($p) -or ($p -notmatch '^\d+$')) {
            try { $script:AppendLog.Invoke("[ERROR] Enter a valid Host and numeric Port before resetting host keys.`r`n") } catch {}
            try {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "Enter a valid Host and numeric Port first.",
                    "Reset Host Keys",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                ) | Out-Null
            } catch {}
            return
        }

        try { $script:AppendLog.Invoke("[INFO] Resetting host keys for $h (port $p)...`r`n") } catch {}

        try {
            Reset-HostKeysForServer -HostName $h -Port $p
            try { $script:AppendLog.Invoke("[INFO] Host keys reset complete for $h (port $p).`r`n") } catch {}
        } catch {
            try { $script:AppendLog.Invoke("[ERROR] Host key reset failed: $($_.Exception.Message)`r`n") } catch {}
        }
    })
}

$StartButton.Add_Click({

    try {
        $LogBox.Document.Blocks.Clear()
        $LogBox.Document.Blocks.Add((New-Object System.Windows.Documents.Paragraph))
    } catch {}

    try { Start-LogTail } catch {}

    try { & $script:SetUIEnabledSB $false } catch {}
    try { $StartButton.IsEnabled = $false } catch {}
    try { $StartButton.Content = "Working..." } catch {}
    try { $script:IsBackgroundRunning = $true } catch {}

    $ui_serverHost = ""
    $ui_port       = ""
    $ui_user       = ""
    $ui_pass       = ""
    $ui_keyPath    = ""
    $ui_newUser    = ""
    $ui_newPass    = ""
    $ui_disableRootChecked = $false

    $ui_uploadP12  = $false
    $ui_p12Path    = ""
    $ui_p12Pass    = ""
    $ui_p12Alias   = ""
    $ui_p12NodeId  = ""

    try { $ui_serverHost = $HostBox.Text.Trim() } catch {}
    try { $ui_port       = $PortBox.Text.Trim() } catch {}
    try { $ui_user       = $UserBox.Text.Trim() } catch {}
    try { $ui_pass       = $PasswordBox.Password } catch {}
    try { $ui_keyPath    = $KeyBox.Text.Trim() } catch {}
    try { $ui_newUser    = $NewUserBox.Text.Trim() } catch {}
    try { $ui_newPass    = $NewPasswordBox.Password } catch {}
    try { $ui_disableRootChecked = ($DisableRootCheck.IsChecked -eq $true) } catch {}

    try { $ui_uploadP12 = ($UploadP12Check -and ($UploadP12Check.IsChecked -eq $true)) } catch {}
    try { $ui_p12Path   = if ($P12PathBox) { $P12PathBox.Text.Trim() } else { "" } } catch {}

    if ($ui_uploadP12 -eq $true) {

        if ([string]::IsNullOrWhiteSpace($ui_p12Path) -or -not (Test-Path -LiteralPath $ui_p12Path)) {
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] Upload P12 is enabled, but the selected file was not found.`r`n") } } catch {}
            try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
            return
        }

        $ext = ""
        try { $ext = [IO.Path]::GetExtension($ui_p12Path).ToLower() } catch { $ext = "" }
        if ($ext -ne ".p12") {
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] Upload P12 is enabled, but the selected file is not a .p12 file.`r`n") } } catch {}
            try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
            return
        }

        try { if ($P12PasswordBox) { $ui_p12Pass = $P12PasswordBox.Password } } catch { $ui_p12Pass = "" }
        if ([string]::IsNullOrWhiteSpace($ui_p12Pass)) {
            try {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "Please enter your P12 passphrase in the P12 Upload section before starting setup.",
                    "P12 Passphrase Required",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                ) | Out-Null
            } catch {}
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] Upload P12 is enabled, but P12 Passphrase is blank.`r`n") } } catch {}
            try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
            return
        }

        try {
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($ui_p12Path, $ui_p12Pass, 'Exportable')

            if (-not [string]::IsNullOrWhiteSpace($cert.FriendlyName)) {
                $ui_p12Alias = $cert.FriendlyName
                try { if ($script:AppendLog) { $script:AppendLog.Invoke("[INFO] P12 alias detected locally (friendlyName): $ui_p12Alias`r`n") } } catch {}
            } else {
                $ui_p12Alias = ""
                try { if ($script:AppendLog) { $script:AppendLog.Invoke("[WARN] P12 passphrase validated, but no friendlyName/alias was found. Upload will continue.`r`n") } } catch {}
            }

            try {
                $ui_p12NodeId = Get-NodeIdFromP12Local -P12Path $ui_p12Path -P12Pass $ui_p12Pass

                if ($ui_p12NodeId -match '^[0-9a-f]{128}$') {

                    $shortId = Get-ShortNodeId $ui_p12NodeId

                    try {
                        if ($script:AppendLog) {
                            $script:AppendLog.Invoke("__GUIONLY__Node ID extraction successful...`r`n")
                            if ($shortId) {
                                $script:AppendLog.Invoke("__GUIONLY__Short ID: $shortId`r`n")
                            }
                        }
                    } catch {}

                } else {

                    $ui_p12NodeId = ""
                    try { if ($script:AppendLog) { $script:AppendLog.Invoke("[WARN] Node ID could not be extracted locally (unexpected key type or format).`r`n") } } catch {}
                }

            } catch {
                $ui_p12NodeId = ""
                try { if ($script:AppendLog) { $script:AppendLog.Invoke("[WARN] Local Node ID extraction crashed: $($_.Exception.Message)`r`n") } } catch {}
            }

            try { $script:LastValidatedP12Alias = $ui_p12Alias } catch {}
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[INFO] P12 passphrase validated locally. Upload will proceed.`r`n") } } catch {}

        } catch {
            try { if ($script:AppendLog) { $script:AppendLog.Invoke("[ERROR] P12 passphrase is incorrect (or P12 is invalid).`r`n") } } catch {}
            try {
                [System.Windows.MessageBox]::Show(
                    $MainWindow,
                    "The P12 passphrase appears to be incorrect (or the P12 file is invalid).`r`n`r`nPlease re-enter the correct passphrase and try again.",
                    "P12 Passphrase Incorrect",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
            } catch {}
            try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
            return
        }
    }

    if (-not (Ensure-SshAgentSessionPreStart -KeyPath $ui_keyPath -Passphrase $ui_pass)) {
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("[FATAL] Cannot continue without enabling ssh-agent for encrypted key automation.`r`n") } } catch {}
        try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
        return
    }

    $ui_agentPubKey = ""

    try { if ($script:AppendLog) { $script:AppendLog.Invoke("[INFO] Start clicked. Launching background setup...`r`n") } } catch {}

    try {
        $script:TaskResult.ServerHost = ""
        $script:TaskResult.SshPort    = ""
        $script:TaskResult.User       = ""
        $script:TaskResult.Key        = ""
        $script:TaskResult.Prompt     = 0
        $script:TaskResult.Success    = 0
        $script:TaskResult.RootDisabled = 0

        $script:TaskResult.StatusTitle   = ""
        $script:TaskResult.StatusMessage = ""
        $script:TaskResult.StatusIcon    = "Information"
        $script:TaskResult.ShowStatus    = 0

        $script:TaskResult.P12Alias      = ""
        $script:TaskResult.P12RemotePath = ""
        $script:TaskResult.P12LocalPath  = ""
        $script:TaskResult.P12NodeId     = ""
        $script:TaskResult.ShowP12Alias  = 0
    } catch {}

    $ui_createNonRoot = $false
    try { $ui_createNonRoot = ($CreateNonRootCheck -and ($CreateNonRootCheck.IsChecked -eq $true)) } catch {}

    $ctx = @{
        ui_serverHost         = $ui_serverHost
        ui_port               = $ui_port
        ui_user               = $ui_user
        ui_pass               = $ui_pass
        ui_keyPath            = $ui_keyPath
        ui_newUser            = $ui_newUser
        ui_newPass            = $ui_newPass
        ui_disableRootChecked = $ui_disableRootChecked

        ui_createNonRoot      = $ui_createNonRoot

        ui_uploadP12          = $ui_uploadP12
        ui_p12Path            = $ui_p12Path
        ui_p12Pass            = $ui_p12Pass
        ui_p12Alias           = $ui_p12Alias
        ui_p12NodeId          = $ui_p12NodeId
        ui_agentPubKey        = $ui_agentPubKey

        TaskResult            = $script:TaskResult
    }

    $task = Invoke-BackgroundUI -Context $ctx -Work {

        function Safe-AppendLog {
            param([string]$t)
            try { if ($AppendLog) { $AppendLog.Invoke($t) } } catch {}
        }

        try {
            Add-Content -Path $ServerToolkitLogPath -Value ("WORKER_ENTERED " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}

        try {
            Safe-AppendLog "=== StartButton clicked at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===`r`n"

            $serverHost = $ui_serverHost
            $port       = $ui_port
            $user       = $ui_user
            $pass       = $ui_pass
            $keyPath    = $ui_keyPath
            $newUser    = $ui_newUser
            $newPass    = $ui_newPass

            $disableRootChecked = ($ui_disableRootChecked -eq $true)
            $doCreateNonRoot    = ($ui_createNonRoot -eq $true)

            $doP12 = ($ui_uploadP12 -eq $true)
            $p12LocalPath = $ui_p12Path
            $p12Alias     = $ui_p12Alias
            $p12NodeId    = $ui_p12NodeId

            try {
                if ($TaskResult) {
                    # Fill these so we can use them later if the connection is verified
                    $TaskResult.ServerHost = $serverHost
                    $TaskResult.SshPort    = $port
                    $TaskResult.User       = $user
                    $TaskResult.Key        = $keyPath

                    $TaskResult.Prompt     = 0
                    $TaskResult.Success    = 0
                }
            } catch {}

            if ([string]::IsNullOrWhiteSpace($serverHost) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($port)) {
                Safe-AppendLog "[ERROR] Host, Port, and Username are required fields.`r`n"
                return
            }
            if ($port -notmatch '^\d+$') {
                Safe-AppendLog "[ERROR] Port must be a number.`r`n"
                return
            }

            if ([string]::IsNullOrWhiteSpace($newUser)) {
                if ($user -eq "root") { $newUser = "nodeadmin" } else { $newUser = $user }
            }

            Reset-HostKeysForServer -HostName $serverHost -Port $port

            $authMode = "password"
            if (-not [string]::IsNullOrWhiteSpace($keyPath)) { $authMode = "key" }

            if ($authMode -eq "key") {
                Safe-AppendLog "Valid OpenSSH/PEM private key detected - using as-is: $keyPath`r`n"
                Safe-AppendLog "Using key-based authentication for $user@$serverHost (Password is treated as key passphrase).`r`n"

                $agentOk = Ensure-SshAgentWithKey -KeyPath $keyPath -Passphrase $pass
                if (-not $agentOk) {
                    Safe-AppendLog "[INFO] Waiting for ssh-add. Then click Start Setup again.`r`n"
                    return
                }
            } else {
                Safe-AppendLog "Using password-based authentication for $user@$serverHost.`r`n"
                if (-not (Ensure-PuttyTools)) {
                    Safe-AppendLog "[FATAL] plink.exe required for password-only mode.`r`n"
                    return
                }
            }

            function Run-RemoteCommand {
                param([string]$cmd, [int]$TimeoutMs = 60000, [string]$StdinText = $null)

                if ($authMode -eq "key") {
                    return (Run-SshCommand -RemoteCommand $cmd -RemoteHost $serverHost -Port $port -User $user -KeyPath $keyPath -TimeoutMs $TimeoutMs -StdinText $StdinText)
                } else {
                    $args = "-batch -ssh -P $port"
                    if (-not [string]::IsNullOrWhiteSpace($pass)) { $args += " -pw `"$pass`"" }
                    $args += " -l $user $serverHost `"$cmd`""

                    $psi = New-Object System.Diagnostics.ProcessStartInfo
                    $psi.FileName = $global:PlinkPath
                    $psi.Arguments = $args
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError  = $true
                    $psi.CreateNoWindow = $true

                    $p = New-Object System.Diagnostics.Process
                    $p.StartInfo = $psi
                    $null = $p.Start()
                    $p.WaitForExit()

                    $o = ""
                    $e = ""
                    try { $o = $p.StandardOutput.ReadToEnd() } catch {}
                    try { $e = $p.StandardError.ReadToEnd() } catch {}

                    if ($e) {
                        foreach ($ln in ($e -split "`r?`n")) { if ($ln) { Safe-AppendLog "[ERROR] $ln`r`n" } }
                    }
                    if ($o) {
                        foreach ($ln in ($o -split "`r?`n")) { if ($ln) { Safe-AppendLog "$ln`r`n" } }
                    }

                    return $p.ExitCode
                }
            }

            Safe-AppendLog "Connecting to $serverHost on port $port using OpenSSH...`r`n"
            $probeExit = Run-RemoteCommand -cmd "echo CONNECTED" -TimeoutMs 8000
            if ($probeExit -ne 0) {
                Safe-AppendLog "[FATAL] SSH connection failed. Aborting.`r`n"
                try {
                    if ($TaskResult) {
                        $TaskResult.Success = 0
                        $TaskResult.Prompt  = 0
                    }
                } catch {}
                return
            }

            try {
                if ($TaskResult) {
                    $TaskResult.Success = 1
                    $TaskResult.Prompt  = 1
                }
            } catch {}

            if ($doCreateNonRoot) {

                Safe-AppendLog "Checking whether user '$newUser' already exists...`r`n"
                $existsExit = Run-RemoteCommand -cmd "bash -lc 'id $newUser >/dev/null 2>&1'" -TimeoutMs 8000

                if ($existsExit -ne 0) {
                    Safe-AppendLog "Creating '$newUser' user...`r`n"
                    $cExit = Run-RemoteCommand -cmd "bash -lc 'timeout 10 useradd -m -s /bin/bash $newUser'" -TimeoutMs 20000
                    if ($cExit -ne 0) {
                        Safe-AppendLog "[FATAL] Failed to create user '$newUser'. Aborting.`r`n"
                        return
                    }
                    Safe-AppendLog "User '$newUser' created successfully.`r`n"
                } else {
                    Safe-AppendLog "[WARN] User '$newUser' already exists on this server.`r`n"
                }

                Safe-AppendLog "Adding '$newUser' to sudo group...`r`n"
                $gExit = Run-RemoteCommand -cmd "bash -lc 'timeout 10 usermod -aG sudo $newUser'" -TimeoutMs 20000
                if ($gExit -ne 0) {
                    Safe-AppendLog "[FATAL] Failed to add user '$newUser' to sudo group. Aborting.`r`n"
                    return
                }
                Safe-AppendLog "User '$newUser' is now in sudo group.`r`n"

                Safe-AppendLog "Ensuring home directory and SSH permissions...`r`n"
                $hExit = Run-RemoteCommand -cmd "bash -lc 'set -e; u=$newUser; mkdir -p /home/$newUser/.ssh; chown -R ${newUser}:$newUser /home/$newUser; chmod 700 /home/$newUser/.ssh'" -TimeoutMs 20000
                if ($hExit -ne 0) {
                    Safe-AppendLog "[FATAL] Failed to ensure home/ssh permissions for '$newUser'. Aborting.`r`n"
                    return
                }

                if (-not [string]::IsNullOrWhiteSpace($newPass)) {
                    Safe-AppendLog "Setting password for '$newUser'...`r`n"
                    $pair = "$newUser`:$newPass"
                    $b64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))
                    $pExit = Run-RemoteCommand -cmd "bash -lc `"echo $b64 | base64 -d | chpasswd`"" -TimeoutMs 180000
                    if ($pExit -ne 0) {
                        Safe-AppendLog "[FATAL] Failed to set password for '$newUser'. Aborting.`r`n"
                        return
                    }
                    Safe-AppendLog "Password set for '$newUser'.`r`n"
                }

                Safe-AppendLog "Authorizing SSH key for '$newUser' (agent pubkey -> authorized_keys)...`r`n"

                $kExit = 1
                $agentPubKey = ""

                try {
                    $sshAddExe2 = Join-Path $env:WINDIR "System32\OpenSSH\ssh-add.exe"
                    if (-not (Test-Path $sshAddExe2)) {
                        try { $sshAddExe2 = (Get-Command ssh-add -ErrorAction Stop).Source } catch { $sshAddExe2 = $null }
                    }

                    if ($sshAddExe2) {
                        $pubs2 = & $sshAddExe2 -L 2>$null
                        foreach ($ln2 in ($pubs2 -split "`r?`n")) {
                            if ($ln2 -and $ln2 -match '^(ssh-(rsa|ed25519)|ecdsa-sha2-nistp\d+)\s+') {
                                $agentPubKey = $ln2.Trim()
                                break
                            }
                        }
                    }
                } catch {
                    $agentPubKey = ""
                }

                if (-not [string]::IsNullOrWhiteSpace($agentPubKey)) {

                    $pubB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($agentPubKey))

                    $cmdAuth = "bash -lc 'set -e; u=$newUser; " +
                            "mkdir -p /home/$newUser/.ssh; " +
                            "printf %s $pubB64 | base64 -d > /home/$newUser/.ssh/authorized_keys; " +
                            "chown -R ${newUser}:$newUser /home/$newUser/.ssh; " +
                            "chmod 700 /home/$newUser/.ssh; " +
                            "chmod 600 /home/$newUser/.ssh/authorized_keys; " +
                            "echo AUTH_KEYS_WRITTEN'"

                    $kExit = Run-RemoteCommand -cmd $cmdAuth -TimeoutMs 20000

                } else {

                    Safe-AppendLog "[FATAL] Could not read a public key from ssh-agent after unlock. Aborting.`r`n"
                    return
                }

                if ($kExit -ne 0) {
                    Safe-AppendLog "[FATAL] Failed to authorize SSH key for '$newUser'. Aborting.`r`n"
                    return
                }

                Safe-AppendLog "SSH key authorized for '$newUser'.`r`n"

                Safe-AppendLog "Testing login for new user '$newUser'...`r`n"
                if ($authMode -eq "key") {
                    $testExit = (Run-SshCommand -RemoteCommand "echo OK" -RemoteHost $serverHost -Port $port -User $newUser -KeyPath $keyPath -TimeoutMs 8000)
                    if ($testExit -ne 0) {
                        Safe-AppendLog "[FATAL] New user '$newUser' login test failed. Aborting.`r`n"
                        return
                    }
                }
                Safe-AppendLog "New user '$newUser' login test successful.`r`n"

                if ($disableRootChecked -and $user -eq "root") {
                    Safe-AppendLog "Disabling root SSH login and global password auth...`r`n"

                    $cmdDisableRoot = @"
bash -lc '
set -e
mkdir -p /etc/ssh/sshd_config.d
cat >/etc/ssh/sshd_config.d/99-server-toolkit.conf <<EOF
PermitRootLogin no
PasswordAuthentication no
KbdInteractiveAuthentication no
ChallengeResponseAuthentication no
EOF
sshd -t
( sleep 1; systemctl reload sshd 2>/dev/null || systemctl reload ssh 2>/dev/null || systemctl restart sshd 2>/dev/null || systemctl restart ssh 2>/dev/null || true ) >/dev/null 2>&1 &
exit 0
'
"@
                    $dExit = Run-RemoteCommand -cmd $cmdDisableRoot -TimeoutMs 20000
                    if ($dExit -ne 0) {
                        Safe-AppendLog "[FATAL] Failed to disable root login. Aborting.`r`n"
                        return
                    }

                    try { if ($TaskResult) { $TaskResult.RootDisabled = 1 } } catch {}

                    Safe-AppendLog "Root login and password authentication disabled.`r`n"
                }

                Safe-AppendLog "[INFO] Switching operations to '$newUser' for non-root tasks...`r`n"
                $user = $newUser
                try { if ($TaskResult) { $TaskResult.User = $user } } catch {}
            }

            $remoteP12Path = ""
            if ($doP12) {

                Safe-AppendLog "[INFO] Upload P12 enabled. Starting secure upload...`r`n"

                $p12Leaf = (Split-Path -Leaf $p12LocalPath)
                $homeDirResolved = if ($user -eq "root") { "/root" } else { "/home/$user" }
                $remoteP12Path = "$homeDirResolved/$p12Leaf"

                $existsExit = 1
                if ($authMode -eq "key") {
                    $existsExit = (Run-SshCommand -RemoteCommand "bash -lc 'test -f ""$remoteP12Path""'" -RemoteHost $serverHost -Port $port -User $user -KeyPath $keyPath -TimeoutMs 8000)
                } else {
                    $existsExit = Run-RemoteCommand -cmd "bash -lc 'test -f ""$remoteP12Path""'" -TimeoutMs 8000
                }

                if ($existsExit -eq 0 -and $P12OverwritePromptFunc) {
                    $evt = New-Object System.Threading.ManualResetEventSlim($false)
                    $box = [hashtable]::Synchronized(@{ Decision = "skip" })

                    $uiAsk = [Action]{
                        try { $box.Decision = $P12OverwritePromptFunc.Invoke($remoteP12Path, $p12LocalPath) } catch { $box.Decision = "skip" }
                        try { $evt.Set() } catch {}
                    }

                    try { $null = $UiDispatcher.BeginInvoke($uiAsk) } catch { try { $uiAsk.Invoke() } catch {} }
                    try { $null = $evt.Wait([TimeSpan]::FromMinutes(5)) } catch {}

                    if ($box.Decision -eq "skip") {
                        Safe-AppendLog "[INFO] Remote P12 exists. Skipping P12 handling and continuing setup.`r`n"
                        $doP12 = $false
                    }
                    elseif ($box.Decision -eq "backup") {
                        Safe-AppendLog "[INFO] Remote P12 exists. Creating backup before overwrite...`r`n"
                        Run-RemoteCommand -cmd "bash -lc 'set -e; ts=$(date +%Y%m%d-%H%M%S); cp ""$remoteP12Path"" ""$remoteP12Path.bak-$ts"" || true'" -TimeoutMs 20000 | Out-Null
                    }
                }

                if ($doP12) {

                    if ($authMode -eq "key") {
                        $scpExe = Join-Path $env:WINDIR "System32\OpenSSH\scp.exe"
                        if (-not (Test-Path $scpExe)) {
                            try { $scpExe = (Get-Command scp -ErrorAction Stop).Source } catch { $scpExe = $null }
                        }
                        if (-not $scpExe) {
                            Safe-AppendLog "[FATAL] scp.exe not found. Cannot upload P12.`r`n"
                            return
                        }

                        $knownHostsPath = Join-Path $env:USERPROFILE ".ssh\known_hosts"
                        $scpArgs = @(
                            "-p",
                            "-P", $port,
                            "-o", "StrictHostKeyChecking=accept-new",
                            "-o", "UserKnownHostsFile=$knownHostsPath",
                            "-o", "GlobalKnownHostsFile=NUL"
                        )
                        if ($keyPath) { $scpArgs += @("-i", "`"$keyPath`"") }

                        $scpArgs += @(
                            "`"$p12LocalPath`"",
                            "${user}@${serverHost}:$remoteP12Path"
                        )

                        Safe-AppendLog "[INFO] Uploading P12 via SCP as '$user'...`r`n"
                        $psi = New-Object System.Diagnostics.ProcessStartInfo
                        $psi.FileName = $scpExe
                        $psi.Arguments = ($scpArgs -join " ")
                        $psi.UseShellExecute = $false
                        $psi.RedirectStandardError = $true
                        $psi.RedirectStandardOutput = $true
                        $psi.CreateNoWindow = $true

                        $proc = New-Object System.Diagnostics.Process
                        $proc.StartInfo = $psi
                        $proc.Start() | Out-Null
                        $proc.WaitForExit()

                        if ($proc.ExitCode -ne 0) {
                            $err = ""
                            try { $err = $proc.StandardError.ReadToEnd() } catch {}
                            Safe-AppendLog "[FATAL] P12 upload failed: $err`r`n"
                            return
                        }

                        Safe-AppendLog "[INFO] P12 uploaded to: $remoteP12Path`r`n"
                    }
                    else {
                        if (-not $global:PscpPath -or -not (Test-Path $global:PscpPath)) {
                            Safe-AppendLog "[FATAL] pscp.exe not found. Cannot upload P12 in password mode.`r`n"
                            return
                        }

                        Safe-AppendLog "[INFO] Uploading P12 via PSCP as '$user' (password mode)...`r`n"

                        $pscpArgs = @("-batch","-scp","-p","-P",$port)
                        $pscpArgs += @("-pw","`"$newPass`"","`"$p12LocalPath`"","${user}@${serverHost}:$remoteP12Path")

                        $psi = New-Object System.Diagnostics.ProcessStartInfo
                        $psi.FileName = $global:PscpPath
                        $psi.Arguments = ($pscpArgs -join " ")
                        $psi.UseShellExecute = $false
                        $psi.RedirectStandardError = $true
                        $psi.RedirectStandardOutput = $true
                        $psi.CreateNoWindow = $true

                        $proc = New-Object System.Diagnostics.Process
                        $proc.StartInfo = $psi
                        $proc.Start() | Out-Null
                        $proc.WaitForExit()

                        if ($proc.ExitCode -ne 0) {
                            $err = ""
                            try { $err = $proc.StandardError.ReadToEnd() } catch {}
                            Safe-AppendLog "[FATAL] P12 upload failed (pscp): $err`r`n"
                            return
                        }

                        Safe-AppendLog "[INFO] P12 uploaded to: $remoteP12Path`r`n"
                    }

                    $chmodExit = Run-RemoteCommand -cmd "bash -lc 'chmod 600 ""$remoteP12Path""'" -TimeoutMs 20000
                    if ($chmodExit -ne 0) {
                        Safe-AppendLog "[FATAL] Failed to set permissions on: $remoteP12Path`r`n"
                        return
                    }

                    Safe-AppendLog "[INFO] P12 installed: $remoteP12Path`r`n"

                    try {
                        if ($TaskResult) {
                            $TaskResult.P12RemotePath = $remoteP12Path
                            $TaskResult.P12LocalPath  = $p12LocalPath
                            $TaskResult.P12Alias      = $p12Alias
                            $TaskResult.P12NodeId     = ""
                            if (-not [string]::IsNullOrWhiteSpace($p12NodeId)) { $TaskResult.P12NodeId = $p12NodeId }
                            $TaskResult.ShowP12Alias  = 1
                        }
                    } catch {}
                }
            }

            try {
                if ($doP12 -eq $true) {
                    Safe-AppendLog "[INFO] Setup complete - P12 upload successful.`r`n"
                } else {
                    Safe-AppendLog "[INFO] Setup complete.`r`n"
                }
            } catch {}

            try {
                if ($TaskResult) {
                    $TaskResult.ServerHost = $serverHost
                    $TaskResult.SshPort    = $port
                    $TaskResult.User       = $user
                    $TaskResult.Key        = $keyPath
                    $TaskResult.Prompt     = 1
                }
            } catch {}

        } catch {
            Safe-AppendLog "[FATAL] Background task crashed: $($_.Exception.Message)`r`n"
            if ($_.ScriptStackTrace) { Safe-AppendLog "[FATAL] Stack:`r`n$($_.ScriptStackTrace)`r`n" }
        }

        try {
            Add-Content -Path $ServerToolkitLogPath -Value ("WORKER_FINALLY_REACHED " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + "`r`n") -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch {}
    }

    if (-not $task) {
        try { if ($script:AppendLog) { $script:AppendLog.Invoke("[FATAL] Background setup did not start. Re-enabling UI.`r`n") } } catch {}
        try { if ($UnlockUIAction) { $UnlockUIAction.Invoke() } } catch {}
        return
    }

    try { if ($script:AppendLog) { $script:AppendLog.Invoke("[INFO] Background task started. TaskId=$($task.Id) Status=$($task.Status)`r`n") } } catch {}
    try { $script:ActiveTask = $task } catch {}
    try { Start-TaskWatchTimer -Task $task } catch {}

})


$MainWindow.add_Closing({
    param($sender, $e)

    try {
        if ($script:IsBackgroundRunning) {

            $e.Cancel = $true

            try {
                Add-Content -Path $global:ServerToolkitLogPath `
                    -Value ((Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " [WARN] Close requested while background task is running.") `
                    -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}

            $msg = @()
            $msg += "Setup is still running."
            $msg += ""
            $msg += "Do you want to FORCE STOP the running task and close Server Toolkit?"
            $msg += ""
            $msg += "Yes  = Stop + Close"
            $msg += "No   = Keep running"

            $res = [System.Windows.MessageBox]::Show(
                $MainWindow,
                ($msg -join "`r`n"),
                "Close while running?",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )

            if ($res -eq [System.Windows.MessageBoxResult]::Yes) {

                try { Stop-ActiveBackgroundTask -Reason "user-close" } catch {}
                try { Signal-SshAgentSessionEnd } catch {}

                $e.Cancel = $false
                return
            }

            return
        }

        try { Signal-SshAgentSessionEnd } catch {}
    } catch {}
})

function Should-RunAgentCleanup {
    try {
        if ($script:SshAgentTouched -eq $true) { return $true }
    } catch {}

    try {
        $f = Get-ServerToolkitOwnedPubKeysFile
        if ($f -and (Test-Path -LiteralPath $f)) { return $true }
    } catch {}

    return $false
}

$MainWindow.add_Closed({
    try { Signal-SshAgentSessionEnd } catch {}
    try {
        if (Should-RunAgentCleanup) {
            Cleanup-AgentOwnedKeys
        }
    } catch {}
})

try {
    Register-EngineEvent -SourceIdentifier "ServerToolkit_Exiting" -InputObject ([AppDomain]::CurrentDomain) -EventName "ProcessExit" -Action {
        try {
            if (Get-Command Should-RunAgentCleanup -ErrorAction SilentlyContinue) {
                if (Should-RunAgentCleanup) { Cleanup-AgentOwnedKeys }
            } else {
                Cleanup-AgentOwnedKeys
            }
        } catch {}
    } | Out-Null
} catch {}

try {

    Add-Content -Path $global:ServerToolkitLogPath `
        -Value ("DEBUG: Launching MainWindow via ShowDialog() (clean)" + "`r`n") `
        -Encoding UTF8 -ErrorAction SilentlyContinue

    $MainWindow.WindowState   = [System.Windows.WindowState]::Normal
    $MainWindow.ShowInTaskbar = $true
    $MainWindow.Topmost       = $false

    $script:StartupWatchdogTimer = $null
    $script:StartupWatchdogArmed = $true

    try {
        $script:StartupWatchdogTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:StartupWatchdogTimer.Interval = [TimeSpan]::FromSeconds(12)

        $script:StartupWatchdogTimer.Add_Tick({
            try { $script:StartupWatchdogTimer.Stop() } catch {}
            if (-not $script:StartupWatchdogArmed) { return }

            try {
                Add-Content -Path $global:ServerToolkitLogPath -Value (
                    (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                    " [FATAL] STARTUP_WATCHDOG_TIMEOUT: window did not render within 12 seconds. Forcing process exit.`r`n"
                ) -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}

            try { [Environment]::Exit(86) } catch {}
            try { Stop-Process -Id $PID -Force } catch {}
        })

        $script:StartupWatchdogTimer.Start() | Out-Null
    } catch {}

    try {
        $MainWindow.Add_ContentRendered({
            try { $script:StartupWatchdogArmed = $false } catch {}
            try { if ($script:StartupWatchdogTimer) { $script:StartupWatchdogTimer.Stop() } } catch {}

            try {
                if ($HostBox -and $HostBox.IsEnabled) {
                    $HostBox.Focus()
                    $HostBox.CaretIndex = $HostBox.Text.Length
                }
            } catch {}

            try {
                Add-Content -Path $global:ServerToolkitLogPath -Value (
                    (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") +
                    " DEBUG: ContentRendered fired; startup watchdog disarmed; HostBox focused.`r`n"
                ) -Encoding UTF8 -ErrorAction SilentlyContinue
            } catch {}
        }) | Out-Null
    } catch {}

    $null = $MainWindow.ShowDialog()

} catch {

    try {
        Add-Content -Path $global:ServerToolkitLogPath `
            -Value ("[FATAL] Window start failed: " + $_.Exception.ToString() + "`r`n") `
            -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}

    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show(
            "Server Toolkit failed to start.`r`n`r`n" +
            "See log:`r`n$global:ServerToolkitLogPath`r`n`r`n" +
            $_.Exception.ToString(),
            "Server Toolkit Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    } catch {}

    throw
}
