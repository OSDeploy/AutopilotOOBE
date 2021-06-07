#=======================================================================
#   PowershellWindow Functions
#=======================================================================
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
function Show-PowershellWindow() {
    $null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
function Hide-PowershellWindow() {
    $null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}

Hide-PowershellWindow
#=======================================================================
#   MahApps.Metro
#=======================================================================
# Assign current script directory to a global variable
$Global:MyScriptDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)

# Load presentationframework and Dlls for the MahApps.Metro theme
[System.Reflection.Assembly]::LoadWithPartialName("presentationframework") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$Global:MyScriptDir\assembly\System.Windows.Interactivity.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$Global:MyScriptDir\assembly\MahApps.Metro.dll") | Out-Null

# Set console size and title
$host.ui.RawUI.WindowTitle = "Start-AutopilotOOBE"
#=======================================================================
#   LoadForm
#=======================================================================
function LoadForm {
    [CmdletBinding()]
    param()

    [xml]$Global:xmlWPF = @"
    <Controls:MetroWindow
        xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x = "http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d = "http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:Controls = "http://metro.mahapps.com/winfx/xaml/controls"

        Title = ""
        BorderBrush = "{DynamicResource AccentColorBrush}"
        BorderThickness = "1"
        Width = "980"
        Height = "670"
        Background = "#004275"
        
        ResizeMode = "CanResizeWithGrip"
        WindowStartupLocation = "CenterScreen"
        WindowStyle = "None"
    >

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <!-- MahApps.Metro resource dictionaries. Make sure that all file names are Case Sensitive! -->
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
            </ResourceDictionary.MergedDictionaries>

            <Style TargetType="{x:Type Window}">
                <Setter Property="FontFamily" Value="Segoe UI" />
                <Setter Property="FontWeight" Value="Light" />
                <Setter Property="Background" Value="#1f1f1f" />
                <Setter Property="Foreground" Value="white" />
            </Style>

            <Style TargetType="{x:Type Button}">
                <Setter Property="Background" Value="#FF1D3245" />
                <Setter Property="Foreground" Value="#FFE8EDF9" />
                <Setter Property="FontSize" Value="15" />
                <Setter Property="SnapsToDevicePixels" Value="True" />

                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button" >
                            <Border Name = "Border" 
                                BorderThickness = "1"
                                Padding = "4,2" 
                                BorderBrush = "#336891" 
                                CornerRadius = "1" 
                                Background = "#0078d7">
                                <ContentPresenter HorizontalAlignment = "Center" VerticalAlignment = "Center" TextBlock.TextAlignment = "Center" />
                            </Border>

                            <ControlTemplate.Triggers>
                                <Trigger Property = "IsFocused" Value = "False">
                                    <Setter TargetName = "Border" Property = "BorderBrush" Value = "#336891" />
                                    <Setter Property = "Button.Background" Value = "#336891" />
                                </Trigger>
                                <Trigger Property = "IsMouseOver" Value="True">
                                    <Setter TargetName = "Border" Property = "BorderBrush" Value = "#FFE8EDF9" />
                                </Trigger>
                                <Trigger Property = "IsEnabled" Value = "False">
                                    <Setter TargetName = "Border" Property = "BorderBrush" Value = "#336891" />
                                    <Setter Property = "Button.Foreground" Value = "#336891" />
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>


            <Style TargetType="{x:Type Label}">
                <Setter Property = "FontFamily" Value = "Segoe UI" />
            </Style>

            <Style TargetType="{x:Type TextBox}">
                <Setter Property = "FontFamily" Value = "Segoe UI" />
            </Style>

            <Style TargetType="{x:Type ComboBox}">
                <Setter Property = "FontFamily" Value = "Segoe UI" />
            </Style>

        </ResourceDictionary>
    </Window.Resources>

        <Grid>
            <StackPanel Width = "700">
                <Label Name = "Title1"
                Content = "Autopilot Azure Join"
                FontFamily = "Segoe UI Light"
                FontSize = "40"
                Foreground = "White"
                HorizontalAlignment = "Center"
                Margin = "10,10,10,10"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Label Name = "GroupTagLabel"
                Content = "GroupTag:"
                FontFamily = "Segoe UI"
                FontSize = "16"
                Foreground = "White"
                HorizontalAlignment = "Stretch"
                Margin = "0,90,0,0"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <TextBox Name = "GroupTagTextBox"
                Background = "White"
                FontFamily = "Segoe UI"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Stretch"
                Margin = "200,80,0,0"
                Width = "500"
                Padding = "8"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Label Name = "AddToGroupLabel"
                Content = "AddToGroup:"
                FontFamily = "Segoe UI"
                FontSize = "16"
                Foreground = "White"
                HorizontalAlignment = "Stretch"
                Margin = "0,150,0,0"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <TextBox Name = "AddToGroupTextBox"
                Background = "White"
                FontFamily = "Segoe UI"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Stretch"
                Margin = "200,140,0,0"
                Width = "500"
                Padding = "8"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Label Name = "AssignedUserLabel"
                Content = "AssignedUser:"
                FontFamily = "Segoe UI"
                FontSize = "16"
                Foreground = "White"
                HorizontalAlignment = "Stretch"
                Margin = "0,210,0,0"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <TextBox Name = "AssignedUserTextBox"
                Background = "White"
                FontFamily = "Segoe UI"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Stretch"
                Margin = "200,200,0,0"
                Width = "500"
                Padding = "8"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Label Name = "AssignedComputerNameLabel"
                Content = "AssignedComputerName:"
                FontFamily = "Segoe UI"
                FontSize = "16"
                Foreground = "White"
                HorizontalAlignment = "Stretch"
                Margin = "0,270,0,0"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <TextBox Name = "AssignedComputerNameTextBox"
                Background = "Gray"
                FontFamily = "Segoe UI"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Stretch"
                Margin = "200,260,0,0"
                Text = "Azure AD Join Only"
                Width = "500"
                Padding = "8"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Label Name = "PostActionLabel"
                Content = "PostAction:"
                FontFamily = "Segoe UI"
                FontSize = "16"
                Foreground = "White"
                HorizontalAlignment = "Stretch"
                Margin = "0,330,0,0"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <ComboBox Name = "PostActionComboBox"
                Background = "White"
                FontFamily = "Segoe UI"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Stretch"
                Margin = "200,320,0,0"
                Width = "500"
                Padding = "8"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <CheckBox Name = "AssignCheckbox"
                HorizontalAlignment = "Stretch"
                Background = "White"
                FontFamily = "Segoe UI"
                FontSize = "16"
                Foreground = "White"
                Margin = "200,380,0,0"
                >Assign: Wait for Intune to assign an Autopilot profile to the device
                </CheckBox>
            </StackPanel>

            <StackPanel Width = "700">
                <Button Name = "RegisterButton"
                Content = "Register"
                FontFamily = "Segoe UI"
                FontSize = "16"
                FontWeight = "Bold"
                Height = "40"
                HorizontalAlignment = "Right"
                Margin = "0,420,0,0"
                Width = "170"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Button Name = "RunButton"
                Content = "Run"
                FontSize = "16"
                FontWeight = "Bold"
                Height = "40"
                HorizontalAlignment = "Left"
                Margin = "0,510,0,0"
                Width = "100"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <ComboBox Name = "RunComboBox"
                Background = "White"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Left"
                Margin = "100,510,0,0"
                Padding = "8"
                Width = "600"
                />
            </StackPanel>

            <StackPanel Width = "700">
                <Button Name = "StartButton"
                Content = "Docs"
                FontSize = "16"
                FontWeight = "Bold"
                Height = "40"
                HorizontalAlignment = "Left"
                Margin = "0,570,0,0"
                Width = "100"
                />
            </StackPanel>
            <StackPanel Width = "700">
                <TextBox Name = "StartTextBox"
                Background = "White"
                FontSize = "16"
                FontWeight = "Bold"
                Foreground = "Black"
                Height = "40"
                HorizontalAlignment = "Left"
                Margin = "100,570,0,0"
                Padding = "8"
                Width = "600"
                />
            </StackPanel>
        </Grid>
    </Controls:MetroWindow>
"@

    # Add WPF and Windows Forms assemblies
    Try {
        Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms
    } 
    Catch {
        Throw "Failed to load Windows Presentation Framework assemblies."
    }

    #Create the XAML reader using a new XML node reader
    $Global:xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))

    #Create hooks to each named object in the XAML
    $xmlWPF.SelectNodes("//*[@Name]") | foreach {
        Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global
    }
}
#=======================================================================
#   LoadForm
#=======================================================================
LoadForm
#=======================================================================
#   Parameters
#=======================================================================
$AutopilotOOBEParams = (Get-Command Start-AutopilotOOBE).Parameters
#=======================================================================
#   Parameter GroupTag
#=======================================================================
$GroupTagTextBox.Text = $Global:AutopilotOOBE.GroupTag
#=======================================================================
#   Parameter AddToGroup
#=======================================================================
$AddToGroupTextBox.Text = $Global:AutopilotOOBE.AddToGroup
#=======================================================================
#   Parameter PostAction
#=======================================================================
#$AutopilotOOBEParams["PostAction"].Attributes.ValidValues | ForEach-Object {
#    $PostActionComboBox.Items.Add($_) | Out-Null
#}
#$PostActionComboBox.SelectedValue = $Global:AutopilotOOBE.PostAction
$PostActionComboBox.Items.Add('') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /quit') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /reboot') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /shutdown') | Out-Null

if ($Global:AutopilotOOBE.PostAction -eq 'Sysprep') {$PostActionComboBox.SelectedValue = 'Sysprep /oobe /quit'}
if ($Global:AutopilotOOBE.PostAction -eq 'SysprepReboot') {$PostActionComboBox.SelectedValue = 'Sysprep /oobe /reboot'}
if ($Global:AutopilotOOBE.PostAction -eq 'SysprepShutdown') {$PostActionComboBox.SelectedValue = 'Sysprep /oobe /shutdown'}
#=======================================================================
#   Parameter Assign
#=======================================================================
if ($Global:AutopilotOOBE.Assign -eq $true) {$AssignCheckBox.IsChecked = $true}
#=======================================================================
#   Parameter Run
#=======================================================================
#$AutopilotOOBEParams["Run"].Attributes.ValidValues | ForEach-Object {
#    $RunComboBox.Items.Add($_) | Out-Null
#}
#$RunComboBox.SelectedValue = $Global:AutopilotOOBE.Run
$RunComboBox.Items.Add('Command Prompt') | Out-Null
$RunComboBox.Items.Add('PowerShell') | Out-Null
$RunComboBox.Items.Add('PowerShell ISE') | Out-Null
$RunComboBox.Items.Add('Open Windows Explorer') | Out-Null
$RunComboBox.Items.Add('Open Windows Settings') | Out-Null
$RunComboBox.Items.Add('Open Network and Wireless Settings') | Out-Null
$RunComboBox.Items.Add('Restart Computer') | Out-Null
$RunComboBox.Items.Add('Shutdown Computer') | Out-Null
$RunComboBox.Items.Add('Sysprep /oobe /quit') | Out-Null
$RunComboBox.Items.Add('Sysprep /oobe /reboot') | Out-Null
$RunComboBox.Items.Add('Sysprep /oobe /shutdown') | Out-Null
$RunComboBox.Items.Add('MDMDiagnosticsTool -out C:\Temp') | Out-Null
$RunComboBox.Items.Add('MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab') | Out-Null
$RunComboBox.Items.Add('MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\Autopilot.cab') | Out-Null

if ($Global:AutopilotOOBE.Run -eq 'CommandPrompt') {$RunComboBox.SelectedValue = 'Command Prompt'}
if ($Global:AutopilotOOBE.Run -eq 'PowerShell') {$RunComboBox.SelectedValue = 'PowerShell'}
if ($Global:AutopilotOOBE.Run -eq 'PowerShellISE') {$RunComboBox.SelectedValue = 'PowerShell ISE'}
if ($Global:AutopilotOOBE.Run -eq 'WindowsExplorer') {$RunComboBox.SelectedValue = 'Open Windows Explorer'}
if ($Global:AutopilotOOBE.Run -eq 'WindowsSettings') {$RunComboBox.SelectedValue = 'Open Windows Settings'}
if ($Global:AutopilotOOBE.Run -eq 'NetworkingWireless') {$RunComboBox.SelectedValue = 'Open Network and Wireless Settings'}
if ($Global:AutopilotOOBE.Run -eq 'Restart') {$RunComboBox.SelectedValue = 'Restart Computer'}
if ($Global:AutopilotOOBE.Run -eq 'Shutdown') {$RunComboBox.SelectedValue = 'Shutdown Computer'}
if ($Global:AutopilotOOBE.Run -eq 'Sysprep') {$RunComboBox.SelectedValue = 'Sysprep /oobe /quit'}
if ($Global:AutopilotOOBE.Run -eq 'SysprepReboot') {$RunComboBox.SelectedValue = 'Sysprep /oobe /reboot'}
if ($Global:AutopilotOOBE.Run -eq 'SysprepShutdown') {$RunComboBox.SelectedValue = 'Sysprep /oobe /shutdown'}
if ($Global:AutopilotOOBE.Run -eq 'MDMDiag') {$RunComboBox.SelectedValue = 'MDMDiagnosticsTool -out C:\Temp'}
if ($Global:AutopilotOOBE.Run -eq 'MDMDiagAutopilot') {$RunComboBox.SelectedValue = 'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab'}
if ($Global:AutopilotOOBE.Run -eq 'MDMDiagAutopilotTPM') {$RunComboBox.SelectedValue = 'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\Autopilot.cab'}

$RunButton.add_Click( {
    if ($RunComboBox.SelectedValue -eq 'Command Prompt') {Start-Process Cmd.exe}
    if ($RunComboBox.SelectedValue -eq 'PowerShell') {Start-Process PowerShell.exe -ArgumentList "-Nologo"}
    if ($RunComboBox.SelectedValue -eq 'PowerShell ISE') {Start-Process PowerShell_ISE.exe}
    if ($RunComboBox.SelectedValue -eq 'Open Windows Explorer') {Start-Process Explorer.exe}
    if ($RunComboBox.SelectedValue -eq 'Open Windows Settings') {Start-Process ms-settings:}
    if ($RunComboBox.SelectedValue -eq 'Open Network and Wireless Settings') {Start-Process ms-availablenetworks:}
    if ($RunComboBox.SelectedValue -eq 'Restart Computer') {Restart-Computer}
    if ($RunComboBox.SelectedValue -eq 'Shutdown Computer') {Stop-Computer}
    if ($RunComboBox.SelectedValue -eq 'Sysprep /oobe /quit') {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/quit"}
    if ($RunComboBox.SelectedValue -eq 'Sysprep /oobe /reboot') {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/reboot"}
    if ($RunComboBox.SelectedValue -eq 'Sysprep /oobe /shutdown') {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/shutdown"}
    if ($RunComboBox.SelectedValue -eq 'MDMDiagnosticsTool -out C:\Temp') {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-out C:\Temp"}
    if ($RunComboBox.SelectedValue -eq 'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab') {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-area Autopilot","-cab C:\Temp\Autopilot.cab"}
    if ($RunComboBox.SelectedValue -eq 'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\Autopilot.cab') {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-area Autopilot;TPM","-cab C:\Temp\Autopilot.cab"}
})
#=======================================================================
#   Parameter Start
#=======================================================================
$StartTextBox.Text = $Global:AutopilotOOBE.Start

$StartButton.add_Click( {
    Write-Host -ForegroundColor Cyan "Run: $($StartTextBox.Text)"
    try {
        Start-Process $StartTextBox.Text
    }
    catch {
        Write-Warning "Could not execute $($StartTextBox.Text)"
    }
})
#=======================================================================
#   RegisterButton
#=======================================================================
$RegisterButton.add_Click( {
    $xamGUI.Close()
    Show-PowershellWindow

    Write-Host -ForegroundColor Cyan "Online: $true"
    $Params = @{
        Online = $true
    }

    if ($AssignCheckbox.IsChecked) {
        Write-Host -ForegroundColor Cyan "Assign: $true" 
        $Params.Assign = $true
    }

    if ($AddToGroupTextBox.Text -gt 0) {
        $Params.AddToGroup = $AddToGroupTextBox.Text
        Write-Host -ForegroundColor Cyan "AddToGroup: $($Params.AddToGroup)" 
    }

    if ($GroupTagTextBox.Text -gt 0) {
        $Params.GroupTag = $GroupTagTextBox.Text
        Write-Host -ForegroundColor Cyan "GroupTag: $($Params.GroupTag)" 
    }

    if ($AssignedUserTextBox.Text -gt 0) {
        $Params.AssignedUser = $AssignedUserTextBox.Text
        Write-Host -ForegroundColor Cyan "AssignedUser: $($Params.AssignedUser)" 
    }

    if (($AssignedComputerNameTextBox.Text -gt 0) -and ($AssignedComputerNameTextBox.Text -notmatch 'Azure AD Join Only')) {
        $Params.AssignedComputerName = $AssignedComputerNameTextBox.Text
        Write-Host -ForegroundColor Cyan "AssignedComputerName: $($Params.AssignedComputerName)" 
    }

    Write-Host -ForegroundColor Cyan "Install-Script Get-WindowsAutoPilotInfo"
    Start-Sleep -Seconds 3
    Install-Script Get-WindowsAutoPilotInfo -Force

    Write-Host -ForegroundColor Cyan "Get-WindowsAutoPilotInfo"
    Start-Sleep -Seconds 3

    Get-WindowsAutoPilotInfo @Params
    Start-Sleep -Seconds 3

    if ($PostActionComboBox.SelectedValue -match 'Sysprep') {
        Start-Sleep -Seconds 5
    }
    if ($PostActionComboBox.SelectedValue -match 'quit') {
        Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/quit" -Wait
    }
    if ($PostActionComboBox.SelectedValue -match 'reboot') {
        Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/reboot" -Wait
    }
    if ($PostActionComboBox.SelectedValue -match 'shutdown') {
        Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/shutdown" -Wait
    }
    if ($PostActionComboBox.SelectedValue -notmatch 'Sysprep') {
        & "$($MyInvocation.MyCommand.Module.ModuleBase)\AutopilotOOBE.ps1"
    }
})
#=======================================================================
#   ShowDialog
#=======================================================================
$xamGUI.ShowDialog() | Out-Null
#=======================================================================