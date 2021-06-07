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
[System.Reflection.Assembly]::LoadFrom("$Global:MyScriptDir\assembly\MahApps.Metro.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$Global:MyScriptDir\assembly\System.Windows.Interactivity.dll") | Out-Null

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
        xmlns:iconPacks="http://metro.mahapps.com/winfx/xaml/iconpacks"
        xmlns:Controls = "clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"

        Title = ""
        BorderBrush = "{DynamicResource AccentColorBrush}"
        BorderThickness = "2"
        Width = "980"
        Height = "670"
        Background = "#004275"
        
        ResizeMode = "CanResizeWithGrip"
        WindowStartupLocation = "CenterScreen"
        WindowStyle = "None">

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
        <StackPanel>
            <Label Name = "SidebarModuleVersion"
            Content = ""
            FontFamily = "Segoe UI" FontSize = "11"
            Foreground = "White"
            HorizontalAlignment = "Left"
            Margin = "0,0,0,0"
            />
        </StackPanel>
        <StackPanel>
            <Label Name = "SidebarManufacturer"
            Content = ""
            FontFamily = "Segoe UI" FontSize = "11"
            Foreground = "White"
            HorizontalAlignment = "Left"
            Margin = "0,15,0,0"
            />
        </StackPanel>
        <StackPanel>
            <Label Name = "SidebarModel"
            Content = ""
            FontFamily = "Segoe UI" FontSize = "11"
            Foreground = "White"
            HorizontalAlignment = "Left"
            Margin = "0,30,0,0"
            />
        </StackPanel>
        <StackPanel>
            <Label Name = "SidebarSerialNumber"
            Content = ""
            FontFamily = "Segoe UI" FontSize = "11"
            Foreground = "White"
            HorizontalAlignment = "Left"
            Margin = "0,45,0,0"
            />
        </StackPanel>
        <StackPanel>
            <Label Name = "SidebarBiosVersion"
            Content = ""
            FontFamily = "Segoe UI" FontSize = "11"
            Foreground = "White"
            HorizontalAlignment = "Left"
            Margin = "0,60,0,0"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Label Name = "TitleMain"
            Content = "Join Autopilot OOBE"
            FontFamily = "Segoe UI Light" FontSize = "46"
            Foreground = "White"
            HorizontalAlignment = "Center"
            Margin = "0,0,0,0"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Label Name = "GroupTagLabel"
            Content = "GroupTag:"
            FontFamily = "Segoe UI" FontSize = "15"
            Foreground = "White"
            HorizontalAlignment = "Stretch"
            Margin = "0,100,0,0"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <TextBox Name = "GroupTagTextBox"
            Background = "#002846"
            BorderThickness = "2"
            FontFamily = "Segoe UI"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "White"
            Height = "40"
            HorizontalAlignment = "Stretch"
            Margin = "180,90,0,0"
            Width = "410"
            Padding = "8"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Label Name = "AddToGroupLabel"
            Content = "AddToGroup:"
            FontFamily = "Segoe UI"
            FontSize = "15"
            Foreground = "White"
            HorizontalAlignment = "Stretch"
            Margin = "0,150,0,0"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <TextBox Name = "AddToGroupTextBox"
            Background = "#002846"
            BorderThickness = "2"
            FontFamily = "Segoe UI"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "White"
            Height = "40"
            HorizontalAlignment = "Stretch"
            Margin = "180,140,0,0"
            Width = "410"
            Padding = "8"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Label Name = "AssignedUserLabel"
            Content = "AssignedUser:"
            FontFamily = "Segoe UI"
            FontSize = "15"
            Foreground = "White"
            HorizontalAlignment = "Stretch"
            Margin = "0,200,0,0"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <TextBox Name = "AssignedUserTextBox"
            Background = "#002846"
            BorderThickness = "2"
            FontFamily = "Segoe UI"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "White"
            Height = "40"
            HorizontalAlignment = "Stretch"
            Margin = "180,190,0,0"
            Width = "410"
            Padding = "8"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Label Name = "AssignedComputerNameLabel"
            Content = "AssignedComputerName:"
            FontFamily = "Segoe UI"
            FontSize = "15"
            Foreground = "White"
            HorizontalAlignment = "Stretch"
            Margin = "0,250,0,0"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <TextBox Name = "AssignedComputerNameTextBox"
            Background = "#002846"
            BorderThickness = "2"
            FontFamily = "Segoe UI"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "White"
            Height = "40"
            HorizontalAlignment = "Stretch"
            Margin = "180,240,0,0"
            Text = "Azure AD Join Only"
            Width = "410"
            Padding = "8"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Label Name = "PostActionLabel"
            Content = "PostAction:"
            FontFamily = "Segoe UI"
            FontSize = "15"
            Foreground = "White"
            HorizontalAlignment = "Stretch"
            Margin = "0,300,0,0"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <ComboBox Name = "PostActionComboBox"
            Background = "#002846"
            BorderThickness = "2"
            FontFamily = "Segoe UI"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "Black"
            Height = "40"
            HorizontalAlignment = "Stretch"
            Margin = "180,290,0,0"
            Width = "410"
            Padding = "8"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <CheckBox Name = "AssignCheckbox"
            HorizontalAlignment = "Stretch"
            Background = "#002846"
            BorderThickness = "2"
            FontFamily = "Segoe UI"
            FontSize = "15"
            Foreground = "White"
            Margin = "5,350,0,0"
            >Assign: Wait for Intune to assign an Autopilot profile for this device
            </CheckBox>
        </StackPanel>

        <StackPanel Width = "600">
            <Button Name = "RegisterButton"
            Content = "Register"
            FontFamily = "Segoe UI"
            FontSize = "15"
            Height = "40"
            HorizontalAlignment = "Left"
            Margin = "0,390,0,0"
            Width = "170"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Button Name = "RunButton"
            Content = "Run"
            FontSize = "15"
            FontWeight = "Normal"
            Height = "40"
            HorizontalAlignment = "Left"
            Margin = "0,520,0,0"
            Width = "100"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <ComboBox Name = "RunComboBox"
            Background = "#002846"
            BorderThickness = "2"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "Black"
            Height = "40"
            HorizontalAlignment = "Left"
            Margin = "100,520,0,0"
            Padding = "8"
            Width = "500"
            />
        </StackPanel>

        <StackPanel Width = "600">
            <Button Name = "DocsButton"
            Content = "Docs"
            FontSize = "15"
            FontWeight = "Normal"
            Height = "40"
            HorizontalAlignment = "Left"
            Margin = "0,570,0,0"
            Width = "100"
            />
        </StackPanel>
        <StackPanel Width = "600">
            <TextBox Name = "DocsTextBox"
            Background = "#002846"
            BorderThickness = "2"
            FontSize = "15"
            FontWeight = "Normal"
            Foreground = "White"
            Height = "40"
            HorizontalAlignment = "Left"
            Margin = "100,570,0,0"
            Padding = "8"
            Width = "500"
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
#   Sidebar
#=======================================================================
$ModuleVersion = (Get-Module -Name AutopilotOOBE | Sort-Object Version | Select-Object Version -Last 1).Version
$SidebarModuleVersion.Content = "Version: $ModuleVersion"

$SidebarManufacturer.Content = ((Get-CimInstance -ClassName CIM_ComputerSystem).Manufacturer).Trim()

if ($SidebarManufacturer.Content -match 'Lenovo') {
    $SidebarModel.Content = ((Get-CimInstance -ClassName Win32_ComputerSystemProduct).Version).Trim()
}
else {
    $SidebarModel.Content = ((Get-CimInstance -ClassName CIM_ComputerSystem).Model).Trim()
}

$SerialNumber = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
$SidebarSerialNumber.Content = "Serial Number: $SerialNumber"

$BiosVersion = ((Get-CimInstance -ClassName Win32_BIOS).SMBIOSBIOSVersion).Trim()
$SidebarBiosVersion.Content = "BIOS Version: $BiosVersion"
#=======================================================================
#   Parameters
#=======================================================================
$AutopilotOOBEParams = (Get-Command Start-AutopilotOOBE).Parameters

$AddToGroupTextBox.Text = $Global:AutopilotOOBE.AddToGroup
$GroupTagTextBox.Text = $Global:AutopilotOOBE.GroupTag
$TitleMain.Content = $Global:AutopilotOOBE.Title
#=======================================================================
#   Parameter PostAction
#=======================================================================
#$AutopilotOOBEParams["PostAction"].Attributes.ValidValues | ForEach-Object {
#    $PostActionComboBox.Items.Add($_) | Out-Null
#}
#$PostActionComboBox.SelectedValue = $Global:AutopilotOOBE.PostAction
$PostActionComboBox.Items.Add('None') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /quit') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /reboot') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /shutdown') | Out-Null

if ($Global:AutopilotOOBE.PostAction -eq 'None') {$PostActionComboBox.SelectedValue = 'None'}
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
#   Parameter Docs
#=======================================================================
$DocsTextBox.Text = $Global:AutopilotOOBE.Docs

$DocsButton.add_Click( {
    Write-Host -ForegroundColor Cyan "Run: $($DocsTextBox.Text)"
    try {
        Start-Process $DocsTextBox.Text
    }
    catch {
        Write-Warning "Could not execute $($DocsTextBox.Text)"
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