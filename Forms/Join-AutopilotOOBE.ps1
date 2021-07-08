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
    Param(
     [Parameter(Mandatory = $False, Position = 1)]
     [string]$XamlPath
    )

    # Import the XAML code
    [xml]$Global:xmlWPF = Get-Content -Path $XamlPath

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
LoadForm -XamlPath (Join-Path $Global:MyScriptDir 'Join-AutopilotOOBE.xaml')

$MDMEventLog = @'
<ViewerConfig>
	<QueryConfig>
		<QueryParams>
			<UserQuery/>
		</QueryParams>
		<QueryNode>
			<Name LanguageNeutralValue="MDMDiagnosticsTool">MDMDiagnosticsTool</Name>
			<Description>MDMDiagnosticsTool</Description>
			<SuppressQueryExecutionErrors>1</SuppressQueryExecutionErrors>
			<QueryList>
				<Query>
					<Select Path="Microsoft-Windows-AAD/Operational">*</Select>
					<Select Path="Microsoft-Windows-AppXDeployment-Server/Operational">*</Select>
					<Select Path="Microsoft-Windows-AssignedAccess/Admin">*</Select>
					<Select Path="Microsoft-Windows-AssignedAccess/Operational">*</Select>
					<Select Path="Microsoft-Windows-AssignedAccessBroker/Admin">*</Select>
					<Select Path="Microsoft-Windows-AssignedAccessBroker/Operational">*</Select>
					<Select Path="Microsoft-Windows-Crypto-NCrypt/Operational">*</Select>
					<Select Path="Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin">*</Select>
					<Select Path="Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational">*</Select>
					<Select Path="Microsoft-Windows-ModernDeployment-Diagnostics-Provider/Autopilot">*</Select>
					<Select Path="Microsoft-Windows-ModernDeployment-Diagnostics-Provider/ManagementService">*</Select>
					<Select Path="Microsoft-Windows-Provisioning-Diagnostics-Provider/Admin">*</Select>
					<Select Path="Microsoft-Windows-Shell-Core/Operational">*</Select>
					<Select Path="Microsoft-Windows-User Device Registration/Admin">*</Select>
				</Query>
			</QueryList>
		</QueryNode>
	</QueryConfig>
</ViewerConfig>
'@
#=======================================================================
#   Sidebar
#=======================================================================
$ModuleVersion = (Get-Module -Name AutopilotOOBE | Sort-Object Version | Select-Object Version -Last 1).Version
$SidebarModuleVersion.Content = "ver $ModuleVersion"

try {
    $Tpm = (Get-CimInstance -Namespace "root\CIMV2\Security\MicrosoftTPM" -ClassName Win32_Tpm).SpecVersion
}
catch {}

if ($Tpm -match '2.0') {
    $SidebarTpmVersion.Content = "TPM: 2.0"
    $SidebarTpmVersion.Background = "Green"
}
elseif ($Tpm -match '1.2') {
    $SidebarTpmVersion.Content = "TPM: 1.2"
    $SidebarTpmVersion.Background = "Red"
}
else {
    $SidebarTpmVersion.Visibility = "Collapsed"
}

$SidebarManufacturer.Content = ((Get-CimInstance -ClassName CIM_ComputerSystem).Manufacturer).Trim()

if ($SidebarManufacturer.Content -match 'Lenovo') {
    $SidebarModel.Content = ((Get-CimInstance -ClassName Win32_ComputerSystemProduct).Version).Trim()
}
else {
    $SidebarModel.Content = ((Get-CimInstance -ClassName CIM_ComputerSystem).Model).Trim()
}

$SerialNumber = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
$SidebarSerialNumber.Content = $SerialNumber

$BiosVersion = ((Get-CimInstance -ClassName Win32_BIOS).SMBIOSBIOSVersion).Trim()
$SidebarBiosVersion.Content = "BIOS $BiosVersion"
#=======================================================================
#   Parameters
#=======================================================================
$AutopilotOOBEParams = (Get-Command Start-AutopilotOOBE).Parameters
#=======================================================================
#   Parameter Title
#=======================================================================
$TitleMain.Content = $Global:AutopilotOOBE.Title
#=======================================================================
#   Windows Version Info
#=======================================================================
$Global:GetRegCurrentVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

$SubTitleProductName = ($Global:GetRegCurrentVersion).ProductName

if ($Global:GetRegCurrentVersion.DisplayVersion -gt 0) {
    $SubTitleDisplayVersion = ($Global:GetRegCurrentVersion).DisplayVersion
}
else {
    $SubTitleDisplayVersion = ($Global:GetRegCurrentVersion).ReleaseId
}

$SubTitleBuildNumber = "$($Global:GetRegCurrentVersion.CurrentBuild).$($Global:GetRegCurrentVersion.UBR)"

$TitleMinor.Content = "$SubTitleProductName $SubTitleDisplayVersion ($SubTitleBuildNumber)"
#=======================================================================
#   Parameter GroupTag
#=======================================================================
$Global:AutopilotOOBE.GroupTagOptions | ForEach-Object {
    $GroupTagComboBox.Items.Add($_) | Out-Null
}

if ($Global:AutopilotOOBE.GroupTag) {
    $GroupTagComboBox.Text = $Global:AutopilotOOBE.GroupTag
}

if ($Disabled -contains 'GroupTag') {$GroupTagComboBox.IsEnabled = $false}

if ($Hidden -contains 'GroupTag') {
    $StackPanelGroupTag = $Global:xamGUI.FindName("StackPanelGroupTag")
    $StackPanelGroupTag.Visibility = 'Collapsed'
}
#=======================================================================
#   Parameter AddToGroup
#=======================================================================
$Global:AutopilotOOBE.AddToGroupOptions | ForEach-Object {
    $AddToGroupComboBox.Items.Add($_) | Out-Null
}

if ($Global:AutopilotOOBE.AddToGroup) {
    $AddToGroupComboBox.Text = $Global:AutopilotOOBE.AddToGroup
}

if ($Disabled -contains 'AddToGroup') {$AddToGroupComboBox.IsEnabled = $false}

if ($Hidden -contains 'AddToGroup') {
    $StackPanelAddToGroup = $Global:xamGUI.FindName("StackPanelAddToGroup")
    $StackPanelAddToGroup.Visibility = 'Collapsed'
}
#=======================================================================
#   Parameter AssignedUser
#=======================================================================
$AssignedUserTextBox.Text = $Global:AutopilotOOBE.AssignedUserExample
if ($Global:AutopilotOOBE.AssignedUser -gt 0) {$AssignedUserTextBox.Text = $Global:AutopilotOOBE.AssignedUser}

if ($Disabled -contains 'AssignedUser') {$AssignedUserTextBox.IsEnabled = $false}

if ($Hidden -contains 'AssignedUser') {
    $StackPanelAssignedUser = $Global:xamGUI.FindName("StackPanelAssignedUser")
    $StackPanelAssignedUser.Visibility = 'Collapsed'
}
#=======================================================================
#   Parameter AssignedComputerName
#=======================================================================
$AssignedComputerNameTextBox.Text = $Global:AutopilotOOBE.AssignedComputerNameExample
if ($Global:AutopilotOOBE.AssignedComputerName -gt 0) {$AssignedComputerNameTextBox.Text = $Global:AutopilotOOBE.AssignedComputerName}

if ($Disabled -contains 'AssignedComputerName') {$AssignedComputerNameTextBox.IsEnabled = $false}

if ($Hidden -contains 'AssignedComputerName') {
    $StackPanelAssignedComputerName = $Global:xamGUI.FindName("StackPanelAssignedComputerName")
    $StackPanelAssignedComputerName.Visibility = 'Collapsed'
}
#=======================================================================
#   Parameter PostAction
#=======================================================================
#$AutopilotOOBEParams["PostAction"].Attributes.ValidValues | ForEach-Object {
#    $PostActionComboBox.Items.Add($_) | Out-Null
#}
#$PostActionComboBox.SelectedValue = $Global:AutopilotOOBE.PostAction
$PostActionComboBox.Items.Add('Quit') | Out-Null
$PostActionComboBox.Items.Add('Restart Computer') | Out-Null
$PostActionComboBox.Items.Add('Shutdown Computer') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /quit') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /reboot') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /oobe /shutdown') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /generalize /oobe /reboot') | Out-Null
$PostActionComboBox.Items.Add('Sysprep /generalize /oobe /shutdown') | Out-Null

if ($Global:AutopilotOOBE.PostAction -eq 'Quit') {$PostActionComboBox.SelectedValue = 'Quit'}
if ($Global:AutopilotOOBE.PostAction -eq 'Restart') {$PostActionComboBox.SelectedValue = 'Restart Computer'}
if ($Global:AutopilotOOBE.PostAction -eq 'Shutdown') {$PostActionComboBox.SelectedValue = 'Shutdown Computer'}
if ($Global:AutopilotOOBE.PostAction -eq 'Sysprep') {$PostActionComboBox.SelectedValue = 'Sysprep /oobe /quit'}
if ($Global:AutopilotOOBE.PostAction -eq 'SysprepReboot') {$PostActionComboBox.SelectedValue = 'Sysprep /oobe /reboot'}
if ($Global:AutopilotOOBE.PostAction -eq 'SysprepShutdown') {$PostActionComboBox.SelectedValue = 'Sysprep /oobe /shutdown'}
if ($Global:AutopilotOOBE.PostAction -eq 'GeneralizeReboot') {$PostActionComboBox.SelectedValue = 'Sysprep /generalize /oobe /reboot'}
if ($Global:AutopilotOOBE.PostAction -eq 'GeneralizeShutdown') {$PostActionComboBox.SelectedValue = 'Sysprep /generalize /oobe /shutdown'}

if ($Disabled -contains 'PostAction') {$PostActionComboBox.IsEnabled = $false}

if ($Hidden -contains 'PostAction') {
    $StackPanelPostAction = $Global:xamGUI.FindName("StackPanelPostAction")
    $StackPanelPostAction.Visibility = 'Collapsed'
}
#=======================================================================
#   Parameter Assign
#=======================================================================
if ($Global:AutopilotOOBE.Assign -eq $true) {$AssignCheckBox.IsChecked = $true}

if ($Disabled -contains 'Assign') {$AssignCheckBox.IsEnabled = $false}

if ($Hidden -contains 'Assign') {
    $StackPanelAssign = $Global:xamGUI.FindName("StackPanelAssign")
    $StackPanelAssign.Visibility = 'Collapsed'
}
#=======================================================================
#   Register
#=======================================================================
if ($Hidden -contains 'Register') {
    $StackPanelRegister = $Global:xamGUI.FindName("StackPanelRegister")
    $StackPanelRegister.Visibility = 'Collapsed'

    if ($Global:RegAutoPilot.CloudAssignedForcedEnrollment -eq 1) {
        $CloudAssignedForcedEnrollment = 'Yes'
    }
    else {
        $CloudAssignedForcedEnrollment = 'No'
    }

    if ($Global:RegAutoPilot.IsDevicePersonalized -eq 1) {
        $IsDevicePersonalized = 'Yes'
    }
    else {
        $IsDevicePersonalized = 'No'
    }

    if ($Global:RegAutoPilot.CloudAssignedLanguage) {
        $CloudAssignedLanguage = $Global:RegAutoPilot.CloudAssignedLanguage
    }
    else {
        $CloudAssignedLanguage = 'Operating System Default'
    }
    
    if (($Global:RegAutoPilot.CloudAssignedOobeConfig -band 512) -gt 0) {$PatchDownload = 'Yes'} else {$PatchDownload = 'No'}
    if (($Global:RegAutoPilot.CloudAssignedOobeConfig -band 128) -gt 0) {$TPMRequired = 'Yes'} else {$TPMRequired = 'No'}
    if (($Global:RegAutoPilot.CloudAssignedOobeConfig -band 64) -gt 0) {$DeviceAuth = 'Yes'} else {$DeviceAuth = 'No'}
    if (($Global:RegAutoPilot.CloudAssignedOobeConfig -band 32) -gt 0) {$TPMAttestation = 'Yes'} else {$TPMAttestation = 'No'}
    if (($Global:RegAutoPilot.CloudAssignedOobeConfig -band 4) -gt 0) {$SkipExpress = 'Yes'} else {$SkipExpress = 'No'}
    if (($Global:RegAutoPilot.CloudAssignedOobeConfig -band 2) -gt 0) {$DisallowAdmin = 'Yes'} else {$DisallowAdmin = 'No'}

    $InformationLabel.Content = @"
    Azure AD Tenant: $($Global:RegAutoPilot.CloudAssignedTenantDomain)
    Azure AD Tenant ID: $($Global:RegAutoPilot.CloudAssignedTenantId)
    MDM ID: $($Global:RegAutoPilot.CloudAssignedMdmId)
    Autopilot Service Correlation ID: $($Global:RegAutoPilot.AutopilotServiceCorrelationId)

    AAD Device Auth: $DeviceAuth
    AAD TPM Required: $TPMRequired
    Disallow Admin: $DisallowAdmin
    Enable Patch Download: $PatchDownload
    Forced Enrollment: $CloudAssignedForcedEnrollment
    Is Device Personalized: $IsDevicePersonalized
    Language: $CloudAssignedLanguage
    Skip Express Settings: $SkipExpress
    Telemetry Level: $($Global:RegAutoPilot.CloudAssignedTelemetryLevel)
    TPM Attestation: $TPMAttestation
"@
}
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
$RunComboBox.Items.Add('Sysprep /audit /reboot') | Out-Null
$RunComboBox.Items.Add('Event Viewer') | Out-Null
$RunComboBox.Items.Add('Enable Windows Security Appx') | Out-Null
$RunComboBox.Items.Add('Get-TPM') | Out-Null
$RunComboBox.Items.Add('Clear-TPM') | Out-Null
$RunComboBox.Items.Add('Initialize-TPM') | Out-Null
$RunComboBox.Items.Add('Get-AutopilotDiagnostics') | Out-Null
$RunComboBox.Items.Add('MDMDiagnosticsTool -out C:\Temp') | Out-Null
$RunComboBox.Items.Add('MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab') | Out-Null
$RunComboBox.Items.Add('MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\Autopilot.cab') | Out-Null
$RunComboBox.Items.Add('Update-MyDellBios') | Out-Null

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
if ($Global:AutopilotOOBE.Run -eq 'SysprepAudit') {$RunComboBox.SelectedValue = 'Sysprep /audit /reboot'}
if ($Global:AutopilotOOBE.Run -eq 'EventViewer') {$RunComboBox.SelectedValue = 'Event Viewer'}
if ($Global:AutopilotOOBE.Run -eq 'AutopilotDiagnostics') {$RunComboBox.SelectedValue = 'Get-AutopilotDiagnostics'}
if ($Global:AutopilotOOBE.Run -eq 'MDMDiag') {$RunComboBox.SelectedValue = 'MDMDiagnosticsTool -out C:\Temp'}
if ($Global:AutopilotOOBE.Run -eq 'MDMDiagAutopilot') {$RunComboBox.SelectedValue = 'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab'}
if ($Global:AutopilotOOBE.Run -eq 'MDMDiagAutopilotTPM') {$RunComboBox.SelectedValue = 'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\Autopilot.cab'}
if ($Global:AutopilotOOBE.Run -eq 'UpdateMyDellBios') {$RunComboBox.SelectedValue = 'Update-MyDellBios'}

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
    if ($RunComboBox.SelectedValue -eq 'Sysprep /audit /reboot') {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/audit", "/reboot"}
    if ($RunComboBox.SelectedValue -eq 'Event Viewer') {
        $MDMEventLog | Set-Content -Path "$env:ProgramData\Microsoft\Event Viewer\Views\MDMDiagnosticsTool.xml"
        Restart-Service -Name EventLog -Force -ErrorAction Ignore
        Show-EventLog
    }
    if ($RunComboBox.SelectedValue -eq 'Enable Windows Security Appx') {
        Start-Process PowerShell.exe -ArgumentList "Add-AppxPackage -Register -DisableDevelopmentMode 'C:\Windows\SystemApps\Microsoft.Windows.SecHealthUI_cw5n1h2txyewy\AppXManifest.xml';start windowsdefender:"
    }
    if ($RunComboBox.SelectedValue -eq 'Get-TPM') {
        Start-Process PowerShell.exe -ArgumentList "-NoExit -Command Get-TPM"
    }
    if ($RunComboBox.SelectedValue -eq 'Clear-TPM') {
        Start-Process PowerShell.exe -ArgumentList "-NoExit -Command Clear-TPM;Write-Warning 'You should restart the computer at this time'"
    }
    if ($RunComboBox.SelectedValue -eq 'Initialize-TPM') {
        Start-Process PowerShell.exe -ArgumentList "-NoExit -Command Initialize-Tpm -AllowClear -AllowPhysicalPresence;Write-Warning 'You should restart the computer at this time'"
    }
    if ($RunComboBox.SelectedValue -eq 'Get-AutopilotDiagnostics') {
        Show-PowershellWindow
        Install-Script Get-AutopilotDiagnostics -Force -Verbose
        Get-AutopilotDiagnostics -Online
        Pause
    }
    if ($RunComboBox.SelectedValue -eq 'MDMDiagnosticsTool -out C:\Temp') {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-out C:\Temp"}
    if ($RunComboBox.SelectedValue -eq 'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab') {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-area Autopilot","-cab C:\Temp\Autopilot.cab"}
    if ($RunComboBox.SelectedValue -eq 'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\Autopilot.cab') {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-area Autopilot;TPM","-cab C:\Temp\Autopilot.cab"}
    if ($RunComboBox.SelectedValue -eq 'Update-MyDellBios') {
        Show-PowershellWindow
        Install-Module OSD -Force
        Start-Sleep -Seconds 2
        Get-MyDellBios
        Update-MyDellBios
    }
})

if ($Hidden -contains 'Run') {
    $StackPanelRun = $Global:xamGUI.FindName("StackPanelRun")
    $StackPanelRun.Visibility = 'Collapsed'
}
#=======================================================================
#   Parameter Docs
#=======================================================================
$DocsComboBox.Items.Add('Windows Autopilot Documentation') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot Overview') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot User-Driven Mode') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot for Pre-Provisioned Deployment') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot Deployment for Existing Devices') | Out-Null
$DocsComboBox.Items.Add('Manually register devices with Windows Autopilot') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot Troubleshooting Overview') | Out-Null
$DocsComboBox.Items.Add('Troubleshoot Autopilot Device Import and Enrollment') | Out-Null
$DocsComboBox.Items.Add('Troubleshoot Autopilot OOBE Issues') | Out-Null
$DocsComboBox.Items.Add('Troubleshoot Azure Active Directory Join Issues') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot Known Issues') | Out-Null
$DocsComboBox.Items.Add('Windows Autopilot Resolved Issues') | Out-Null
$DocsComboBox.Items.Add('Sysprep Overview') | Out-Null
$DocsComboBox.Items.Add('Sysprep Audit Mode Overview') | Out-Null
$DocsComboBox.Items.Add('Sysprep Command-Line Options') | Out-Null


if ($Hidden -contains 'Register') {
    $DocsComboBox.SelectedValue = 'Troubleshoot Autopilot OOBE Issues'
}
else {
    $DocsComboBox.SelectedValue = 'Windows Autopilot Documentation'
}


if ($Global:AutopilotOOBE.Docs) {
    $DocsComboBox.Items.Add($Global:AutopilotOOBE.Docs) | Out-Null
    $DocsComboBox.SelectedValue = $Global:AutopilotOOBE.Docs
}

$DocsButton.add_Click( {
    Write-Host -ForegroundColor Cyan "Run: $($DocsComboBox.SelectedValue)"

    if ($DocsComboBox.SelectedValue -eq 'Windows Autopilot Documentation') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot Overview') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/windows-autopilot'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot User-Driven Mode') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/user-driven'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot for Pre-Provisioned Deployment') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/pre-provision'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot Deployment for Existing Devices') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/existing-devices'}
    elseif ($DocsComboBox.SelectedValue -eq 'Manually register devices with Windows Autopilot') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/add-devices'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot Troubleshooting Overview') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshooting'}
    elseif ($DocsComboBox.SelectedValue -eq 'Troubleshoot Autopilot Device Import and Enrollment') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshoot-device-enrollment'}
    elseif ($DocsComboBox.SelectedValue -eq 'Troubleshoot Autopilot OOBE Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshoot-oobe'}
    elseif ($DocsComboBox.SelectedValue -eq 'Troubleshoot Azure Active Directory Join Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshoot-aad-join'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot Known Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/known-issues'}
    elseif ($DocsComboBox.SelectedValue -eq 'Windows Autopilot Resolved Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/resolved-issues'}
    elseif ($DocsComboBox.SelectedValue -eq 'Sysprep Overview') {Start-Process 'https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/sysprep--system-preparation--overview'}
    elseif ($DocsComboBox.SelectedValue -eq 'Sysprep Audit Mode Overview') {Start-Process 'https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/audit-mode-overview'}
    elseif ($DocsComboBox.SelectedValue -eq 'Sysprep Command-Line Options') {Start-Process 'https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/sysprep-command-line-options'}
    else {
        try {
            Start-Process $DocsComboBox.SelectedValue
        }
        catch {
            Write-Warning "Could not execute $($DocsComboBox.SelectedValue)"
        }
    }
})

if ($Hidden -contains 'Docs') {
    $StackPanelDocs = $Global:xamGUI.FindName("StackPanelDocs")
    $StackPanelDocs.Visibility = 'Collapsed'
}
#=======================================================================
#   RegisterButton
#=======================================================================
$RegisterButton.add_Click( {
    $xamGUI.Close()
    Show-PowershellWindow

    $Params = @{
        Online = $true
    }

    if ($AssignCheckbox.IsChecked) {
        $Params.Assign = $true
    }

    if ($AddToGroupComboBox.Text -gt 0) {
        $Params.AddToGroup = $AddToGroupComboBox.Text
    }

    if ($GroupTagComboBox.Text -gt 0) {
        $Params.GroupTag = $GroupTagComboBox.Text
    }

    if (($AssignedUserTextBox.Text -gt 0) -and ($AssignedUserTextBox.Text -notmatch $Global:AutopilotOOBE.AssignedUserExample)) {
        $Params.AssignedUser = $AssignedUserTextBox.Text
    }

    if (($AssignedComputerNameTextBox.Text -gt 0) -and ($AssignedComputerNameTextBox.Text -notmatch $Global:AutopilotOOBE.AssignedComputerNameExample)) {
        $Params.AssignedComputerName = $AssignedComputerNameTextBox.Text
    }

    Write-Host -ForegroundColor Cyan "Install-Script Get-WindowsAutoPilotInfo"
    if ($Global:AutopilotOOBE.Demo -ne $true) {
        Start-Sleep -Seconds 3
        Install-Script Get-WindowsAutoPilotInfo -Force -Verbose
    }

    Write-Host ($Params | Out-String)
    Write-Host -ForegroundColor Cyan "Get-WindowsAutoPilotInfo @Params"

    if ($Global:AutopilotOOBE.Demo -ne $true) {
        Start-Sleep -Seconds 3
        Get-WindowsAutoPilotInfo @Params
    }


    if ($PostActionComboBox.SelectedValue -eq 'Restart Computer') {Restart-Computer}
    if ($PostActionComboBox.SelectedValue -eq 'Shutdown Computer') {Stop-Computer}

    if ($PostActionComboBox.SelectedValue -match 'Sysprep') {
        Write-Host -ForegroundColor Cyan "Executing Sysprep"

        if ($Global:AutopilotOOBE.Demo -ne $true) {
            if ($PostActionComboBox.SelectedValue -match 'quit') {
                Start-Sleep -Seconds 3
                Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/quit" -Wait
            }
            
            elseif (($PostActionComboBox.SelectedValue -match 'reboot') -and ($PostActionComboBox.SelectedValue -match 'generalize')) {
                Start-Sleep -Seconds 3
                Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "generalize", "/oobe", "/reboot" -Wait
            }
            elseif (($PostActionComboBox.SelectedValue -match 'shutdown') -and ($PostActionComboBox.SelectedValue -match 'generalize')) {
                Start-Sleep -Seconds 3
                Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "generalize", "/oobe", "/reboot" -Wait
            }

            elseif ($PostActionComboBox.SelectedValue -match 'reboot') {
                Start-Sleep -Seconds 3
                Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/reboot" -Wait
            }
            elseif ($PostActionComboBox.SelectedValue -match 'shutdown') {
                Start-Sleep -Seconds 3
                Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/shutdown" -Wait
            }
        }
    }

<#     if ($PostActionComboBox.SelectedValue -match 'None') {
        & "$Global:MyScriptDir\Join-AutopilotOOBE.ps1"
    } #>
})
#=======================================================================
#   ShowDialog
#=======================================================================
$xamGUI.ShowDialog() | Out-Null
#=======================================================================