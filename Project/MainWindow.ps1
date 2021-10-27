# PoSHPF - Version 1.2
# Grab all resources (MahApps, etc), all XAML files, and any potential static resources
$Global:resources = Get-ChildItem -Path "$PSScriptRoot\Resources\*.dll" -ErrorAction SilentlyContinue
$Global:XAML = Get-ChildItem -Path "$PSScriptRoot\*.xaml" | Where-Object {$_.Name -ne 'App.xaml'} -ErrorAction SilentlyContinue #Changed path and exclude App.xaml
$Global:MediaResources = Get-ChildItem -Path "$PSScriptRoot\Media" -ErrorAction SilentlyContinue

# This class allows the synchronized hashtable to be available across threads,
# but also passes a couple of methods along with it to do GUI things via the
# object's dispatcher.
class SyncClass 
{
    #Hashtable containing all forms/windows and controls - automatically created when newing up
    [hashtable]$SyncHash = [hashtable]::Synchronized(@{}) 
    
    # method to close the window - pass window name
    [void]CloseWindow($windowName){ 
        $this.SyncHash.$windowName.Dispatcher.Invoke([action]{$this.SyncHash.$windowName.Close()},"Normal") 
    }
    
    # method to update GUI - pass object name, property and value   
    [void]UpdateElement($object,$property,$value){ 
        $this.SyncHash.$object.Dispatcher.Invoke([action]{ $this.SyncHash.$object.$property = $value },"Normal") 
    } 
}
$Global:SyncClass = [SyncClass]::new() # create a new instance of this SyncClass to use.

###################
## Import Resources
###################
# Load WPF Assembly
Add-Type -assemblyName PresentationFramework

# Load Resources
foreach($dll in $resources) { [System.Reflection.Assembly]::LoadFrom("$($dll.FullName)") | out-null }

##############
## Import XAML
##############
$xp = '[^a-zA-Z_0-9]' # All characters that are not a-Z, 0-9, or _
$vx = @()             # An array of XAML files loaded

foreach($x in $XAML) { 
    # Items from XAML that are known to cause issues
    # when PowerShell parses them.
    $xamlToRemove = @(
        'mc:Ignorable="d"',
        "x:Class=`"(.*?)`"",
        "xmlns:local=`"(.*?)`""
    )

    $xaml = Get-Content $x.FullName # Load XAML
    $xaml = $xaml -replace "x:N",'N' # Rename x:Name to just Name (for consumption in variables later)
    foreach($xtr in $xamlToRemove){ $xaml = $xaml -replace $xtr } # Remove items from $xamlToRemove
    
    # Create a new variable to store the XAML as XML
    New-Variable -Name "xaml$(($x.BaseName) -replace $xp, '_')" -Value ($xaml -as [xml]) -Force
    
    # Add XAML to list of XAML documents processed
    $vx += "$(($x.BaseName) -replace $xp, '_')"
}
#######################
## Add Media Resources
#######################
$imageFileTypes = @(".jpg",".bmp",".gif",".tif",".png") # Supported image filetypes
$avFileTypes = @(".mp3",".wav",".wmv") # Supported audio/visual filetypes
$xp = '[^a-zA-Z_0-9]' # All characters that are not a-Z, 0-9, or _
if($MediaResources.Count -gt 0){
    ## Okay... the following code is just silly. I know
    ## but hear me out. Adding the nodes to the elements
    ## directly caused big issues - mainly surrounding the
    ## "x:" namespace identifiers. This is a hacky fix but
    ## it does the trick.
    foreach($v in $vx)
    {
        $xml = ((Get-Variable -Name "xaml$($v)").Value) # Load the XML

        # add the resources needed for strings
        $xml.DocumentElement.SetAttribute("xmlns:sys","clr-namespace:System;assembly=System")

        # if the document doesn't already have a "Window.Resources" create it
        if($null -eq ($xml.DocumentElement.'Window.Resources')){ 
            $fragment = "<Window.Resources>" 
            $fragment += "<ResourceDictionary>"
        }
        
        # Add each StaticResource with the key of the base name and source to the full name
        foreach($sr in $MediaResources)
        {
            $srname = "$($sr.BaseName -replace $xp, '_')$($sr.Extension.Substring(1).ToUpper())" #convert name to basename + Uppercase Extension
            if($sr.Extension -in $imageFileTypes){ $fragment += "<BitmapImage x:Key=`"$srname`" UriSource=`"$($sr.FullName)`" />" }
            if($sr.Extension -in $avFileTypes){ 
                $uri = [System.Uri]::new($sr.FullName)
                $fragment += "<sys:Uri x:Key=`"$srname`">$uri</sys:Uri>" 
            }    
        }

        # if the document doesn't already have a "Window.Resources" close it
        if($null -eq ($xml.DocumentElement.'Window.Resources'))
        {
            $fragment += "</ResourceDictionary>"
            $fragment += "</Window.Resources>"
            $xml.DocumentElement.InnerXml = $fragment + $xml.DocumentElement.InnerXml
        }
        # otherwise just add the fragment to the existing resource dictionary
        else
        {
            $xml.DocumentElement.'Window.Resources'.ResourceDictionary.InnerXml += $fragment
        }

        # Reset the value of the variable
        (Get-Variable -Name "xaml$($v)").Value = $xml
    }
}
#################
## Create "Forms"
#################
$forms = @()
foreach($x in $vx)
{
    $Reader = (New-Object System.Xml.XmlNodeReader ((Get-Variable -Name "xaml$($x)").Value)) #load the xaml we created earlier into XmlNodeReader
    New-Variable -Name "form$($x)" -Value ([Windows.Markup.XamlReader]::Load($Reader)) -Force #load the xaml into XamlReader
    $forms += "form$($x)" #add the form name to our array
    $SyncClass.SyncHash.Add("form$($x)", (Get-Variable -Name "form$($x)").Value) #add the form object to our synched hashtable
}
#################################
## Create Controls (Buttons, etc)
#################################
$controls = @()
$xp = '[^a-zA-Z_0-9]' # All characters that are not a-Z, 0-9, or _
foreach($x in $vx)
{
    $xaml = (Get-Variable -Name "xaml$($x)").Value #load the xaml we created earlier
    $xaml.SelectNodes("//*[@Name]") | %{ #find all nodes with a "Name" attribute
        $cname = "form$($x)Control$(($_.Name -replace $xp, '_'))"
        Set-Variable -Name "$cname" -Value $SyncClass.SyncHash."form$($x)".FindName($_.Name) #create a variale to hold the control/object
        $controls += (Get-Variable -Name "form$($x)Control$($_.Name)").Name #add the control name to our array
        $SyncClass.SyncHash.Add($cname, $SyncClass.SyncHash."form$($x)".FindName($_.Name)) #add the control directly to the hashtable
    }
}
############################
## FORMS AND CONTROLS OUTPUT
############################
<# Write-Host -ForegroundColor Cyan "The following forms were created:"
$forms | %{ Write-Host -ForegroundColor Yellow "  `$$_"} #output all forms to screen
if($controls.Count -gt 0){
    Write-Host ""
    Write-Host -ForegroundColor Cyan "The following controls were created:"
    $controls | %{ Write-Host -ForegroundColor Yellow "  `$$_"} #output all named controls to screen
} #>
#######################
## DISABLE A/V AUTOPLAY
#######################
foreach($x in $vx)
{
    $carray = @()
    $fts = $syncClass.SyncHash."form$($x)"
    foreach($c in $fts.Content.Children)
    {
        if($c.GetType().Name -eq "MediaElement") #find all controls with the type MediaElement
        {
            $c.LoadedBehavior = "Manual" #Don't autoplay
            $c.UnloadedBehavior = "Stop" #When the window closes, stop the music
            $carray += $c #add the control to an array
        }
    }
    if($carray.Count -gt 0)
    {
        New-Variable -Name "form$($x)PoSHPFCleanupAudio" -Value $carray -Force # Store the controls in an array to be accessed later
        $syncClass.SyncHash."form$($x)".Add_Closed({
            foreach($c in (Get-Variable "form$($x)PoSHPFCleanupAudio").Value)
            {
                $c.Source = $null #stops any currently playing media
            }
        })
    }
}

#####################
## RUNSPACE FUNCTIONS
#####################
## Yo dawg... Runspace to clean up Runspaces
## Thank you Boe Prox / Stephen Owen
#region RSCleanup
$Script:JobCleanup = [hashtable]::Synchronized(@{}) 
$Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList)) #hashtable to store all these runspaces
$jobCleanup.Flag = $True #cleanup jobs
$newRunspace =[runspacefactory]::CreateRunspace() #create a new runspace for this job to cleanup jobs to live
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup) #pass the jobCleanup variable to the runspace
$newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) #pass the jobs variable to the runspace
$jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
    #Routine to handle completed runspaces
    Do {    
        Foreach($runspace in $jobs) {            
            If ($runspace.Runspace.isCompleted) {                         #if runspace is complete
                [void]$runspace.powershell.EndInvoke($runspace.Runspace)  #then end the script
                $runspace.powershell.dispose()                            #dispose of the memory
                $runspace.Runspace = $null                                #additional garbage collection
                $runspace.powershell = $null                              #additional garbage collection
            } 
        }
        #Clean out unused runspace jobs
        $temphash = $jobs.clone()
        $temphash | Where {
            $_.runspace -eq $Null
        } | ForEach {
            $jobs.remove($_)
        }        
        Start-Sleep -Seconds 1 #lets not kill the processor here 
    } while ($jobCleanup.Flag)
})
$jobCleanup.PowerShell.Runspace = $newRunspace
$jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke() 
#endregion RSCleanup

#This function creates a new runspace for a script block to execute
#so that you can do your long running tasks not in the UI thread.
#Also the SyncClass is passed to this runspace so you can do UI
#updates from this thread as well.
function Start-BackgroundScriptBlock($scriptBlock){
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncClass",$SyncClass) 
    $PowerShell = [PowerShell]::Create().AddScript($scriptBlock)
    $PowerShell.Runspace = $newRunspace
    $PowerShell.BeginInvoke()

    #Add it to the job list so that we can make sure it is cleaned up
<#     [void]$Jobs.Add(
        [pscustomobject]@{
            PowerShell = $PowerShell
            Runspace = $PowerShell.BeginInvoke()
        }
    ) #>
}
#================================================
#   Customizations
#================================================
[string]$ModuleVersion = Get-Module -Name AutopilotOOBE | Sort-Object -Property Version | Select-Object -ExpandProperty Version -Last 1
#================================================
#   Window Functions
#   Minimize Command and PowerShell Windows
#================================================
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
function Hide-CmdWindow() {
    $CMDProcess = Get-Process -Name cmd -ErrorAction Ignore
    foreach ($Item in $CMDProcess) {
        $null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $Item.id).MainWindowHandle, 2)
    }
}
function Hide-PowershellWindow() {
    $null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}
function Show-PowershellWindow() {
    $null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
#================================================
#   Sidebar
#================================================
if (Test-AutopilotOOBEconnection) {
    $formMainWindowControlOnlineStatusLabel.Background = 'Green'
}
else {
    $formMainWindowControlOnlineStatusLabel.Background = 'Red'
}

try {
    $Tpm = (Get-CimInstance -Namespace "root\CIMV2\Security\MicrosoftTPM" -ClassName Win32_Tpm).SpecVersion
}
catch {}

if ($Tpm -match '2.0') {
    $formMainWindowControlTpmVersionLabel.Content = "TPM: 2.0"
    $formMainWindowControlTpmVersionLabel.Background = "Green"
}
elseif ($Tpm -match '1.2') {
    $formMainWindowControlTpmVersionLabel.Content = "TPM: 1.2"
    $formMainWindowControlTpmVersionLabel.Background = "Red"
}
else {
    $formMainWindowControlTpmVersionLabel.Content = "TPM"
    $formMainWindowControlTpmVersionLabel.Background = "Red"
    #$formMainWindowControlTpmVersionLabel.Visibility = "Collapsed"
}

$formMainWindowControlCSManufacturerControl.Content = ((Get-CimInstance -ClassName CIM_ComputerSystem).Manufacturer).Trim()

if ($formMainWindowControlCSManufacturerControl.Content -match 'Lenovo') {
    $formMainWindowControlCSModelControl.Content = ((Get-CimInstance -ClassName Win32_ComputerSystemProduct).Version).Trim()
}
else {
    $formMainWindowControlCSModelControl.Content = ((Get-CimInstance -ClassName CIM_ComputerSystem).Model).Trim()
}

$SerialNumber = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
$formMainWindowControlSerialNumberLabel.Content = $SerialNumber

$BiosVersion = ((Get-CimInstance -ClassName Win32_BIOS).SMBIOSBIOSVersion).Trim()
$formMainWindowControlBiosVersionLabel.Content = "BIOS $BiosVersion"
#================================================
#   Parameters
#================================================
$AutopilotOOBEParams = (Get-Command Start-AutopilotOOBE).Parameters
#================================================
#   Heading
#================================================
$formMainWindowControlHeading.Content = $Global:AutopilotOOBE.Title
#================================================
#   SubHeading
#================================================
$Global:GetRegCurrentVersion = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

$SubTitleProductName = ($Global:GetRegCurrentVersion).ProductName

if ($Global:GetRegCurrentVersion.DisplayVersion -gt 0) {
    $SubTitleDisplayVersion = ($Global:GetRegCurrentVersion).DisplayVersion
}
else {
    $SubTitleDisplayVersion = ($Global:GetRegCurrentVersion).ReleaseId
}

$SubTitleBuildNumber = "$($Global:GetRegCurrentVersion.CurrentBuild).$($Global:GetRegCurrentVersion.UBR)"

$formMainWindowControlSubHeading.Content = "$SubTitleProductName $SubTitleDisplayVersion ($SubTitleBuildNumber)"
#================================================
#   GroupTag Control
#================================================
# Disable the Control
if ($Disabled -contains 'GroupTag') {
    $formMainWindowControlGroupTagComboBox.IsEnabled = $false
}

# Hide the Control
if ($Hidden -contains 'GroupTag') {
    $formMainWindowControlGroupTagStackPanel.Visibility = 'Collapsed'
}

# Populate the ComboBox
$Global:AutopilotOOBE.GroupTagOptions | ForEach-Object {
    $formMainWindowControlGroupTagComboBox.Items.Add($_) | Out-Null
}

# Set the ComboBox Default
if ($Global:AutopilotOOBE.GroupTag) {
    $formMainWindowControlGroupTagComboBox.Text = $Global:AutopilotOOBE.GroupTag
}
#================================================
#   AddToGroup Control
#================================================
# Disable the Control
if ($Disabled -contains 'AddToGroup') {
    $formMainWindowControlAddToGroupComboBox.IsEnabled = $false
}

# Hide the Control
if ($Hidden -contains 'AddToGroup') {
    $formMainWindowControlAddToGroupStackPanel.Visibility = 'Collapsed'
}

# Populate the Control
$Global:AutopilotOOBE.AddToGroupOptions | ForEach-Object {
    $formMainWindowControlAddToGroupComboBox.Items.Add($_) | Out-Null
}

# Set the Default
if ($Global:AutopilotOOBE.AddToGroup) {
    $formMainWindowControlAddToGroupComboBox.Text = $Global:AutopilotOOBE.AddToGroup
}
#================================================
#   AssignedComputerName Control
#================================================
# Disable the Control
if ($Disabled -contains 'AssignedComputerName') {
    $formMainWindowControlAssignedComputerNameTextBox.IsEnabled = $false
}

# Hide the Control
if ($Hidden -contains 'AssignedComputerName') {
    $formMainWindowControlAssignedComputerNameStackPanel.Visibility = 'Collapsed'
}

# Populate the Control
$formMainWindowControlAssignedComputerNameTextBox.Text = $Global:AutopilotOOBE.AssignedComputerNameExample
if ($Global:AutopilotOOBE.AssignedComputerName -gt 0) {
    $formMainWindowControlAssignedComputerNameTextBox.Text = $Global:AutopilotOOBE.AssignedComputerName
}
#================================================
#   PostAction Control
#================================================
# Disable the Control
if ($Disabled -contains 'PostAction') {$formMainWindowControlPostActionComboBox.IsEnabled = $false}

# Hide the Control
if ($Hidden -contains 'PostAction') {
    $formMainWindowControlPostActionStackPanel.Visibility = 'Collapsed'
}

# Values
$PostActionComboBoxValues = @(
    'Quit',
    'Restart Computer',
    'Shutdown Computer',
    'Sysprep /oobe /quit',
    'Sysprep /oobe /reboot',
    'Sysprep /oobe /shutdown',
    'Sysprep /generalize /oobe /reboot',
    'Sysprep /generalize /oobe /shutdown'
)

# Populate the ComboBox
$PostActionComboBoxValues | ForEach-Object {
    $formMainWindowControlPostActionComboBox.Items.Add($_) | Out-Null
}

# Set the Default
switch ($Global:AutopilotOOBE.PostAction) {
    'Quit'                  {$formMainWindowControlPostActionComboBox.SelectedValue = 'Quit'}
    'Restart'               {$formMainWindowControlPostActionComboBox.SelectedValue = 'Restart Computer'}
    'Shutdown'              {$formMainWindowControlPostActionComboBox.SelectedValue = 'Shutdown Computer'}
    'Sysprep'               {$formMainWindowControlPostActionComboBox.SelectedValue = 'Sysprep /oobe /quit'}
    'SysprepReboot'         {$formMainWindowControlPostActionComboBox.SelectedValue = 'Sysprep /oobe /reboot'}
    'SysprepShutdown'       {$formMainWindowControlPostActionComboBox.SelectedValue = 'Sysprep /oobe /shutdown'}
    'GeneralizeReboot'      {$formMainWindowControlPostActionComboBox.SelectedValue = 'Sysprep /generalize /oobe /reboot'}
    'GeneralizeShutdown'    {$formMainWindowControlPostActionComboBox.SelectedValue = 'Sysprep /generalize /oobe /shutdown'}
    Default                 {$formMainWindowControlPostActionComboBox.SelectedValue = 'Quit'}
}
#================================================
#   Assign CheckBox
#================================================
# Disable the Control
if ($Disabled -contains 'Assign') {$formMainWindowControlAssignCheckBox.IsEnabled = $false}

# Hide the Control
if ($Hidden -contains 'Assign') {
    $formMainWindowControlAssignStackPanel.Visibility = 'Collapsed'
}

# Set the Default
if ($Global:AutopilotOOBE.Assign -eq $true) {
    $formMainWindowControlAssignCheckBox.IsChecked = $true
}
#================================================
#   Register Control
#================================================
# Hide the Control
if ($Hidden -contains 'Register') {
    $formMainWindowControlRegisterStackPanel.Visibility = 'Collapsed'

    if ($Global:RegAutoPilot.CloudAssignedForcedEnrollment -eq 1) {
        $CloudAssignedForcedEnrollment = 'Yes'
        $formMainWindow.Title = "AutopilotOOBE $ModuleVersion : Quit to OOBE"
    }
    else {
        $CloudAssignedForcedEnrollment = 'No'
        #$formMainWindow.Title = "AutopilotOOBE $ModuleVersion Device Not Registered"
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

    $formMainWindowControlInformationLabel.Content = @"
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
else {
    $formMainWindow.Title = "AutopilotOOBE $ModuleVersion : Register Device"
}
#================================================
#   Run Controls
#================================================
# Hide the Control
if ($Hidden -contains 'Run') {
    $formMainWindowControlRunStackPanel.Visibility = 'Collapsed'
}

# Values
$RunComboBoxValues = @(
    'Restart Computer',
    'Shutdown Computer',
    'Command Prompt',
    'PowerShell',
    'PowerShell ISE',
    'Open Event Viewer',
    'Open Windows Explorer',
    'Show Network and Wireless Settings',
    'Show Windows Security',
    'Show Windows Settings',
    'AutopilotDiagnostics',
    'AutopilotDiagnostics Online',
    'MDMDiagnosticsTool -out C:\Temp',
    'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab',
    'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\AutopilotTPM.cab',
    'TPM Get',
    'TPM Clear',
    'TPM Initialize',
    'Sysprep /oobe /quit',
    'Sysprep /oobe /reboot',
    'Sysprep /oobe /shutdown',
    'Sysprep /audit /reboot'
)

# Populate the ComboBox
$RunComboBoxValues | ForEach-Object {
    $formMainWindowControlRunComboBox.Items.Add($_) | Out-Null
}

# Set the ComboBox Default
switch ($Global:AutopilotOOBE.Run) {
    'Restart'                       {$formMainWindowControlRunComboBox.SelectedValue = 'Restart Computer'}
    'Shutdown'                      {$formMainWindowControlRunComboBox.SelectedValue = 'Shutdown Computer'}
    'CommandPrompt'                 {$formMainWindowControlRunComboBox.SelectedValue = 'Command Prompt'}
    'PowerShell'                    {$formMainWindowControlRunComboBox.SelectedValue = 'PowerShell'}
    'PowerShellISE'                 {$formMainWindowControlRunComboBox.SelectedValue = 'PowerShell ISE'}
    'EventViewer'                   {$formMainWindowControlRunComboBox.SelectedValue = 'Open Event Viewer'}
    'NetworkingWireless'            {$formMainWindowControlRunComboBox.SelectedValue = 'Show Network and Wireless Settings'}
    'WindowsExplorer'               {$formMainWindowControlRunComboBox.SelectedValue = 'Open Windows Explorer'}
    'WindowsSettings'               {$formMainWindowControlRunComboBox.SelectedValue = 'Show Windows Settings'}
    'AutopilotDiagnostics'          {$formMainWindowControlRunComboBox.SelectedValue = 'AutopilotDiagnostics'}
    'AutopilotDiagnosticsOnline'    {$formMainWindowControlRunComboBox.SelectedValue = 'AutopilotDiagnostics Online'}
    'MDMDiag'                       {$formMainWindowControlRunComboBox.SelectedValue = 'MDMDiagnosticsTool -out C:\Temp'}
    'MDMDiagAutopilot'              {$formMainWindowControlRunComboBox.SelectedValue = 'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab'}
    'MDMDiagAutopilotTPM'           {$formMainWindowControlRunComboBox.SelectedValue = 'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\AutopilotTPM.cab'}
    'Sysprep'                       {$formMainWindowControlRunComboBox.SelectedValue = 'Sysprep /oobe /quit'}
    'SysprepReboot'                 {$formMainWindowControlRunComboBox.SelectedValue = 'Sysprep /oobe /reboot'}
    'SysprepShutdown'               {$formMainWindowControlRunComboBox.SelectedValue = 'Sysprep /oobe /shutdown'}
    'SysprepAudit'                  {$formMainWindowControlRunComboBox.SelectedValue = 'Sysprep /audit /reboot'}
    Default                         {$formMainWindowControlRunComboBox.SelectedValue = 'PowerShell'}
}

# Add Click
$formMainWindowControlRunButton.add_Click( {
    switch ($formMainWindowControlRunComboBox.SelectedValue) {
        'Restart Computer'                  {Restart-Computer}
        'Shutdown Computer'                 {Stop-Computer}
        'Command Prompt'                    {Start-Process Cmd.exe}
        'PowerShell'                        {Start-Process PowerShell.exe -ArgumentList "-Nologo"}
        'PowerShell ISE'                    {Start-Process PowerShell_ISE.exe}
        'Open Event Viewer'                 {Start-Process -FilePath PowerShell.exe -ArgumentList '-NoLogo -Window Minimized',"-Command Invoke-AutopilotOOBEcmd EventViewer"}
        'Open Windows Explorer'             {Start-Process Explorer.exe}
        'Show Network and Wireless Settings'{Start-Process ms-availablenetworks:}
        'Show Windows Security'             {Start-Process PowerShell.exe -ArgumentList "Add-AppxPackage -Register -DisableDevelopmentMode 'C:\Windows\SystemApps\Microsoft.Windows.SecHealthUI_cw5n1h2txyewy\AppXManifest.xml';start windowsdefender:"}
        'Show Windows Settings'             {Start-Process ms-settings:}
        'AutopilotDiagnostics'              {Start-Process -FilePath PowerShell.exe -ArgumentList '-NoLogo -NoExit',"-Command Invoke-AutopilotOOBEcmd AutopilotDiagnostics"}
        'AutopilotDiagnostics Online'       {Start-Process -FilePath PowerShell.exe -ArgumentList '-NoLogo -NoExit',"-Command Invoke-AutopilotOOBEcmd AutopilotDiagnosticsOnline"}
        'MDMDiagnosticsTool -out C:\Temp'                                       {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-out C:\Temp"}
        'MDMDiagnosticsTool -area Autopilot -cab C:\Temp\Autopilot.cab'         {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-area Autopilot","-cab C:\Temp\Autopilot.cab"}
        'MDMDiagnosticsTool -area Autopilot;TPM -cab C:\Temp\AutopilotTPM.cab'  {Start-Process MDMDiagnosticsTool.exe -ArgumentList "-area Autopilot;TPM","-cab C:\Temp\AutopilotTPM.cab"}
        'TPM Get'                           {Start-Process -FilePath PowerShell.exe -ArgumentList '-NoLogo -NoExit',"-Command Invoke-AutopilotOOBEcmd GetTpm"}
        'TPM Clear'                         {Start-Process -FilePath PowerShell.exe -ArgumentList '-NoLogo -NoExit',"-Command Invoke-AutopilotOOBEcmd ClearTpm"}
        'TPM Initialize'                    {Start-Process -FilePath PowerShell.exe -ArgumentList '-NoLogo -NoExit',"-Command Invoke-AutopilotOOBEcmd InitializeTpm"}
        'Sysprep /oobe /quit'               {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/quit"}
        'Sysprep /oobe /reboot'             {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/reboot"}
        'Sysprep /oobe /shutdown'           {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/shutdown"}
        'Sysprep /audit /reboot'            {Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/audit", "/reboot"}
        Default {}
    }
})
#================================================
#   Parameter Docs
#================================================
if ($Hidden -contains 'Docs') {
    $formMainWindowControlDocsStackPanel.Visibility = 'Collapsed'
}

$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot Documentation') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot Overview') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot User-Driven Mode') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot for Pre-Provisioned Deployment') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot Deployment for Existing Devices') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Manually register devices with Windows Autopilot') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot Troubleshooting Overview') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Troubleshoot Autopilot Device Import and Enrollment') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Troubleshoot Autopilot OOBE Issues') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Troubleshoot Azure Active Directory Join Issues') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot Known Issues') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Windows Autopilot Resolved Issues') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Sysprep Overview') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Sysprep Audit Mode Overview') | Out-Null
$formMainWindowControlDocsComboBox.Items.Add('Sysprep Command-Line Options') | Out-Null

if ($Hidden -contains 'Register') {
    $formMainWindowControlDocsComboBox.SelectedValue = 'Troubleshoot Autopilot OOBE Issues'
}
else {
    $formMainWindowControlDocsComboBox.SelectedValue = 'Windows Autopilot Documentation'
}

if ($Global:AutopilotOOBE.Docs) {
    $formMainWindowControlDocsComboBox.Items.Add($Global:AutopilotOOBE.Docs) | Out-Null
    $formMainWindowControlDocsComboBox.SelectedValue = $Global:AutopilotOOBE.Docs
}

$formMainWindowControlDocsButton.add_Click( {
    Write-Host -ForegroundColor Cyan "Run: $($formMainWindowControlDocsComboBox.SelectedValue)"

    if ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot Documentation') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot Overview') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/windows-autopilot'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot User-Driven Mode') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/user-driven'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot for Pre-Provisioned Deployment') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/pre-provision'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot Deployment for Existing Devices') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/existing-devices'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Manually register devices with Windows Autopilot') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/add-devices'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot Troubleshooting Overview') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshooting'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Troubleshoot Autopilot Device Import and Enrollment') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshoot-device-enrollment'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Troubleshoot Autopilot OOBE Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshoot-oobe'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Troubleshoot Azure Active Directory Join Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/troubleshoot-aad-join'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot Known Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/known-issues'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Windows Autopilot Resolved Issues') {Start-Process 'https://docs.microsoft.com/en-us/mem/autopilot/resolved-issues'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Sysprep Overview') {Start-Process 'https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/sysprep--system-preparation--overview'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Sysprep Audit Mode Overview') {Start-Process 'https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/audit-mode-overview'}
    elseif ($formMainWindowControlDocsComboBox.SelectedValue -eq 'Sysprep Command-Line Options') {Start-Process 'https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/sysprep-command-line-options'}
    else {
        try {
            Start-Process $formMainWindowControlDocsComboBox.SelectedValue
        }
        catch {
            Write-Warning "Could not execute $($formMainWindowControlDocsComboBox.SelectedValue)"
        }
    }
})
#================================================
#   RegisterButton
#================================================
if ($env:UserName -ne 'defaultuser0') {
    $formMainWindowControlRegisterButton.IsEnabled = $false
}

$formMainWindowControlRegisterButton.add_Click( {
    $formMainWindow.Close()
    Show-PowershellWindow

    $Params = @{
        Online = $true
    }

    if ($formMainWindowControlAssignCheckbox.IsChecked) {
        $Params.Assign = $true
    }

    if ($formMainWindowControlAddToGroupComboBox.Text -gt 0) {
        $Params.AddToGroup = $formMainWindowControlAddToGroupComboBox.Text
    }

    if ($formMainWindowControlGroupTagComboBox.Text -gt 0) {
        $Params.GroupTag = $formMainWindowControlGroupTagComboBox.Text
    }

    if (($formMainWindowControlAssignedComputerNameTextBox.Text -gt 0) -and ($formMainWindowControlAssignedComputerNameTextBox.Text -notmatch $Global:AutopilotOOBE.AssignedComputerNameExample)) {
        $Params.AssignedComputerName = $formMainWindowControlAssignedComputerNameTextBox.Text
    }

    $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-AutopilotOOBE.log"
    Start-Transcript -Path (Join-Path "$env:SystemRoot\Temp" $Transcript) -ErrorAction Ignore

    Write-Host -ForegroundColor Cyan "Install-Script Get-WindowsAutoPilotInfo"
    Start-Sleep -Seconds 3
    Install-Script Get-WindowsAutoPilotInfo -Force -Verbose

    Write-Host ($Params | Out-String)
    Write-Host -ForegroundColor Cyan "Get-WindowsAutoPilotInfo @Params"

    Start-Sleep -Seconds 3
    $formMainWindow.Title = "AutopilotOOBE $ModuleVersion : Registering Device"
    Get-WindowsAutoPilotInfo @Params
    $formMainWindow.Title = "AutopilotOOBE $ModuleVersion : Restart Device"

    if ((Get-Process -Name powershell).MainWindowTitle -match 'Running') {
        Write-Warning "Waiting for Start-OOBEDeploy to finish"
    }

    while ((Get-Process -Name powershell).MainWindowTitle -match 'Running') {
        Start-Sleep -Seconds 10
    }

    if ($formMainWindowControlPostActionComboBox.SelectedValue -eq 'Restart Computer') {Restart-Computer}
    if ($formMainWindowControlPostActionComboBox.SelectedValue -eq 'Shutdown Computer') {Stop-Computer}

    if ($formMainWindowControlPostActionComboBox.SelectedValue -match 'Sysprep') {
        Write-Host -ForegroundColor Cyan "Executing Sysprep"

        if ($formMainWindowControlPostActionComboBox.SelectedValue -match 'quit') {
            Start-Sleep -Seconds 3
            Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/quit" -Wait
        }
        
        elseif (($formMainWindowControlPostActionComboBox.SelectedValue -match 'reboot') -and ($formMainWindowControlPostActionComboBox.SelectedValue -match 'generalize')) {
            Start-Sleep -Seconds 3
            Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "generalize", "/oobe", "/reboot" -Wait
        }
        elseif (($formMainWindowControlPostActionComboBox.SelectedValue -match 'shutdown') -and ($formMainWindowControlPostActionComboBox.SelectedValue -match 'generalize')) {
            Start-Sleep -Seconds 3
            Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "generalize", "/oobe", "/reboot" -Wait
        }

        elseif ($formMainWindowControlPostActionComboBox.SelectedValue -match 'reboot') {
            Start-Sleep -Seconds 3
            Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/reboot" -Wait
        }
        elseif ($formMainWindowControlPostActionComboBox.SelectedValue -match 'shutdown') {
            Start-Sleep -Seconds 3
            Start-Process "$env:SystemRoot\System32\Sysprep\Sysprep.exe" -ArgumentList "/oobe", "/shutdown" -Wait
        }
    }
})
#================================================
#   Hide Windows
#================================================
Hide-CmdWindow
Hide-PowershellWindow
########################
## WIRE UP YOUR CONTROLS
########################
# simple example: $formMainWindowControlButton.Add_Click({ your code })
#
# example with BackgroundScriptBlock and UpdateElement
# $formmainControlButton.Add_Click({
#     $sb = {
#         $SyncClass.UpdateElement("formmainControlProgress","Value",25)
#     }
#     Start-BackgroundScriptBlock $sb
# })

############################
###### DISPLAY DIALOG ######
############################
[void]$formMainWindow.ShowDialog()

##########################
##### SCRIPT CLEANUP #####
##########################
$jobCleanup.Flag = $false #Stop Cleaning Jobs
$jobCleanup.PowerShell.Runspace.Close() #Close the runspace
$jobCleanup.PowerShell.Dispose() #Remove the runspace from memory