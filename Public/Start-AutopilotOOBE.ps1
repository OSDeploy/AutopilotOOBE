function Start-AutopilotOOBE {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$CustomProfile,

        [ValidateSet (
            'GroupTag',
            'AddToGroup',
            'AssignedUser',
            'AssignedComputerName',
            'PostAction',
            'Assign'
        )]
        [string[]]$Disabled,

        [ValidateSet (
            'GroupTag',
            'AddToGroup',
            'AssignedUser',
            'AssignedComputerName',
            'PostAction',
            'Assign',
            'Register',
            'Run',
            'Docs'
        )]
        [string[]]$Hidden,

        [string]$AddToGroup,
        [string[]]$AddToGroupOptions,
        [switch]$Assign,
        [string]$AssignedUser,
        [string]$AssignedUserExample = 'someone@example.com',
        [string]$AssignedComputerName,
        [string]$AssignedComputerNameExample = 'Azure AD Join Only',
        [string]$GroupTag,
        [string[]]$GroupTagOptions,
        [ValidateSet (
            'Quit',
            'Restart',
            'Shutdown',
            'Sysprep',
            'SysprepReboot',
            'SysprepShutdown',
            'GeneralizeReboot',
            'GeneralizeShutdown'
        )]
        [string]$PostAction = 'Quit',
        [ValidateSet (
            'CommandPrompt',
            'PowerShell',
            'PowerShellISE',
            'WindowsExplorer',
            'WindowsSettings',
            'NetworkingWireless',
            'Restart',
            'Shutdown',
            'Sysprep',
            'SysprepReboot',
            'SysprepShutdown',
            'SysprepAudit',
            'EventViewer',
            'GetAutopilotDiagnostics',
            'GetAutopilotDiagnosticsOnline',
            'MDMDiag',
            'MDMDiagAutopilot',
            'MDMDiagAutopilotTPM'
        )]
        [string]$Run = 'PowerShell',
        [string]$Docs,
        [string]$Title = 'Autopilot Manual Registration'
    )
    #=======================================================================
    #   WinPE and WinOS Start
    #=======================================================================
    if ($env:SystemDrive -eq 'X:') {
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Green "Start-AutopilotOOBE in WinPE"
        $ProgramDataOSDeploy = 'C:\ProgramData\OSDeploy'
        $JsonPath = "$ProgramDataOSDeploy\OSDeploy.AutopilotOOBE.json"
    }
    if ($env:SystemDrive -ne 'X:') {
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Green "Start-AutopilotOOBE"
        $ProgramDataOSDeploy = "$env:ProgramData\OSDeploy"
        $JsonPath = "$ProgramDataOSDeploy\OSDeploy.AutopilotOOBE.json"
    }
    #=======================================================================
    #   WinOS Transcript
    #=======================================================================
    if ($env:SystemDrive -ne 'X:') {
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Start-Transcript"
        $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-AutopilotOOBE.log"
        Start-Transcript -Path (Join-Path "$env:SystemRoot\Temp" $Transcript) -ErrorAction Ignore
    }
    #=======================================================================
    #   WinOS Console Disable Line Wrap
    #=======================================================================
    reg add HKCU\Console /v LineWrap /t REG_DWORD /d 0 /f
    #=======================================================================
    #   Custom Profile Sample Variables
    #=======================================================================
    if ($CustomProfile -eq 'Sample') {
        $Title = 'Sample Autopilot Registration'
        $AddToGroup = 'Administrators'
        $AssignedUserExample = 'someone@osdeploy.com'
        $AssignedComputerName = 'OSD-' + ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $PostAction = 'Shutdown'
        $Assign = $true
        $Run = 'PowerShell'
        $Docs = 'https://www.osdeploy.com/'
        $Hidden = 'GroupTag'
    }
    #=======================================================================
    #   Custom Profile
    #=======================================================================
    if ($CustomProfile) {
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Loading AutopilotOOBE Custom Profile $CustomProfile"

        $CustomProfileJson = Get-ChildItem "$($MyInvocation.MyCommand.Module.ModuleBase)\CustomProfile" *.json | Where-Object {$_.BaseName -eq $CustomProfile} | Select-Object -First 1

        if ($CustomProfileJson) {
            Write-Host -ForegroundColor DarkGray "Saving Module CustomProfile to $JsonPath"
            if (!(Test-Path "$ProgramDataOSDeploy")) {New-Item "$ProgramDataOSDeploy" -ItemType Directory -Force | Out-Null}
            Copy-Item -Path $CustomProfileJson.FullName -Destination $JsonPath -Force -ErrorAction Ignore
        }
    }
    #=======================================================================
    #   Import Json
    #=======================================================================
    if (Test-Path $JsonPath) {
        Write-Host -ForegroundColor DarkGray "Importing Configuration $JsonPath"
        $ImportAutopilotOOBE = @()
        $ImportAutopilotOOBE = Get-Content -Raw -Path $JsonPath | ConvertFrom-Json
    
        $ImportAutopilotOOBE.PSObject.Properties | ForEach-Object {
            if ($_.Value -match 'IsPresent=True') {
                $_.Value = $true
            }
            if ($_.Value -match 'IsPresent=False') {
                $_.Value = $false
            }
            if ($null -eq $_.Value) {
                Continue
            }
            Set-Variable -Name $_.Name -Value $_.Value -Force
        }
    }
    #=======================================================================
    #   WinOS PSGallery
    #=======================================================================
    if ($env:SystemDrive -ne 'X:') {
        $PSGalleryIP = (Get-PSRepository -Name PSGallery).InstallationPolicy
        if ($PSGalleryIP -eq 'Untrusted') {
            Write-Host -ForegroundColor DarkGray "========================================================================="
            Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Set-PSRepository -Name PSGallery -InstallationPolicy Trusted"
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    }
    #=======================================================================
    #   WinOS Autopilot Registered
    #=======================================================================
    if (($env:SystemDrive -ne 'X:') -and ($Global:RegAutoPilot.CloudAssignedForcedEnrollment -eq 1)) {
        $Disabled = 'GroupTag','AddToGroup','AssignedUser','AssignedComputerName','PostAction','Assign'
        $Hidden = 'GroupTag','AddToGroup','AssignedUser','AssignedComputerName','PostAction','Assign','Register'
        $Run = 'MDMDiagAutopilotTPM'
        $Title = 'Autopilot Registration Information'
    }
    #=======================================================================
    #   WinPE and WinOS Configuration Json
    #=======================================================================
    $Global:AutopilotOOBE = [ordered]@{
        AddToGroup = $AddToGroup
        AddToGroupOptions = $AddToGroupOptions
        Assign = $Assign
        AssignedUser = $AssignedUser
        AssignedUserExample = $AssignedUserExample
        AssignedComputerName = $AssignedComputerName
        AssignedComputerNameExample = $AssignedComputerNameExample
        Disabled = $Disabled
        GroupTag = $GroupTag
        GroupTagOptions = $GroupTagOptions
        Hidden = $Hidden
        PostAction = $PostAction
        Run = $Run
        Docs = $Docs
        Title = $Title
    }

    if ($env:SystemDrive -eq 'X:') {
        if (!(Test-Path "$ProgramDataOSDeploy")) {New-Item "$ProgramDataOSDeploy" -ItemType Directory -Force | Out-Null}
        Write-Host -ForegroundColor DarkGray "Exporting Configuration $ProgramDataOSDeploy\OSDeploy.AutopilotOOBE.json"
        @($Global:AutopilotOOBE.Keys) | ForEach-Object { 
            if (-not $Global:AutopilotOOBE[$_]) { $Global:AutopilotOOBE.Remove($_) }
        }
        $Global:AutopilotOOBE | ConvertTo-Json | Out-File "$ProgramDataOSDeploy\OSDeploy.AutopilotOOBE.json" -Force
    }
    else {
        Write-Host -ForegroundColor DarkGray "Exporting Configuration $env:Temp\OSDeploy.AutopilotOOBE.json"
        @($Global:AutopilotOOBE.Keys) | ForEach-Object { 
            if (-not $Global:AutopilotOOBE[$_]) { $Global:AutopilotOOBE.Remove($_) }
        }
        $Global:AutopilotOOBE | ConvertTo-Json | Out-File "$env:Temp\OSDeploy.AutopilotOOBE.json" -Force
        #=======================================================================
        #   Watch-AutopilotOOBEeventlog
        #=======================================================================
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Watch-AutopilotOOBEeventlog"
        Write-Host -ForegroundColor DarkCyan 'The EventLog is being monitored for MDM Diagnostic Events in a minimized window'
        Write-Host -ForegroundColor DarkCyan 'Use Alt+Tab to view the progress in the separate PowerShell session'
        Start-Process PowerShell.exe -WindowStyle Minimized -ArgumentList "-NoExit -Command Watch-AutopilotOOBEeventlog"
        #=======================================================================
        #   Test-AutopilotOOBEnetwork
        #=======================================================================
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Test-AutopilotOOBEnetwork"
        Write-Host -ForegroundColor DarkCyan 'Required Autopilot network addresses are being tested in a minimized window'
        Write-Host -ForegroundColor DarkCyan 'Use Alt+Tab to view the progress in the separate PowerShell session'
        Start-Process PowerShell.exe -WindowStyle Minimized -ArgumentList "-NoExit -Command Test-AutopilotOOBEnetwork"
        #=======================================================================
        #   Date Time
        #=======================================================================
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Verify Date and Time"
        Write-Host -ForegroundColor DarkCyan 'Make sure the Time is set properly in the System BIOS as this can cause issues'
        Get-Date
        Get-TimeZone
        Start-Sleep -Seconds 5
        #=======================================================================
        #   Test-AutopilotRegistry
        #=======================================================================
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Test-AutopilotRegistry"
        Write-Host -ForegroundColor DarkCyan 'Gathering Autopilot Registration information from the Registry'
        $Global:RegAutoPilot = Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Provisioning\Diagnostics\AutoPilot'
        
        Write-Host -ForegroundColor Gray "IsAutoPilotDisabled: $($Global:RegAutoPilot.IsAutoPilotDisabled)"
        Write-Host -ForegroundColor Gray "CloudAssignedForcedEnrollment: $($Global:RegAutoPilot.CloudAssignedForcedEnrollment)"
        Write-Host -ForegroundColor Gray "CloudAssignedTenantDomain: $($Global:RegAutoPilot.CloudAssignedTenantDomain)"
        Write-Host -ForegroundColor Gray "CloudAssignedTenantId: $($Global:RegAutoPilot.CloudAssignedTenantId)"
        Write-Host -ForegroundColor Gray "CloudAssignedTenantUpn: $($Global:RegAutoPilot.CloudAssignedTenantUpn)"
        Write-Host -ForegroundColor Gray "CloudAssignedLanguage: $($Global:RegAutoPilot.CloudAssignedLanguage)"
    
        if ($Global:RegAutoPilot.CloudAssignedForcedEnrollment -eq 1) {
            Write-Host -ForegroundColor Gray "TenantId: $($Global:RegAutoPilot.TenantId)"
            Write-Host -ForegroundColor Gray "CloudAssignedMdmId: $($Global:RegAutoPilot.CloudAssignedMdmId)"
            Write-Host -ForegroundColor Gray "AutopilotServiceCorrelationId: $($Global:RegAutoPilot.AutopilotServiceCorrelationId)"
            Write-Host -ForegroundColor Gray "CloudAssignedOobeConfig: $($Global:RegAutoPilot.CloudAssignedOobeConfig)"
            Write-Host -ForegroundColor Gray "CloudAssignedTelemetryLevel: $($Global:RegAutoPilot.CloudAssignedTelemetryLevel)"
            Write-Host -ForegroundColor Gray "IsDevicePersonalized: $($Global:RegAutoPilot.IsDevicePersonalized)"
            Write-Host -ForegroundColor Gray "SetTelemetryLevel_Succeeded_With_Level: $($Global:RegAutoPilot.SetTelemetryLevel_Succeeded_With_Level)"
            Write-Host -ForegroundColor Gray "IsForcedEnrollmentEnabled: $($Global:RegAutoPilot.IsForcedEnrollmentEnabled)"
            Write-Host -ForegroundColor Green "This device has already been Autopilot Registered. Registration will not be enabled"
            Start-Sleep -Seconds 2
        }
        #=======================================================================
        #   Launch
        #=======================================================================
        Write-Host -ForegroundColor DarkGray "========================================================================="
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Starting AutopilotOOBE GUI"
        Start-Sleep -Seconds 2
        & "$($MyInvocation.MyCommand.Module.ModuleBase)\Project\MainWindow.ps1"
        #=======================================================================
    }
}