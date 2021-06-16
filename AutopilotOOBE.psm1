function Start-AutopilotOOBE {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$CustomProfile,
        [switch]$Demo,

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
            'MDMDiag',
            'MDMDiagAutopilot',
            'MDMDiagAutopilotTPM',
            'UpdateMyDellBios'
        )]
        [string]$Run = 'PowerShell',
        [string]$Docs,
        [string]$Title = 'Autopilot Manual Enrollment',
        [switch]$Test
    )
    #=======================================================================
    #   Transcript
    #=======================================================================
    $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-AutopilotOOBE.log"
    Start-Transcript -Path (Join-Path "$env:SystemRoot\Temp" $Transcript) -ErrorAction Ignore
    #=======================================================================
    #   Profile OSDeploy
    #=======================================================================
    if ($CustomProfile -in 'OSD','OSDeploy','OSDeploy.com') {
        $Title = 'OSDeploy Autopilot Enrollment'
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
    #   Profile SeguraOSD
    #=======================================================================
    if ($CustomProfile -match 'SeguraOSD') {
        $Title = 'SeguraOSD Autopilot Enrollment'
        $GroupTag = 'Twitter'
        $AssignedComputerName = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $PostAction = 'Restart'
        $Assign = $true
        $Run = 'WindowsSettings'
        $Docs = 'https://twitter.com/SeguraOSD'
        $Hidden = 'AddToGroup','AssignedUser'
    }
    #=======================================================================
    #   Profile Baker Hughes
    #=======================================================================
    if ($CustomProfile -eq 'BH') {
        $Title = 'Baker Hughes Autopilot Enrollment'
        $Assign = $true
        $Hidden = 'AddToGroup','AssignedComputerName','AssignedUser'
        $GroupTag = 'Enterprise'
        $GroupTagOptions = 'Development','Enterprise'
        $Run = 'NetworkingWireless'
        $Test = $true

        if (-NOT (Get-Module -Name OSD -ListAvailable)) {
            Install-Module OSD -Force
            Import-Module OSD -Force
        }
        if ((Get-MyComputerManufacturer -Brief) -eq 'Dell') {
            Update-MyDellBios
            Start-Sleep -Seconds 2
        }
    }
    #=======================================================================
    #   Profile HalfMan
    #=======================================================================
    if ($CustomProfile -eq 'HalfMan') {
        $Title = 'Autopilot Enrollment'
        $Hidden = 'GroupTag'
        $AddToGroup = 'Azr_crp_ent_modern_workplace_devices'
        $AddToGroupOptions = 'Azr_crp_ent_modern_workplace_devices'
    }
    #=======================================================================
    #   Set Global Variable
    #=======================================================================
    $Global:AutopilotOOBE = @{
        AddToGroup = $AddToGroup
        AddToGroupOptions = $AddToGroupOptions
        Assign = $Assign
        AssignedUser = $AssignedUser
        AssignedUserExample = $AssignedUserExample
        AssignedComputerName = $AssignedComputerName
        AssignedComputerNameExample = $AssignedComputerNameExample
        Disabled = $Disabled
        Demo = $Demo
        GroupTag = $GroupTag
        GroupTagOptions = $GroupTagOptions
        Hidden = $Hidden
        PostAction = $PostAction
        Run = $Run
        Docs = $Docs
        Title = $Title
    }
    #=======================================================================
    #   Test
    #   https://docs.microsoft.com/en-us/mem/autopilot/networking-requirements
    #=======================================================================
    if ($Test) {
        Write-Host -ForegroundColor Cyan "Testing Windows Autopilot networking requirements"
        Write-Host -ForegroundColor Cyan "https://docs.microsoft.com/en-us/mem/autopilot/networking-requirements"
        #=======================================================================
        #   PowerShell Gallery
        #=======================================================================
        $TestPSGallery = @(
            'powershellgallery.com'
        )
        Write-Host ""
        Write-Host -ForegroundColor Cyan "PowerShell Gallery"
        foreach ($Item in $TestPSGallery){
            try {
                if (Test-NetConnection -ComputerName $Item -Port 443 -InformationLevel Quiet -ErrorAction Stop) {
                    Write-Host -ForegroundColor Green $Item
                }
                else {
                }
            }
            catch {
                Write-Host -ForegroundColor Red $Item
            }
        }
        #=======================================================================
        #   Windows Autopilot Deployment Service
        #=======================================================================
        $TestWADS = @(
            'cs.dds.microsoft.com'
            'login.live.com'
            'ztd.dds.microsoft.com'
        )
        Write-Host ""
        Write-Host -ForegroundColor Cyan "Windows Autopilot Deployment Service"
        foreach ($Item in $TestWADS){
            try {
                if (Test-NetConnection -ComputerName $Item -Port 443 -InformationLevel Quiet -ErrorAction Stop) {
                    Write-Host -ForegroundColor Green $Item
                }
                else {
                }
            }
            catch {
                Write-Host -ForegroundColor Red $Item
            }
        }
        #=======================================================================
        #   Windows Activation
        #=======================================================================
        $TestWA = @(
            'activation.sls.microsoft.com'
            'activation-v2.sls.microsoft.com'
            'crl.microsoft.com'
            'displaycatalog.mp.microsoft.com'
            'displaycatalog.md.mp.microsoft.com'
            'go.microsoft.com'
            'licensing.mp.microsoft.com'
            'licensing.md.mp.microsoft.com'
            'purchase.mp.microsoft.com'
            'validation.sls.microsoft.com'
            'validation-v2.sls.microsoft.com'
        )
        Write-Host ""
        Write-Host -ForegroundColor Cyan "Windows Activation"
        foreach ($Item in $TestWA){
            try {
                if (Test-NetConnection -ComputerName $Item -Port 443 -InformationLevel Quiet -ErrorAction Stop) {
                    Write-Host -ForegroundColor Green $Item
                }
                else {
                }
            }
            catch {
                Write-Host -ForegroundColor Red $Item
            }
        }
        #=======================================================================
        #   Windows Update
        #=======================================================================
        $TestWU = @(
            'prod.do.dsp.mp.microsoft.com'
            'emdl.ws.microsoft.com'
            'delivery.mp.microsoft.com'
            'dl.delivery.mp.microsoft.com'
            'tsfe.trafficshaping.dsp.mp.microsoft.com'
            'update.microsoft.com'
            'windowsupdate.com'
        )
<#         Write-Host ""
        Write-Host -ForegroundColor Cyan "Windows Update"
        foreach ($Item in $TestWU){
            try {
                if (Test-NetConnection -ComputerName $Item -Port 443 -InformationLevel Quiet -ErrorAction Stop) {
                    Write-Host -ForegroundColor Green $Item
                }
                else {
                }
            }
            catch {
                Write-Host -ForegroundColor Red $Item
            }
        } #>
        #=======================================================================
        #   Autopilot self-Deploying mode and Autopilot pre-provisioning
        #=======================================================================
        $TestTPM = @(
            'ekcert.spserv.microsoft.com'
            'ekop.intel.com'
            'ftpm.amd.com'
        )
        Write-Host ""
        Write-Host -ForegroundColor Cyan "Autopilot self-Deploying mode and Autopilot pre-provisioning"
        foreach ($Item in $TestTPM){
            try {
                if (Test-NetConnection -ComputerName $Item -Port 443 -InformationLevel Quiet -ErrorAction Stop) {
                    Write-Host -ForegroundColor Green $Item
                }
                else {
                }
            }
            catch {
                Write-Host -ForegroundColor Red $Item
            }
        }
        #=======================================================================
        #   Windows Autopilot Deployment Service
        #   https://raw.githubusercontent.com/Mauvlans/AutoPilot/master/AutopilotNetworkValidation.ps1
        #=======================================================================
        $hash = @(
            'a.manage.microsoft.com'
            'account.azureedge.net'
            'account.live.com'
            'enrollment.manage.microsoft.com'
            'EnterpriseEnrollment.manage.microsoft.com'
            'EnterpriseEnrollment-s.manage.microsoft.com'
            'enterpriseregistration.windows.net'
            'fef.msua06.manage.microsoft.com'
            'i.manage.microsoft.com'
            'login.microsoftonline.com'
            'm.fei.msua01.manage.microsoft.com'
            'm.manage.microsoft.com'
            'manage.microsoft.com'
            'msftconnecttest.com'
            'portal.fei.msua01.manage.microsoft.com'
            'portal.manage.microsoft.com'
            'r.manage.microsoft.com'
            'secure.aadcdn.microsoftonline-p.com'
            'signup.live.com'
            'sts.manage.microsoft.com'
        )

        
    }
    #=======================================================================
    #   Launch
    #=======================================================================
    & "$($MyInvocation.MyCommand.Module.ModuleBase)\Forms\Join-AutopilotOOBE.ps1"
}
#=======================================================================
#   Create Alias
#=======================================================================
New-Alias -Name AutopilotOOBE -Value Start-AutopilotOOBE -Force -ErrorAction SilentlyContinue
Export-ModuleMember -Function Start-AutopilotOOBE -Alias AutopilotOOBE